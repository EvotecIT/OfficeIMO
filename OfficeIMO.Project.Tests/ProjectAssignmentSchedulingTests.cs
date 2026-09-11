namespace OfficeIMO.Project.Tests;

public sealed class ProjectAssignmentSchedulingTests {
    [Theory]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(false, false)]
    public void DelayedAssignmentsMeetBackwardAndFinishConstraintBounds(bool backward, bool finishConstraint) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        var monday = new DateTime(2026, 10, 5, 8, 0, 0); var horizon = monday.AddDays(1).AddHours(9);
        document.Settings.StartDate = monday; document.Settings.FinishDate = horizon; document.Settings.ScheduleFromStart = !backward;
        var task = document.Tasks.Add("Delayed"); task.Duration = ProjectDuration.WorkingDays(1);
        if (finishConstraint) { task.ConstraintType = ProjectConstraintType.MustFinishOn; task.ConstraintDate = horizon; }
        else if (!backward) {
            task.ConstraintType = ProjectConstraintType.AsLateAsPossible;
            var longTask = document.Tasks.Add("Horizon"); longTask.Duration = ProjectDuration.WorkingDays(2);
        }
        var resource = document.Resources.AddWork("Engineer"); resource.StandardRate = 100;
        var assignment = document.Assignments.Add(task, resource, ProjectUnits.Fraction(1)); assignment.DelayMinutes = 60m;
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); result.Report.ThrowIfErrors();
        var plan = result.Tasks.Single(t => t.TaskUid == task.Uid); var assigned = Assert.Single(result.Assignments);
        Assert.Equal(horizon, plan.Finish); Assert.Equal(horizon, assigned.Finish);
        Assert.True(assigned.Start >= plan.Start); Assert.Equal(480m, assigned.Work.Minutes);
        document.ApplySchedule(result); Assert.False(document.AreWorkCostTotalsStale);
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UnassignedAndCostOnlyElapsedTasksKeepContinuousTime(bool costResource) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        var monday = new DateTime(2026, 10, 5, 8, 0, 0); document.Settings.StartDate = monday;
        var task = document.Tasks.Add("Elapsed"); task.Duration = ProjectDuration.ElapsedDays(1);
        if (costResource) document.Assignments.Add(task, document.Resources.AddCost("Travel")).Cost = 100m;
        var result = document.Recalculate(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.Equal(monday.AddDays(1), task.Finish); Assert.True(task.RemainingDuration!.Value.IsElapsed);
        if (costResource) Assert.Equal(task.Finish, result.Assignments.Single().Finish);
        Assert.Equal(task.Finish, document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }).Tasks.Single().Finish);
    }
    [Fact]
    public void AddingAnEffortDrivenAssignmentRedistributesOnlyRemainingWork() {
        using var before = ProjectDocument.Load(ProjectResourceCapacityTests.Fixture("effort-before"));
        using var expected = ProjectDocument.Load(ProjectResourceCapacityTests.Fixture("effort"));
        var reviewer = before.Resources.Single(r => r.Name == "Reviewer");
        foreach (var task in before.AllTasks.Where(t => t.Uid != 0 && !t.IsSummary).ToArray())
            before.Assignments.Add(task, reviewer, ProjectUnits.Fraction(1m));
        var result = before.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true, RedistributeEffortDrivenWork = true });
        Assert.False(result.Report.HasErrors, string.Join("\n", result.Report.Diagnostics.Select(d => d.Message)));
        foreach (var plan in result.Tasks.Where(t => !t.IsSummary)) {
            var task = expected.Tasks.GetByUid(plan.TaskUid);
            Assert.Equal(task.Start, plan.Start); Assert.Equal(task.Finish, plan.Finish);
            Assert.Equal(task.Work!.Value.Minutes, plan.Calculation!.Work.Minutes);
            foreach (var assignment in result.Assignments.Where(a => a.TaskUid == plan.TaskUid)) {
                var source = expected.Assignments.Single(a => a.Task?.Uid == plan.TaskUid && a.Resource?.Uid == assignment.ResourceUid);
                Assert.Equal(source.Work!.Value.Minutes, assignment.Work.Minutes);
            }
        }
    }
    [Fact]
    public void GeneratedNamedContoursMatchProducerWorkingBoundaries() {
        using var document = ProjectDocument.Load(ProjectResourceCapacityTests.Fixture("contours"));
        var expected = document.Assignments.ToDictionary(a => a.Uid, a => a.Finish);
        foreach (var assignment in document.Assignments.Where(a => a.WorkContour != ProjectWorkContour.Turtle))
            foreach (var value in assignment.TimephasedData.Where(v => v.Type == 1).ToArray()) assignment.TimephasedData.Remove(value);
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.False(result.Report.HasErrors, string.Join("\n", result.Report.Diagnostics.Select(d => d.Message)));
        foreach (var assignment in result.Assignments) Assert.Equal(expected[assignment.AssignmentUid], assignment.Finish);
    }
    [Fact]
    public void StatusReschedulingPreservesRecordedActualsAndIsAtomicWhenCanceledOrStale() {
        using var document = ProjectDocument.Load(ProjectResourceCapacityTests.Fixture("progress"));
        document.Settings.StatusDate = new DateTime(2026, 10, 12, 8, 0, 0);
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true, RescheduleRemainingAfterStatusDate = true });
        Assert.False(result.Report.HasErrors, string.Join("\n", result.Report.Diagnostics.Select(d => d.Message)));
        Assert.All(result.Assignments.SelectMany(a => a.Intervals).Where(i => !i.IsActual), i => Assert.True(i.Start >= document.Settings.StatusDate));
        long revision = document.Revision;
        Assert.Throws<OperationCanceledException>(() => document.ApplySchedule(result, new CancellationToken(true)));
        Assert.Equal(revision, document.Revision);
        document.ApplySchedule(result);
        Assert.Throws<InvalidOperationException>(() => document.ApplySchedule(result));
        var limited = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true, MaxIntervals = 1 });
        Assert.True(limited.Report.HasErrors);
        Assert.Throws<InvalidDataException>(() => document.ApplySchedule(limited));
    }
    [Theory]
    [InlineData("rates")]
    [InlineData("calendars")]
    [InlineData("contours")]
    [InlineData("effort")]
    [InlineData("progress")]
    [InlineData("leveling")]
    public void ProducerAssignmentSchedulesPreserveDatesEffortAndCosts(string fixture) {
        using var document = ProjectDocument.Load(ProjectResourceCapacityTests.Fixture(fixture));
        long revision = document.Revision;
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.False(result.Report.HasErrors, string.Join("\n", result.Report.Diagnostics.Select(d => d.Code + ": " + d.Message)));
        Assert.Equal(revision, document.Revision);
        foreach (var plan in result.Tasks.Where(t => !t.IsSummary)) {
            var source = document.Tasks.GetByUid(plan.TaskUid);
            Assert.Equal(source.Start, plan.Start); Assert.Equal(source.Finish, plan.Finish);
            var totals = Assert.IsType<ProjectTaskWorkSchedule>(plan.Calculation);
            Assert.InRange(Math.Abs((source.Work?.Minutes ?? 0m) - totals.Work.Minutes), 0m, .001m);
            Assert.InRange(Math.Abs((source.Cost ?? 0m) - (totals.Cost ?? decimal.MaxValue)), 0m, .02m);
            Assert.Equal(source.PercentComplete ?? 0, totals.PercentComplete);
            Assert.Equal(source.PercentWorkComplete ?? 0, totals.PercentWorkComplete);
        }
        foreach (var plan in result.Assignments) {
            var source = document.Assignments.GetByUid(plan.AssignmentUid);
            Assert.True(source.Start == plan.Start, $"{fixture} assignment {plan.AssignmentUid} start: expected {source.Start:O}, actual {plan.Start:O}");
            Assert.True(source.Finish == plan.Finish, $"{fixture} assignment {plan.AssignmentUid} finish: expected {source.Finish:O}, actual {plan.Finish:O}");
            Assert.InRange(Math.Abs((source.Work?.Minutes ?? 0m) - plan.Work.Minutes), 0m, .001m);
            Assert.InRange(Math.Abs((source.ActualWork?.Minutes ?? 0m) - plan.ActualWork.Minutes), 0m, .001m);
            Assert.InRange(Math.Abs((source.Cost ?? 0m) - (plan.Cost ?? decimal.MaxValue)), 0m, .02m);
        }
        var actuals = document.Assignments.ToDictionary(a => a.Uid, a => a.TimephasedData.Where(v => v.Type == 2 || v.Type == 3).ToArray());
        document.ApplySchedule(result);
        foreach (var assignment in document.Assignments)
            Assert.Equal(actuals[assignment.Uid], assignment.TimephasedData.Where(v => v.Type == 2 || v.Type == 3).ToArray());
        var again = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.False(again.Report.HasErrors, string.Join("\n", again.Report.Diagnostics.Select(d => d.Message)));
        foreach (var first in result.Assignments) {
            var second = again.Assignments.Single(a => a.AssignmentUid == first.AssignmentUid);
            Assert.Equal(first.Start, second.Start); Assert.Equal(first.Finish, second.Finish);
            Assert.InRange(Math.Abs(first.Work.Minutes - second.Work.Minutes), 0m, .001m);
            Assert.InRange(Math.Abs((first.Cost ?? 0m) - (second.Cost ?? 0m)), 0m, .02m);
        }
    }
}
