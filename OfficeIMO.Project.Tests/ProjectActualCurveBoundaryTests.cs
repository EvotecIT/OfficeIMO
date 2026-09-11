namespace OfficeIMO.Project.Tests;

public sealed class ProjectActualCurveBoundaryTests {
    private static readonly DateTime Monday = new(2026, 10, 5, 8, 0, 0);

    [Fact]
    public void ConcurrentActualAndRemainingAssignmentsStayStableAfterApplyAndXml() {
        using var document = Create();
        var task = document.Tasks.Add("Concurrent delivery"); task.Type = ProjectTaskType.FixedUnits;
        task.Duration = ProjectDuration.WorkingHours(8);
        var completed = document.Assignments.Add(task, document.Resources.AddWork("Completed engineer"));
        completed.Work = completed.ActualWork = ProjectWork.Hours(4); completed.ActualStart = Monday; completed.ActualFinish = Monday.AddHours(4);
        var remaining = document.Assignments.Add(task, document.Resources.AddWork("Remaining engineer")); remaining.Work = ProjectWork.Hours(8);
        var options = new ProjectScheduleOptions { CalculateAssignments = true };
        var original = document.CalculateSchedule(options); original.Report.ThrowIfErrors();
        Assert.Equal(Monday, original.Assignments.Single(a => a.AssignmentUid == remaining.Uid).Start);
        Assert.Equal(240m, Assert.Single(original.Tasks).Calculation!.RemainingDuration.Value);
        document.ApplySchedule(original);
        using var reopened = document.Clone();
        foreach (var candidate in new[] { document, reopened }) {
            var repeated = candidate.CalculateSchedule(options); repeated.Report.ThrowIfErrors();
            Assert.Equal(original.Assignments.Select(a => (a.AssignmentUid, a.Start, a.Finish, a.Work, a.ActualWork)),
                repeated.Assignments.Select(a => (a.AssignmentUid, a.Start, a.Finish, a.Work, a.ActualWork)));
            Assert.Equal(original.Tasks.Single().Duration, repeated.Tasks.Single().Duration);
            Assert.Equal(240m, repeated.Tasks.Single().Calculation!.RemainingDuration.Value);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ActualCurveTimingSurvivesInferredTaskProgress(bool discontinuous) {
        using var document = Create(); var task = document.Tasks.Add("Parallel progress"); task.Type = ProjectTaskType.FixedUnits;
        task.Duration = ProjectDuration.WorkingHours(discontinuous ? 10 : 8);
        var first = document.Assignments.Add(task, document.Resources.AddWork("First engineer"));
        var second = document.Assignments.Add(task, document.Resources.AddWork("Second engineer"));
        if (discontinuous) {
            first.Work = first.ActualWork = ProjectWork.Hours(4); second.Work = ProjectWork.Hours(8);
            Add(first, 2, Monday, Monday.AddHours(2), "PT2H");
            Add(first, 2, Monday.AddDays(1), Monday.AddDays(1).AddHours(2), "PT2H");
        } else {
            first.Work = ProjectWork.Hours(8); first.ActualWork = ProjectWork.Hours(2);
            Add(first, 2, Monday, Monday.AddHours(2), "PT2H");
            second.Work = second.ActualWork = ProjectWork.Hours(4);
            Add(second, 2, Monday, Monday.AddHours(4), "PT4H");
        }
        var options = new ProjectScheduleOptions { CalculateAssignments = true };
        var original = document.CalculateSchedule(options); original.Report.ThrowIfErrors();
        Assert.Equal(discontinuous ? 360m : 240m, original.Tasks.Single().Calculation!.RemainingDuration.Value);
        document.ApplySchedule(original);
        using var reopened = document.Clone();
        foreach (var candidate in new[] { document, reopened }) {
            var repeated = candidate.CalculateSchedule(options); repeated.Report.ThrowIfErrors();
            Assert.Equal(original.Assignments.Select(a => (a.Start, a.Finish, a.Work, a.ActualWork)),
                repeated.Assignments.Select(a => (a.Start, a.Finish, a.Work, a.ActualWork)));
            Assert.Equal(original.Tasks.Single().Duration, repeated.Tasks.Single().Duration);
            Assert.Equal(original.Tasks.Single().Calculation!.RemainingDuration, repeated.Tasks.Single().Calculation!.RemainingDuration);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RecordedActualDurationRequiresAnActualAnchor(bool assigned) {
        using var document = Create(); var task = document.Tasks.Add("Missing actual start"); task.Type = ProjectTaskType.FixedUnits;
        task.Duration = ProjectDuration.WorkingHours(16); task.ActualDuration = task.RemainingDuration = ProjectDuration.WorkingHours(8);
        if (assigned) document.Assignments.Add(task, document.Resources.AddWork("Engineer")).Work = ProjectWork.Hours(8);
        long revision = document.Revision;
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.True(result.Report.HasErrors); Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result));
        Assert.Equal(revision, document.Revision); Assert.Equal(ProjectDuration.WorkingHours(8), task.RemainingDuration);
        task.ActualStart = Monday;
        var repaired = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); repaired.Report.ThrowIfErrors();
        Assert.Equal(Monday.AddDays(1).AddHours(9), repaired.Tasks.Single().Finish);
        Assert.Equal(480m, repaired.Tasks.Single().Calculation!.RemainingDuration.Value);
        document.ApplySchedule(repaired);
        using var reopened = document.Clone();
        var repeated = reopened.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); repeated.Report.ThrowIfErrors();
        Assert.Equal(repaired.Tasks.Single().Finish, repeated.Tasks.Single().Finish);
        Assert.Equal(repaired.Tasks.Single().Duration, repeated.Tasks.Single().Duration);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AnEarlierResumeCannotOverlapUnassignedRecordedProgress(bool elapsed) {
        using var document = Create(); var task = document.Tasks.Add("Unassigned progress");
        task.Duration = elapsed ? ProjectDuration.ElapsedHours(16) : ProjectDuration.WorkingHours(16);
        task.ActualDuration = task.RemainingDuration = elapsed ? ProjectDuration.ElapsedHours(8) : ProjectDuration.WorkingHours(8);
        task.ActualStart = Monday; task.Resume = Monday.AddHours(2);
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); result.Report.ThrowIfErrors();
        Assert.Equal(elapsed ? Monday.AddHours(16) : Monday.AddDays(1).AddHours(9), result.Tasks.Single().Finish);
        Assert.Equal(480m, result.Tasks.Single().Calculation!.RemainingDuration.Value);
    }

    [Theory]
    [InlineData(3, 4, false)]
    [InlineData(1, 3, true)]
    [InlineData(3, 3, true)]
    public void ActualOvertimeOutsideActualWorkCannotBeSilentlyDropped(int from, int to, bool scalar) {
        using var document = Create();
        var task = document.Tasks.Add("Recorded work"); task.Duration = ProjectDuration.WorkingHours(8);
        var assignment = document.Assignments.Add(task, document.Resources.AddWork("Engineer"));
        assignment.Work = ProjectWork.Hours(8); assignment.ActualWork = ProjectWork.Hours(2);
        if (scalar) assignment.ActualOvertimeWork = ProjectWork.Hours(1);
        Add(assignment, 2, Monday, Monday.AddHours(2), "PT2H");
        Add(assignment, 3, Monday.AddHours(from), Monday.AddHours(to), "PT1H");
        long revision = document.Revision;
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.True(result.Report.HasErrors); Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result));
        Assert.Equal(revision, document.Revision); Assert.Equal(2, assignment.TimephasedData.Count);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ActualOvertimeCoverageAllowsCalendarBreaksButRejectsMissingWorkingTime(bool workingGap) {
        using var document = Create();
        var task = document.Tasks.Add("Recorded work"); task.Duration = ProjectDuration.WorkingHours(8);
        var resource = document.Resources.AddWork("Engineer"); resource.StandardRate = 100; resource.OvertimeRate = 150;
        var assignment = document.Assignments.Add(task, resource);
        assignment.Work = assignment.ActualWork = ProjectWork.Hours(workingGap ? 4 : 6);
        assignment.ActualOvertimeWork = ProjectWork.Hours(1);
        Add(assignment, 2, Monday, Monday.AddHours(workingGap ? 2 : 4), workingGap ? "PT2H" : "PT4H");
        Add(assignment, 2, Monday.AddHours(5), Monday.AddHours(7), "PT2H");
        Add(assignment, 3, Monday, Monday.AddHours(7), "PT1H");
        var options = new ProjectScheduleOptions { CalculateAssignments = true };
        var result = document.CalculateSchedule(options);
        if (workingGap) { Assert.True(result.Report.HasErrors); Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result)); return; }
        result.Report.ThrowIfErrors(); var plan = Assert.Single(result.Assignments);
        Assert.Equal(360m, plan.ActualWork.Minutes); Assert.Equal(60m, plan.Intervals.Sum(i => i.OvertimeWork.Minutes));
        Assert.Equal(650m, decimal.Round(plan.ActualCost!.Value, 2));
        document.ApplySchedule(result);
        using var reopened = document.Clone(); var repeated = reopened.CalculateSchedule(options); repeated.Report.ThrowIfErrors();
        Assert.Equal(plan.ActualWork, Assert.Single(repeated.Assignments).ActualWork);
        Assert.Equal(650m, decimal.Round(Assert.Single(repeated.Assignments).ActualCost!.Value, 2));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void StoredActualDurationRespectsLaterResumeAndResourceCalendar(bool statusDate) {
        using var document = Create();
        var task = document.Tasks.Add("Progress"); task.Duration = ProjectDuration.WorkingHours(16);
        task.ActualStart = Monday; task.ActualDuration = task.RemainingDuration = ProjectDuration.WorkingHours(8);
        var resource = document.Resources.AddWork("Engineer");
        resource.Calendar = document.Calendars.AddStandardWorkingWeek("Afternoon Tuesday");
        resource.Calendar.SetWorkingDay(DayOfWeek.Tuesday, ProjectWorkingTime.Hours(13, 17));
        var assignment = document.Assignments.Add(task, resource); assignment.Work = ProjectWork.Hours(10); assignment.ActualWork = ProjectWork.Hours(2);
        Add(assignment, 2, Monday, Monday.AddHours(2), "PT2H");
        if (statusDate) document.Settings.StatusDate = Monday.AddDays(1).AddHours(6);
        else { task.Stop = Monday.AddHours(9); task.Resume = Monday.AddDays(1).AddHours(1); }
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true, RescheduleRemainingAfterStatusDate = statusDate });
        result.Report.ThrowIfErrors();
        Assert.Equal(Monday.AddDays(1).AddHours(statusDate ? 6 : 5), Assert.Single(result.Assignments).Intervals.Where(i => !i.IsActual).Min(i => i.Start));
        Assert.Equal(480m, Assert.Single(result.Tasks).Calculation!.RemainingDuration.Value);
    }

    [Theory]
    [InlineData(false, ProjectTaskType.FixedDuration)]
    [InlineData(true, ProjectTaskType.FixedDuration)]
    [InlineData(false, ProjectTaskType.FixedUnits)]
    [InlineData(false, ProjectTaskType.FixedWork)]
    public void RemainingWorkStartsAfterTheRecordedTaskActualDuration(bool manual, ProjectTaskType type) {
        using var document = Create();
        var task = document.Tasks.Add("Partly staffed progress"); task.Type = type;
        task.Duration = ProjectDuration.WorkingHours(16); task.ActualStart = Monday;
        task.ActualDuration = task.RemainingDuration = ProjectDuration.WorkingHours(8);
        if (manual) { task.IsManual = true; task.Start = Monday; task.Finish = Monday.AddDays(1).AddHours(9); }
        var resource = document.Resources.AddWork("Engineer"); resource.StandardRate = 100;
        var assignment = document.Assignments.Add(task, resource);
        assignment.Work = ProjectWork.Hours(10); assignment.ActualWork = ProjectWork.Hours(2);
        Add(assignment, 2, Monday, Monday.AddHours(2), "PT2H");
        var options = new ProjectScheduleOptions { CalculateAssignments = true };
        var result = document.CalculateSchedule(options); result.Report.ThrowIfErrors();
        var plan = Assert.Single(result.Assignments);
        Assert.Equal(Monday.AddDays(1), plan.Intervals.Where(i => !i.IsActual).Min(i => i.Start));
        Assert.Equal(480m, Assert.Single(result.Tasks).Calculation!.RemainingDuration.Value);
        Assert.Equal(1000m, plan.Cost);
        document.ApplySchedule(result);
        using var reopened = document.Clone();
        var repeated = reopened.CalculateSchedule(options); repeated.Report.ThrowIfErrors();
        Assert.Equal(Monday.AddDays(1), Assert.Single(repeated.Assignments).Intervals.Where(i => !i.IsActual).Min(i => i.Start));
    }

    private static ProjectDocument Create() {
        var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        document.Settings.StartDate = Monday; return document;
    }
    private static void Add(ProjectAssignment assignment, int type, DateTime start, DateTime finish, string value) {
        var curve = assignment.TimephasedData.Add(); curve.Type = type; curve.Start = start; curve.Finish = finish; curve.Value = value;
    }
}
