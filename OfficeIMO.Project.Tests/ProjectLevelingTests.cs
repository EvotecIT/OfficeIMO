namespace OfficeIMO.Project.Tests;

public sealed class ProjectLevelingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SplitProposalPreservesDelayedAndCompletedAssignments(bool completed) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Monday;
        var first = document.Resources.AddWork("First"); first.StandardRate = 100;
        var second = document.Resources.AddWork("Second"); second.StandardRate = 100;
        var task = document.Tasks.Add("Delivery"); task.Duration = ProjectDuration.WorkingDays(3);
        var a = document.Assignments.Add(task, first, ProjectUnits.Fraction(1)); a.Work = ProjectWork.Hours(24);
        var b = document.Assignments.Add(task, second, ProjectUnits.Fraction(1)); b.Work = ProjectWork.Hours(16); b.DelayMinutes = 480;
        if (completed) {
            b.DelayMinutes = 0; b.Work = ProjectWork.Hours(4); b.ActualWork = ProjectWork.Hours(4); b.RemainingWork = ProjectWork.Hours(0);
            b.ActualStart = Monday; b.ActualFinish = Monday.AddHours(4); b.Stop = b.ActualFinish;
            task.ActualDuration = ProjectDuration.WorkingHours(2);
        }
        var locked = document.Tasks.Add("Reserved"); locked.Duration = ProjectDuration.WorkingDays(1); locked.Priority = 1000;
        locked.ConstraintType = ProjectConstraintType.MustStartOn; locked.ConstraintDate = Monday.AddDays(1);
        document.Assignments.Add(locked, completed ? first : second, ProjectUnits.Fraction(1));
        var result = document.CalculateLeveling(new ProjectLevelingOptions { AllowSplitting = true }); result.Report.ThrowIfErrors(); Assert.NotEmpty(result.Splits);
        document.ApplyLeveling(result);
        Assert.False(document.Settings.ExternallyEdited);
        if (completed) Assert.Equal(ProjectDuration.WorkingMinutes(120), task.ActualDuration);
        using var copy = document.Clone();
        foreach (var candidate in new[] { document, copy }) {
            var after = candidate.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); after.Report.ThrowIfErrors();
            Assert.Empty(candidate.AnalyzeResourceAllocation(after).Overallocations);
            Assert.Equal(result.Schedule.Assignments.Select(p => (p.AssignmentUid, p.Start, p.Finish, p.Work, p.ActualWork)),
                after.Assignments.Select(p => (p.AssignmentUid, p.Start, p.Finish, p.Work, p.ActualWork)));
        }
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SplitLevelingKeepsThePrefixAndActualsAndPersistsRemainingCurves(bool progress) {
        using var document = Example();
        var locked = document.Tasks[0]; locked.Priority = 1000;
        locked.ConstraintType = ProjectConstraintType.MustStartOn; locked.ConstraintDate = Monday.AddDays(1);
        var task = document.Tasks[1]; task.Duration = ProjectDuration.WorkingDays(3);
        var assignment = document.Assignments.Single(a => a.Task == task);
        assignment.Work = ProjectWork.Hours(24);
        if (progress) {
            assignment.ActualWork = ProjectWork.Hours(4); assignment.RemainingWork = ProjectWork.Hours(20);
            assignment.ActualStart = Monday; assignment.Stop = Monday.AddHours(4);
            task.ActualStart = Monday; task.ActualDuration = ProjectDuration.WorkingHours(4);
            task.RemainingDuration = ProjectDuration.WorkingHours(20);
        }
        var original = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); original.Report.ThrowIfErrors();
        var result = document.CalculateLeveling(new ProjectLevelingOptions { AllowSplitting = true }); result.Report.ThrowIfErrors();
        Assert.NotEmpty(result.Splits); Assert.Empty(result.Moves); Assert.Empty(result.Capacity!.Overallocations);
        var plan = result.Schedule.Assignments.Single(a => a.AssignmentUid == assignment.Uid);
        Assert.Equal(Monday, plan.Start); Assert.Equal(Monday.AddDays(3).AddHours(9), plan.Finish);
        Assert.Equal(1440m, plan.Work.Minutes); Assert.Equal(2400m, plan.Cost!.Value, 6);
        Assert.Equal(original.Assignments.Single(a => a.AssignmentUid == assignment.Uid).Intervals.Where(i => i.Start < Monday.AddDays(1)).Select(i => (i.Start, i.Finish, i.Work, i.IsActual)),
            plan.Intervals.Where(i => i.Start < Monday.AddDays(1)).Select(i => (i.Start, i.Finish, i.Work, i.IsActual)));
        Assert.DoesNotContain(plan.Intervals, i => i.Start < Monday.AddDays(1).AddHours(9) && i.Finish > Monday.AddDays(1));
        document.ApplyLeveling(result);
        using var copy = document.Clone();
        foreach (var candidate in new[] { document, copy }) {
            var recalculated = candidate.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); recalculated.Report.ThrowIfErrors();
            var after = recalculated.Assignments.Single(a => a.AssignmentUid == assignment.Uid);
            Assert.Equal(plan.Intervals.Select(i => (i.Start, i.Finish, Math.Round(i.Work.Minutes, 6), i.IsActual)),
                after.Intervals.Select(i => (i.Start, i.Finish, Math.Round(i.Work.Minutes, 6), i.IsActual)));
            Assert.Empty(candidate.AnalyzeResourceAllocation(recalculated).Overallocations);
        }
    }
    [Fact]
    public void TaskSplitProhibitionFallsBackToDelay() {
        using var document = Example();
        document.Tasks[0].Priority = 1000; document.Tasks[0].ConstraintType = ProjectConstraintType.MustStartOn;
        document.Tasks[0].ConstraintDate = Monday.AddDays(1);
        document.Tasks[1].Duration = ProjectDuration.WorkingDays(3); document.Tasks[1].LevelingCanSplit = false;
        var result = document.CalculateLeveling(new ProjectLevelingOptions { AllowSplitting = true }); result.Report.ThrowIfErrors();
        Assert.Empty(result.Splits); Assert.Single(result.Moves);
        Assert.Equal(Monday.AddDays(2), result.Schedule.Tasks.Single(t => t.TaskUid == document.Tasks[1].Uid).Start);
    }
    [Fact]
    public void ProducerDelayUsesTenthsOfMinutesAndAnExplicitCalendarKind() {
        using var document = ProjectDocument.Load(ProjectResourceCapacityTests.Fixture("leveling"));
        var task = document.Tasks.GetByUid(2);
        Assert.Equal(ProjectDuration.ElapsedDays(1), task.LevelingDelay);
        task.LevelingDelay = ProjectDuration.WorkingHours(4);
        using var copy = document.Clone(); Assert.Equal(ProjectDuration.WorkingHours(4), copy.Tasks.GetByUid(2).LevelingDelay);
        var result = copy.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); result.Report.ThrowIfErrors();
        Assert.Equal(Monday.AddHours(5), result.Tasks.Single(t => t.TaskUid == 2).Start);
        string xml = "<Project xmlns=\"http://schemas.microsoft.com/project\"><Tasks><Task><UID>1</UID><LevelingDelay>60</LevelingDelay></Task></Tasks></Project>";
        Assert.Throws<InvalidDataException>(() => ProjectDocument.Parse(xml));
    }
    private static readonly DateTime Monday = new DateTime(2026, 10, 5, 8, 0, 0);
    private static ProjectDocument Example() {
        var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Monday;
        var resource = document.Resources.AddWork("Engineer"); resource.StandardRate = 100;
        for (int index = 0; index < 2; index++) {
            var task = document.Tasks.Add("Task " + index); task.Duration = ProjectDuration.WorkingDays(1);
            document.Assignments.Add(task, resource, ProjectUnits.Fraction(1));
        }
        return document;
    }
    [Fact]
    public void OrdinaryCalculationReportsOverloadsAndExplicitLevelingUsesStableUidOrder() {
        using var document = Example(); long revision = document.Revision;
        var schedule = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.NotEmpty(document.AnalyzeResourceAllocation(schedule).Overallocations);
        var result = document.CalculateLeveling(); result.Report.ThrowIfErrors();
        Assert.Empty(result.Capacity!.Overallocations); Assert.Equal(revision, document.Revision);
        Assert.Equal(Monday, result.Schedule.Tasks[0].Start); Assert.Equal(Monday.AddDays(1), result.Schedule.Tasks[1].Start);
        Assert.Single(result.Moves); Assert.Equal(result.Schedule.Tasks[1].TaskUid, result.Moves[0].TaskUid);
        var again = document.CalculateLeveling();
        Assert.Equal(result.Schedule.Tasks.Select(t => (t.Start, t.Finish)), again.Schedule.Tasks.Select(t => (t.Start, t.Finish)));
        document.ApplyLeveling(result);
        Assert.Equal(Monday.AddDays(1), document.Tasks[1].Start);
        Assert.Throws<InvalidOperationException>(() => document.ApplyLeveling(result));
        var after = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); after.Report.ThrowIfErrors();
        Assert.Empty(document.AnalyzeResourceAllocation(after).Overallocations);
        Assert.Equal(Monday.AddDays(1), after.Tasks[1].Start);
        using var copy = document.Clone();
        var reopened = copy.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); reopened.Report.ThrowIfErrors();
        Assert.Empty(copy.AnalyzeResourceAllocation(reopened).Overallocations);
        Assert.Equal(Monday.AddDays(1), reopened.Tasks[1].Start);
    }
    [Fact]
    public void PriorityAndSuccessorBoundsPropagateWithoutChangingWork() {
        using var document = Example(); document.Tasks[1].Priority = 900;
        var next = document.Tasks.Add("Successor"); next.Duration = ProjectDuration.WorkingDays(1);
        document.Dependencies.Add(document.Tasks[0], next);
        var result = document.CalculateLeveling(); result.Report.ThrowIfErrors();
        Assert.Equal(Monday, result.Schedule.Tasks[1].Start); Assert.Equal(Monday.AddDays(1), result.Schedule.Tasks[0].Start);
        Assert.Equal(Monday.AddDays(2), result.Schedule.Tasks[2].Start);
        Assert.All(result.Schedule.Assignments, a => Assert.Equal(480m, a.Work.Minutes));
    }
    [Theory]
    [InlineData(0)]
    [InlineData(4)]
    [InlineData(24)]
    public void ApplyingDependentMovesCountsPredecessorMovementOnlyOnce(int previousDelayHours) {
        using var document = Example();
        document.Tasks[0].Duration = ProjectDuration.WorkingDays(2);
        var successor = document.Tasks.Add("C"); successor.Duration = ProjectDuration.WorkingDays(1);
        if (previousDelayHours > 0) successor.LevelingDelay = ProjectDuration.WorkingHours(previousDelayHours);
        var locked = document.Tasks.Add("D"); locked.Duration = ProjectDuration.WorkingDays(1);
        locked.ConstraintType = ProjectConstraintType.MustStartOn; locked.ConstraintDate = Monday.AddDays(3); locked.Priority = 1000;
        document.Assignments.Add(successor, document.Resources[0], ProjectUnits.Fraction(1));
        document.Assignments.Add(locked, document.Resources[0], ProjectUnits.Fraction(1));
        document.Dependencies.Add(document.Tasks[1], successor);
        var result = document.CalculateLeveling(); result.Report.ThrowIfErrors();
        Assert.Equal(Monday.AddDays(2), result.Schedule.Tasks.Single(t => t.TaskUid == document.Tasks[1].Uid).Start);
        if (previousDelayHours == 0) Assert.Equal(Monday.AddDays(4), result.Schedule.Tasks.Single(t => t.TaskUid == successor.Uid).Start);
        document.ApplyLeveling(result);
        using var copy = document.Clone();
        foreach (var candidate in new[] { document, copy }) {
            var after = candidate.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); after.Report.ThrowIfErrors();
            Assert.Equal(result.Schedule.Tasks.Select(t => (t.Start, t.Finish)), after.Tasks.Select(t => (t.Start, t.Finish)));
            Assert.Empty(candidate.AnalyzeResourceAllocation(after).Overallocations);
        }
    }
    [Fact]
    public void ProgressLevelingMovesOnlyRemainingWorkAndSurvivesApply() {
        using var document = Example(); var task = document.Tasks[1]; var assignment = document.Assignments.Single(a => a.Task == task);
        assignment.Work = ProjectWork.Hours(8); assignment.ActualWork = ProjectWork.Hours(4); assignment.RemainingWork = ProjectWork.Hours(4);
        assignment.ActualStart = Monday; assignment.Stop = Monday.AddHours(4); task.ActualStart = Monday;
        task.ActualDuration = ProjectDuration.WorkingHours(4); task.RemainingDuration = ProjectDuration.WorkingHours(4);
        document.Tasks[0].ConstraintType = ProjectConstraintType.StartNoEarlierThan; document.Tasks[0].ConstraintDate = Monday.AddHours(5);
        document.Tasks[0].Priority = 1000;
        var result = document.CalculateLeveling(); result.Report.ThrowIfErrors();
        Assert.Empty(result.Capacity!.Overallocations);
        var plan = result.Schedule.Assignments.Single(a => a.AssignmentUid == assignment.Uid);
        Assert.Equal(Monday, plan.Intervals.Single(i => i.IsActual).Start); Assert.Equal(240m, plan.ActualWork.Minutes);
        document.ApplyLeveling(result);
        var after = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); after.Report.ThrowIfErrors();
        Assert.Empty(document.AnalyzeResourceAllocation(after).Overallocations);
        Assert.Equal(plan.Finish, after.Assignments.Single(a => a.AssignmentUid == assignment.Uid).Finish);
    }
    [Fact]
    public void LockedPrioritiesAndExhaustedLimitsRejectApplyWithoutMutation() {
        using var document = Example(); document.Tasks[0].Priority = 1000; document.Tasks[1].Priority = 1000;
        long revision = document.Revision;
        var result = document.CalculateLeveling(); Assert.True(result.Report.HasErrors);
        Assert.Throws<InvalidDataException>(() => document.ApplyLeveling(result)); Assert.Equal(revision, document.Revision);
        document.Tasks[0].Priority = 500; document.Tasks[1].Priority = 500;
        Assert.True(document.CalculateLeveling(new ProjectLevelingOptions { MaxIterations = 1 }).Report.HasErrors);
        Assert.True(document.CalculateLeveling(new ProjectLevelingOptions { MaxDelayDays = 0 }).Report.HasErrors);
        Assert.True(document.CalculateLeveling(new ProjectLevelingOptions { WithinAvailableSlack = true }).Report.HasErrors);
        Assert.Throws<OperationCanceledException>(() => document.CalculateLeveling(cancellationToken: new CancellationToken(true)));
    }
}
