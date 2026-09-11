namespace OfficeIMO.Project.Tests;

public sealed class ProjectInputProjectionTests {
    private static readonly DateTime Monday = new(2026, 10, 5, 8, 0, 0);
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UnassignedWorkComponentsArePreservedOrRejectedWhenConflicting(bool conflicting) {
        using var document = Create(); var task = document.Tasks.Add("Unassigned"); task.Duration = ProjectDuration.WorkingDays(1);
        task.RemainingWork = ProjectWork.Hours(8); if (conflicting) task.Work = ProjectWork.Hours(4);
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        if (conflicting) { Assert.True(result.Report.HasErrors); Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result)); }
        else { document.ApplySchedule(result); Assert.Equal(480m, task.Work!.Value.Minutes); }
        Assert.Equal(480m, task.RemainingWork!.Value.Minutes);
    }
    [Fact]
    public void MissingAssignmentRemainingWorkDoesNotRepeatCompletedEffort() {
        using var document = Create(); var task = document.Tasks.Add("Started"); task.Type = ProjectTaskType.FixedUnits;
        task.Duration = ProjectDuration.WorkingHours(8); task.ActualDuration = ProjectDuration.WorkingHours(4); task.RemainingDuration = ProjectDuration.WorkingHours(4);
        task.Work = ProjectWork.Hours(8); task.ActualWork = ProjectWork.Hours(4); task.RemainingWork = ProjectWork.Hours(4);
        var assignment = document.Assignments.Add(task, document.Resources.AddWork("Engineer")); assignment.ActualWork = ProjectWork.Hours(4);
        assignment.ActualStart = Monday; assignment.Stop = Monday.AddHours(4);
        document.Recalculate(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.Equal(480m, task.Work!.Value.Minutes); Assert.Equal(240m, assignment.RemainingWork!.Value.Minutes); Assert.Equal(Monday.AddHours(9), task.Finish);
        document.Recalculate(new ProjectScheduleOptions { CalculateAssignments = true }); Assert.Equal(480m, task.Work.Value.Minutes);
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OvertimeCurvesRequireCompleteActualWorkIntervals(bool overtimeOnly) {
        using var document = Create(); var task = document.Tasks.Add("Actual overtime"); task.Duration = ProjectDuration.WorkingDays(1);
        var assignment = document.Assignments.Add(task, document.Resources.AddWork("Engineer")); assignment.Work = ProjectWork.Hours(8);
        if (!overtimeOnly) { assignment.ActualWork = ProjectWork.Hours(4); assignment.ActualStart = Monday; assignment.Stop = Monday.AddHours(4); }
        var curve = assignment.TimephasedData.Add(); curve.Uid = assignment.Uid; curve.Type = 3; curve.Start = Monday; curve.Finish = Monday.AddHours(1); curve.Value = "PT1H";
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.True(result.Report.HasErrors); Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result)); Assert.Null(assignment.ActualOvertimeWork);
    }
    [Fact]
    public void BackwardPercentageLagUsesCalculatedFixedWorkDuration() {
        using var document = Create(); document.Settings.ScheduleFromStart = false; document.Settings.FinishDate = Monday.AddDays(4).AddHours(9);
        var first = document.Tasks.Add("Two day effort"); first.Type = ProjectTaskType.FixedWork; first.Duration = ProjectDuration.WorkingDays(1); first.Work = ProjectWork.Hours(16);
        var last = document.Tasks.Add("Successor"); last.Duration = ProjectDuration.WorkingDays(1);
        var resource = document.Resources.AddWork("Engineer"); document.Assignments.Add(first, resource).Work = ProjectWork.Hours(16); document.Assignments.Add(last, resource).Work = ProjectWork.Hours(8);
        document.Dependencies.Add(first, last).LagPercent = 50;
        document.Recalculate(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.Equal(Monday.AddDays(1), first.Start); Assert.Equal(Monday.AddDays(2).AddHours(9), first.Finish); Assert.Equal(Monday.AddDays(4), last.Start);
    }
    [Theory]
    [InlineData(ProjectViewKind.TaskUsage, ProjectViewTimescale.Day)]
    [InlineData(ProjectViewKind.ResourceUsage, ProjectViewTimescale.Week)]
    [InlineData(ProjectViewKind.ResourceHistogram, ProjectViewTimescale.Month)]
    public void UsageBucketsIncludePointOvertimeAndHonorExclusiveFinish(ProjectViewKind kind, ProjectViewTimescale timescale) {
        using var document = Create(); var task = document.Tasks.Add("Overtime point"); task.Duration = ProjectDuration.WorkingDays(1);
        var assignment = document.Assignments.Add(task, document.Resources.AddWork("Engineer")); assignment.Work = ProjectWork.Hours(1); assignment.OvertimeWork = ProjectWork.Hours(1);
        var schedule = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); schedule.Report.ThrowIfErrors();
        var view = document.CreateView(schedule, new ProjectViewOptions { Kind = kind, Timescale = timescale });
        Assert.Equal(1m, Assert.Single(view.Rows).BucketWorkHours.Sum());
        var clipped = document.CreateView(schedule, new ProjectViewOptions { Kind = kind, Timescale = timescale, Start = Monday.AddDays(-1), Finish = Monday });
        Assert.Equal(0m, Assert.Single(clipped.Rows).BucketWorkHours.Sum());
    }
    private static ProjectDocument Create() { var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Monday; return document; }
}
