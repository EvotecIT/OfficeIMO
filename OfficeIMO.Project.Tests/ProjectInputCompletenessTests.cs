namespace OfficeIMO.Project.Tests;

public sealed class ProjectInputCompletenessTests {
    private static readonly DateTime Monday = new(2026, 10, 5, 8, 0, 0);
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CompleteOrExplicitlyRedistributedFixedWorkKeepsItsDeclaredEffort(bool redistribute) {
        using var document = Create(); var task = document.Tasks.Add("Fixed effort"); task.Type = ProjectTaskType.FixedWork;
        task.Duration = ProjectDuration.WorkingDays(1); task.Work = ProjectWork.Hours(16);
        var resource = document.Resources.AddWork("Engineer"); resource.StandardRate = 100m;
        var assignment = document.Assignments.Add(task, resource); if (!redistribute) assignment.Work = ProjectWork.Hours(16);
        document.Recalculate(new ProjectScheduleOptions { CalculateAssignments = true, RedistributeEffortDrivenWork = redistribute });
        Assert.Equal(960m, task.Work.Value.Minutes); Assert.Equal(Monday.AddDays(1).AddHours(9), task.Finish);
    }
    [Fact]
    public void TaskActualWorkCannotBeDiscardedByIncompleteAssignments() {
        using var document = Create(); var task = document.Tasks.Add("Started"); task.Duration = ProjectDuration.WorkingDays(1);
        task.ActualWork = ProjectWork.Hours(4); document.Assignments.Add(task, document.Resources.AddWork("Engineer"));
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.True(result.Report.HasErrors); Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result)); Assert.Equal(240m, task.ActualWork.Value.Minutes);
    }
    [Fact]
    public void ConflictingAssignmentWorkComponentsAreRejected() {
        using var document = Create(); var task = document.Tasks.Add("Started"); task.Duration = ProjectDuration.WorkingDays(1);
        var assignment = document.Assignments.Add(task, document.Resources.AddWork("Engineer"));
        assignment.Work = ProjectWork.Hours(8); assignment.ActualWork = ProjectWork.Hours(4); assignment.RemainingWork = ProjectWork.Hours(8);
        assignment.ActualStart = Monday; assignment.Stop = Monday.AddHours(4);
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.True(result.Report.HasErrors); Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result)); Assert.Equal(480m, assignment.Work.Value.Minutes);
    }
    [Theory]
    [InlineData("duration", false)]
    [InlineData("work", false)]
    [InlineData("assignment", false)]
    [InlineData("duration", true)]
    [InlineData("work", true)]
    [InlineData("assignment", true)]
    public void PercentageOnlyProgressCannotBeOverwritten(string owner, bool backward) {
        using var document = Create(); var task = document.Tasks.Add("Started"); task.Duration = ProjectDuration.WorkingDays(1);
        var assignment = document.Assignments.Add(task, document.Resources.AddWork("Engineer"));
        if (owner == "duration") task.PercentComplete = 50;
        else if (owner == "work") task.PercentWorkComplete = 50;
        else assignment.PercentWorkComplete = 50;
        if (backward) { document.Settings.ScheduleFromStart = false; document.Settings.FinishDate = Monday.AddDays(2); }
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.True(result.Report.HasErrors); Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result));
        Assert.Equal(50, owner == "duration" ? task.PercentComplete : owner == "work" ? task.PercentWorkComplete : assignment.PercentWorkComplete);
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FixedWorkCannotBeReplacedByDurationBasedAssignmentDefaults(bool explicitWrongAssignmentWork) {
        using var document = Create(); var task = document.Tasks.Add("Fixed effort"); task.Type = ProjectTaskType.FixedWork;
        task.Duration = ProjectDuration.WorkingDays(1); task.Work = ProjectWork.Hours(16);
        var assignment = document.Assignments.Add(task, document.Resources.AddWork("Engineer"));
        if (explicitWrongAssignmentWork) assignment.Work = ProjectWork.Hours(8);
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.True(result.Report.HasErrors); Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result)); Assert.Equal(960m, task.Work.Value.Minutes);
    }
    [Fact]
    public void RemovingTheLastAssignmentClearsCurrentResourceTotals() {
        using var document = Create(); var task = document.Tasks.Add("Delivery"); task.Duration = ProjectDuration.WorkingDays(1);
        var resource = document.Resources.AddWork("Engineer"); resource.StandardRate = 100m;
        var assignment = document.Assignments.Add(task, resource);
        document.Recalculate(new ProjectScheduleOptions { CalculateAssignments = true }); Assert.Equal(800m, resource.Cost);
        document.Assignments.Remove(assignment);
        document.Recalculate(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.Equal(0m, resource.Work!.Value.Minutes); Assert.Equal(0m, resource.ActualWork!.Value.Minutes);
        Assert.Equal(0m, resource.RemainingWork!.Value.Minutes); Assert.Equal(0m, resource.Cost); Assert.Equal(0m, resource.ActualCost);
        Assert.False(document.AreWorkCostTotalsStale);
        using var copy = document.Clone(); Assert.Equal(0m, copy.Resources.GetByUid(resource.Uid).Cost);
    }
    private static ProjectDocument Create() { var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Monday; return document; }
}
