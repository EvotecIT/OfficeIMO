namespace OfficeIMO.Project.Tests;

public sealed class ProjectStopLevelingFormulaTests {
    private static readonly DateTime Monday = new(2026, 10, 5, 8, 0, 0);

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void TaskStopAnchorsRemainingDurationAndAssignmentWork(bool assigned, bool resumed) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Monday;
        var task = document.Tasks.Add("Stopped"); task.Duration = ProjectDuration.WorkingHours(8); task.ActualDuration = ProjectDuration.WorkingHours(1);
        task.ActualStart = Monday; task.Stop = Monday.AddHours(3); if (resumed) task.Resume = Monday.AddHours(5);
        if (assigned) {
            var assignment = document.Assignments.Add(task, document.Resources.AddWork("Engineer"));
            assignment.Work = ProjectWork.Hours(8); assignment.ActualWork = ProjectWork.Hours(1); assignment.ActualStart = Monday;
        }
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); result.Report.ThrowIfErrors();
        Assert.Equal(Monday.AddDays(1).AddHours(resumed ? 3 : 2), Assert.Single(result.Tasks).Finish);
        Assert.All(result.Assignments.SelectMany(a => a.Intervals).Where(i => !i.IsActual), i => Assert.True(i.Start >= (task.Resume ?? task.Stop)));
        document.ApplySchedule(result);
        using var copy = document.Clone(); Assert.Equal(task.Stop, copy.Tasks[0].Stop);
        var repeated = copy.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); repeated.Report.ThrowIfErrors();
        Assert.Equal(result.Tasks.Single().Finish, repeated.Tasks.Single().Finish);
    }

    [Theory]
    [InlineData("Remaining Cost", false)]
    [InlineData("205520917", false)]
    [InlineData("Remaining Cost", true)]
    [InlineData("205520917", true)]
    public void ResourceFormulasReadTheIndependentRemainingCost(string field, bool missing) {
        using var document = ProjectDocument.Create(); var resource = document.Resources.AddWork("Engineer");
        resource.Cost = 100; resource.ActualCost = 25; if (!missing) resource.RemainingCost = 10;
        var definition = document.CustomFields.Add(); definition.FieldId = "205521008"; definition.Formula = "[" + field + "]";
        var result = document.CalculateCustomFields();
        if (missing) { Assert.True(result.Report.HasErrors); Assert.Throws<InvalidDataException>(() => document.ApplyCustomFields(result)); return; }
        result.Report.ThrowIfErrors(); Assert.Equal("10", Assert.Single(result.Values).Value);
        document.ApplyCustomFields(result); using var copy = document.Clone();
        Assert.Equal("10", Assert.Single(copy.Resources[0].CustomFields).Value); Assert.Equal(10m, copy.Resources[0].RemainingCost);
    }

    [Theory]
    [InlineData(1, false)]
    [InlineData(3, true)]
    public void LevelingDelayLimitIncludesWeekendCalendarSnapping(int maxDays, bool accepted) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        var friday = Monday.AddDays(4).AddHours(8); document.Settings.StartDate = friday;
        var resource = document.Resources.AddWork("Engineer"); var locked = document.Tasks.Add("Locked"); locked.Duration = ProjectDuration.WorkingHours(1); locked.Priority = 1000;
        var moved = document.Tasks.Add("Movable"); moved.Duration = ProjectDuration.WorkingHours(1);
        document.Assignments.Add(locked, resource); document.Assignments.Add(moved, resource);
        long revision = document.Revision;
        var result = document.CalculateLeveling(new ProjectLevelingOptions { MaxDelayDays = maxDays });
        Assert.Equal(!accepted, result.Report.HasErrors); Assert.Equal(revision, document.Revision);
        if (!accepted) { Assert.Throws<InvalidDataException>(() => document.ApplyLeveling(result)); return; }
        Assert.Equal(Monday.AddDays(7), result.Schedule.Tasks.Single(t => t.TaskUid == moved.Uid).Start);
        document.ApplyLeveling(result); using var copy = document.Clone();
        Assert.Equal(moved.Start, copy.Tasks.GetByUid(moved.Uid).Start);
    }

    [Fact]
    public void LevelingDelayLimitAlsoBoundsDependentCalendarMovement() {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        document.Settings.StartDate = Monday.AddDays(4).AddHours(6);
        var resource = document.Resources.AddWork("Engineer");
        var locked = document.Tasks.Add("Locked"); locked.Duration = ProjectDuration.WorkingHours(1); locked.Priority = 1000;
        var moved = document.Tasks.Add("Movable"); moved.Duration = ProjectDuration.WorkingHours(1);
        document.Assignments.Add(locked, resource); document.Assignments.Add(moved, resource);
        var calendar = document.Calendars.Add("Friday hour");
        foreach (DayOfWeek day in Enum.GetValues(typeof(DayOfWeek)))
            calendar.SetWorkingDay(day, day == DayOfWeek.Friday ? new[] { ProjectWorkingTime.Hours(15, 16) } : Array.Empty<ProjectWorkingTime>());
        var dependent = document.Tasks.Add("Dependent"); dependent.Calendar = calendar; dependent.Duration = ProjectDuration.WorkingHours(1);
        document.Dependencies.Add(moved, dependent);
        var result = document.CalculateLeveling(new ProjectLevelingOptions { MaxDelayDays = 1 });
        Assert.True(result.Report.HasErrors); Assert.Throws<InvalidDataException>(() => document.ApplyLeveling(result));
    }
}
