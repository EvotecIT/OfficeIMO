namespace OfficeIMO.Project.Tests;

public sealed class ProjectStopLevelingFormulaTests {
    private static readonly DateTime Monday = new(2026, 10, 5, 8, 0, 0);

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void CompletedDurationDoesNotMoveToALaterStatusDate(bool assigned, bool actualFinish) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Monday; document.Settings.StatusDate = Monday.AddDays(1);
        var task = document.Tasks.Add("Complete"); task.Type = ProjectTaskType.FixedDuration; task.Duration = task.ActualDuration = ProjectDuration.WorkingHours(8);
        task.RemainingDuration = ProjectDuration.WorkingHours(0); task.ActualStart = Monday; if (actualFinish) task.ActualFinish = Monday.AddHours(9);
        if (assigned) {
            var assignment = document.Assignments.Add(task, document.Resources.AddWork("Engineer")); assignment.Work = assignment.ActualWork = ProjectWork.Hours(8);
            assignment.ActualStart = Monday; assignment.ActualFinish = Monday.AddHours(9);
        }
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true, RescheduleRemainingAfterStatusDate = true }); result.Report.ThrowIfErrors();
        Assert.Equal(Monday.AddHours(9), result.Tasks.Single().Finish);
        document.ApplySchedule(result); using var copy = document.Clone(); Assert.Equal(task.Finish, copy.Tasks[0].Finish);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FixedDurationRemainderHonorsStopBeyondAssignmentCoverage(bool shortRemainingCurve) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Monday;
        var task = document.Tasks.Add("Fixed"); task.Type = ProjectTaskType.FixedDuration; task.Duration = ProjectDuration.WorkingHours(8);
        task.ActualDuration = ProjectDuration.WorkingHours(1); task.RemainingDuration = ProjectDuration.WorkingHours(7); task.ActualStart = Monday; task.Stop = Monday.AddDays(1);
        var assignment = document.Assignments.Add(task, document.Resources.AddWork("Engineer")); assignment.Work = ProjectWork.Hours(shortRemainingCurve ? 2 : 1);
        assignment.ActualWork = ProjectWork.Hours(1); assignment.ActualStart = Monday;
        if (shortRemainingCurve) {
            var curve = assignment.TimephasedData.Add(); curve.Uid = assignment.Uid; curve.Type = 1; curve.Start = task.Stop; curve.Finish = task.Stop.Value.AddHours(1); curve.Value = "PT1H";
        } else assignment.ActualFinish = Monday.AddHours(1);
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); result.Report.ThrowIfErrors();
        Assert.Equal(Monday.AddDays(1).AddHours(8), result.Tasks.Single().Finish);
        document.ApplySchedule(result); using var copy = document.Clone();
        Assert.Equal(result.Tasks.Single().Finish, copy.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }).Tasks.Single().Finish);
    }

    [Theory]
    [InlineData(4, false)]
    [InlineData(7, true)]
    public void LevelingBoundsRemainingWorkAfterResourceCalendarSnapping(int maxDays, bool accepted) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        var thursday = Monday.AddDays(3); document.Settings.StartDate = thursday;
        var calendar = document.Calendars.Add("Sparse resource");
        foreach (DayOfWeek day in Enum.GetValues(typeof(DayOfWeek))) calendar.SetWorkingDay(day,
            day == DayOfWeek.Thursday ? new[] { ProjectWorkingTime.Hours(8, 9) } : day == DayOfWeek.Friday ? new[] { ProjectWorkingTime.Hours(16, 17) } : Array.Empty<ProjectWorkingTime>());
        var resource = document.Resources.AddWork("Engineer"); resource.Calendar = calendar;
        var task = document.Tasks.Add("Progressed"); task.Duration = ProjectDuration.WorkingHours(2); task.ActualDuration = ProjectDuration.WorkingHours(1);
        task.ActualStart = thursday; task.Stop = thursday.AddHours(1);
        var assignment = document.Assignments.Add(task, resource); assignment.Work = ProjectWork.Hours(2); assignment.ActualWork = ProjectWork.Hours(1); assignment.ActualStart = thursday;
        var locked = document.Tasks.Add("Locked"); locked.Duration = ProjectDuration.WorkingHours(1); locked.Priority = 1000;
        locked.ConstraintType = ProjectConstraintType.MustStartOn; locked.ConstraintDate = thursday.AddDays(1).AddHours(8); document.Assignments.Add(locked, resource);
        var result = document.CalculateLeveling(new ProjectLevelingOptions { MaxDelayDays = maxDays });
        Assert.Equal(!accepted, result.Report.HasErrors);
        if (!accepted) { Assert.Throws<InvalidDataException>(() => document.ApplyLeveling(result)); return; }
        var plan = result.Schedule.Assignments.Single(a => a.AssignmentUid == assignment.Uid);
        Assert.Equal(thursday.AddDays(7), plan.Intervals.Single(i => !i.IsActual).Start);
        Assert.Equal(thursday, plan.Intervals.Single(i => i.IsActual).Start);
    }

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
