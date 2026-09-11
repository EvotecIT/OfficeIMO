namespace OfficeIMO.Project.Tests;

public sealed class ProjectProfileQualificationTests {
    private static readonly DateTime Monday = new(2026, 10, 5, 8, 0, 0);
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MissingResourceTypeBlocksSchedulingWithoutChangingStoredSource(bool assignments) {
        using var document = Create(); var task = document.Tasks.Add("Delivery"); task.Duration = ProjectDuration.WorkingDays(1);
        var resource = document.Resources.AddWork("Unspecified"); resource.Type = null;
        document.Assignments.Add(task, resource).Work = ProjectWork.Hours(8);
        using var copy = document.Clone(); Assert.Null(copy.Resources.GetByUid(resource.Uid).Type);
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = assignments });
        Assert.True(result.Report.HasErrors); Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result)); Assert.Null(resource.Type); Assert.Null(task.Work);
    }
    [Theory]
    [InlineData("standard", ProjectResourceType.Work)]
    [InlineData("overtime", ProjectResourceType.Work)]
    [InlineData("per-use", ProjectResourceType.Work)]
    [InlineData("standard", ProjectResourceType.Material)]
    [InlineData("overtime", ProjectResourceType.Material)]
    [InlineData("per-use", ProjectResourceType.Material)]
    public void NegativeFallbackRatesBlockValidationSaveAndCalculation(string field, ProjectResourceType type) {
        using var document = Create(); var task = document.Tasks.Add("Delivery"); task.Duration = ProjectDuration.WorkingDays(1);
        var resource = document.Resources.AddWork("Resource"); resource.Type = type; resource.StandardRate = 100m;
        if (field == "standard") resource.StandardRate = -1m; else if (field == "overtime") resource.OvertimeRate = -1m; else resource.CostPerUse = -1m;
        document.Assignments.Add(task, resource);
        Assert.Contains(document.Validate().Diagnostics, d => d.Code == "PROJECT_RATE_VALUE"); Assert.True(document.AssessSave().HasErrors);
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); Assert.True(result.Report.HasErrors);
        Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result));
    }
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void ManualAssignmentsUseStoredStartInBothDirections(bool backward, bool missingDuration) {
        using var document = Create(); if (backward) { document.Settings.ScheduleFromStart = false; document.Settings.FinishDate = Monday.AddDays(4).AddHours(9); }
        var task = document.Tasks.Add("Manual"); task.IsManual = true; task.Start = Monday.AddDays(1); task.Finish = Monday.AddDays(1).AddHours(9);
        if (!missingDuration) task.Duration = ProjectDuration.WorkingDays(1);
        var assignment = document.Assignments.Add(task, document.Resources.AddWork("Engineer"));
        document.Recalculate(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.Equal(Monday.AddDays(1), task.Start); Assert.Equal(task.Start, assignment.Start); Assert.Equal(task.Finish, assignment.Finish); Assert.Equal(480m, task.Work!.Value.Minutes);
        document.Recalculate(new ProjectScheduleOptions { CalculateAssignments = true }); Assert.Equal(Monday.AddDays(1), assignment.Start);
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ManualAssignmentsRejectWorkOutsideTheStoredSpan(bool dependencyConflict) {
        using var document = Create(); var task = document.Tasks.Add("Manual"); task.IsManual = true; task.Start = Monday.AddDays(1); task.Finish = Monday.AddDays(1).AddHours(9); task.Duration = ProjectDuration.WorkingDays(1);
        var assignment = document.Assignments.Add(task, document.Resources.AddWork("Engineer")); assignment.Work = ProjectWork.Hours(dependencyConflict ? 8 : 16);
        if (dependencyConflict) { var predecessor = document.Tasks.Add("Predecessor"); predecessor.Duration = ProjectDuration.WorkingDays(3); document.Dependencies.Add(predecessor, task); }
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); Assert.True(result.Report.HasErrors);
        Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result)); Assert.Equal(Monday.AddDays(1), task.Start);
    }
    [Theory]
    [InlineData("week", 0, true)]
    [InlineData("week", 1, false)]
    [InlineData("exception", 0, true)]
    [InlineData("exception", 1, false)]
    [InlineData("legacy", 0, true)]
    [InlineData("legacy", 1, false)]
    public void CalendarOverridesUseNonoverlappingInclusiveDates(string kind, int gapDays, bool invalid) {
        using var document = Create(); var calendar = document.Calendar!;
        if (kind == "week") {
            var first = calendar.WorkWeeks.Add(); first.FromDate = Monday.Date; first.ToDate = Monday.AddDays(1).Date;
            var second = calendar.WorkWeeks.Add(); second.FromDate = Monday.AddDays(1 + gapDays).Date; second.ToDate = Monday.AddDays(3).Date;
        } else {
            var first = calendar.Exceptions.Add(); first.FromDate = Monday.Date; first.ToDate = Monday.AddDays(1).Date; first.IsWorking = false;
            if (kind == "exception") { var second = calendar.Exceptions.Add(); second.FromDate = Monday.AddDays(1 + gapDays).Date; second.ToDate = Monday.AddDays(3).Date; second.IsWorking = false; }
            else { var second = calendar.WeekDays.Add(); second.FromDate = Monday.AddDays(1 + gapDays).Date; second.ToDate = Monday.AddDays(3).Date; second.IsWorking = false; }
        }
        Assert.Equal(invalid, document.Validate().HasErrors); Assert.Equal(invalid, document.AssessSave().HasErrors);
    }
    private static ProjectDocument Create() { var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Monday; return document; }
}
