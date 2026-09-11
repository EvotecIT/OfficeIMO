namespace OfficeIMO.Project.Tests;

public sealed class ProjectEarnedValueTests {
    [Fact]
    public void InconsistentBaselineCostCurvesInvalidateCurveMetricsButRetainIndependentActualCost() {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        var monday = new DateTime(2026, 10, 5, 8, 0, 0); document.Settings.StartDate = monday;
        var task = document.Tasks.Add("Delivery"); task.Duration = ProjectDuration.WorkingDays(1);
        var resource = document.Resources.AddWork("Engineer"); resource.StandardRate = 100;
        document.Assignments.Add(task, resource, ProjectUnits.Fraction(1));
        document.CaptureBaseline(document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }));
        task.Baselines[0].Cost = 100m; task.PercentComplete = 50;
        var result = document.AnalyzeEarnedValue(statusDate: monday.AddHours(4)).Tasks.Single();
        Assert.Null(result.PlannedValue); Assert.Null(result.EarnedValue); Assert.Equal(0m, result.ActualCost);
        task.EarnedValueMethod = ProjectEarnedValueMethod.PhysicalPercentComplete; task.PhysicalPercentComplete = 50;
        Assert.Equal(50m, document.AnalyzeEarnedValue(statusDate: monday.AddHours(4)).Tasks.Single().EarnedValue);
    }
    [Fact]
    public void SparseWorkDoesNotReplaceFixedTaskDurationOrEraseIndependentMetrics() {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        var monday = new DateTime(2026, 10, 5, 8, 0, 0); document.Settings.StartDate = monday;
        var task = document.Tasks.Add("Delivery"); task.Type = ProjectTaskType.FixedDuration; task.Duration = ProjectDuration.WorkingDays(2); task.FixedCost = 300;
        var resource = document.Resources.AddWork("Engineer"); resource.StandardRate = 100;
        var assignment = document.Assignments.Add(task, resource, ProjectUnits.Fraction(1)); assignment.Work = ProjectWork.Hours(8); assignment.WorkContour = ProjectWorkContour.Custom;
        var curve = assignment.TimephasedData.Add(); curve.Type = 1; curve.Start = monday; curve.Finish = monday.AddHours(9); curve.Value = "PT8H";
        var schedule = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); schedule.Report.ThrowIfErrors();
        document.CaptureBaseline(schedule); task.PercentComplete = 50;
        var result = document.AnalyzeEarnedValue(statusDate: monday.AddHours(9)).Tasks.Single();
        Assert.Equal(950m, result.PlannedValue); Assert.Equal(950m, result.EarnedValue); Assert.Equal(0m, result.ActualCost);
    }
    [Fact]
    public void BaselineUsesTheIndependentResourceCalendarAndKeepsKnownMetricsWhenProgressIsIncomplete() {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        var monday = new DateTime(2026, 10, 5, 8, 0, 0); document.Settings.StartDate = monday;
        var night = document.Calendars.Add("Night");
        foreach (DayOfWeek day in Enum.GetValues(typeof(DayOfWeek))) night.SetWorkingDay(day, ProjectWorkingTime.Hours(18, 2));
        var resource = document.Resources.AddWork("Night crew"); resource.Calendar = night; resource.StandardRate = 100;
        var task = document.Tasks.Add("Delivery"); task.Duration = ProjectDuration.WorkingDays(1);
        document.Assignments.Add(task, resource, ProjectUnits.Fraction(1));
        var schedule = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); schedule.Report.ThrowIfErrors();
        document.CaptureBaseline(schedule); task.PercentComplete = 50;
        var result = document.AnalyzeEarnedValue(statusDate: new DateTime(2026, 10, 5, 22, 0, 0)).Tasks.Single();
        Assert.Equal(400m, result.PlannedValue); Assert.Equal(400m, result.EarnedValue); Assert.Equal(0m, result.ActualCost);
        task.Baselines[0].Duration = ProjectDuration.WorkingDays(2);
        var incomplete = document.AnalyzeEarnedValue(statusDate: new DateTime(2026, 10, 5, 22, 0, 0)).Tasks.Single();
        Assert.Equal(400m, incomplete.PlannedValue); Assert.Null(incomplete.EarnedValue); Assert.Equal(0m, incomplete.ActualCost);
    }
    [Fact]
    public void SplitBaselineDoesNotAccrueProratedFixedCostOrDurationProgressInsideTheGap() {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        var monday = new DateTime(2026, 10, 5, 8, 0, 0); document.Settings.StartDate = monday;
        var resource = document.Resources.AddWork("Engineer"); resource.StandardRate = 100m;
        var task = document.Tasks.Add("Delivery"); task.Duration = ProjectDuration.WorkingDays(3); task.FixedCost = 300m;
        document.Assignments.Add(task, resource, ProjectUnits.Fraction(1));
        var locked = document.Tasks.Add("Reserved"); locked.Duration = ProjectDuration.WorkingDays(1); locked.Priority = 1000;
        locked.ConstraintType = ProjectConstraintType.MustStartOn; locked.ConstraintDate = monday.AddDays(1);
        document.Assignments.Add(locked, resource, ProjectUnits.Fraction(1));
        document.ApplyLeveling(document.CalculateLeveling(new ProjectLevelingOptions { AllowSplitting = true }));
        document.CaptureBaseline(document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }));
        task.PercentComplete = 50;
        var analysis = document.AnalyzeEarnedValue(statusDate: monday.AddDays(1).AddHours(9));
        var result = analysis.Tasks.Single(t => t.TaskUid == task.Uid);
        Assert.Equal(2700m, result.BudgetAtCompletion!.Value, 6);
        Assert.Equal(900m, result.PlannedValue!.Value, 6);
        Assert.Equal(1350m, result.EarnedValue!.Value, 6);
        using var copy = document.Clone();
        var reopened = copy.AnalyzeEarnedValue(statusDate: monday.AddDays(1).AddHours(9)).Tasks.Single(t => t.TaskUid == task.Uid);
        Assert.Equal(result.PlannedValue, reopened.PlannedValue); Assert.Equal(result.EarnedValue, reopened.EarnedValue);
    }
    [Fact]
    public void BaselineCaptureUsesNumberedOwnerTypesAndRequiresExplicitReplacement() {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        var monday = new DateTime(2026, 10, 5, 8, 0, 0); document.Settings.StartDate = monday; document.Settings.StatusDate = monday.AddHours(4);
        var task = document.Tasks.Add("Delivery"); task.Duration = ProjectDuration.WorkingDays(1);
        var resource = document.Resources.AddWork("Engineer"); resource.StandardRate = 100m; resource.CostPerUse = 25m;
        var assignment = document.Assignments.Add(task, resource, ProjectUnits.Fraction(1));
        var schedule = document.Recalculate(new ProjectScheduleOptions { CalculateAssignments = true });
        schedule = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        long revision = document.Revision;
        Assert.Throws<OperationCanceledException>(() => document.CaptureBaseline(schedule, cancellationToken: new CancellationToken(true)));
        Assert.Equal(revision, document.Revision);
        Assert.Throws<InvalidOperationException>(() => document.CaptureBaseline(schedule, maxIntervals: 1)); Assert.Empty(task.Baselines);
        document.CaptureBaseline(schedule, 2);
        Assert.Equal(825m, task.Baselines.Single().Cost); Assert.Equal(480m, assignment.Baselines.Single().Work!.Value.Minutes);
        Assert.Contains(task.TimephasedData, v => v.Type == 25); Assert.Contains(assignment.TimephasedData, v => v.Type == 22);
        Assert.Contains(resource.TimephasedData, v => v.Type == 27);
        var analysis = document.AnalyzeEarnedValue(2); Assert.Equal(425m, analysis.Tasks.Single().PlannedValue);
        resource.StandardRate = 200m;
        var revised = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.Throws<InvalidOperationException>(() => document.CaptureBaseline(revised, 2));
        Assert.Equal(825m, task.Baselines.Single().Cost);
        using var copy = document.Clone(); Assert.Equal(825m, copy.Tasks.GetByUid(task.Uid).Baselines.Single().Cost);
        document.CaptureBaseline(revised, 2, overwrite: true); Assert.Equal(1625m, task.Baselines.Single().Cost);
    }
    [Fact]
    public void ProducerProgressEarnsTheBaselineCurveIncludingTheStartCharge() {
        using var document = ProjectDocument.Load(ProjectResourceCapacityTests.Fixture("progress"));
        long revision = document.Revision;
        var result = document.AnalyzeEarnedValue();
        var task = result.Tasks.Single(t => t.TaskUid == 1);
        Assert.Equal(1725m, task.BudgetAtCompletion); Assert.Equal(1725m, task.PlannedValue);
        Assert.InRange(task.EarnedValue!.Value, 874.99m, 875.01m); Assert.Equal(875m, task.ActualCost);
        Assert.InRange(task.CostPerformanceIndex!.Value, .99999m, 1.00001m);
        Assert.Equal(revision, document.Revision);
        document.Tasks.GetByUid(1).EarnedValueMethod = ProjectEarnedValueMethod.PhysicalPercentComplete;
        var physical = document.AnalyzeEarnedValue().Tasks.Single(t => t.TaskUid == 1);
        Assert.Equal(431.25m, physical.EarnedValue); Assert.Equal(task.PlannedValue, physical.PlannedValue); Assert.Equal(875m, physical.ActualCost);
    }
    [Fact]
    public void AnEarlierStatusClipsBudgetAndDoesNotInventHistoricalActualCosts() {
        using var document = ProjectDocument.Load(ProjectResourceCapacityTests.Fixture("progress"));
        var result = document.AnalyzeEarnedValue(statusDate: new DateTime(2026, 10, 5, 12, 0, 0));
        var task = result.Tasks.Single(t => t.TaskUid == 1);
        Assert.InRange(task.PlannedValue!.Value, 510.70m, 510.73m);
        Assert.Null(task.ActualCost);
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "PROJECT_EARNED_VALUE_INCOMPLETE");
        Assert.Null(document.AnalyzeEarnedValue(10).Tasks.Single(t => t.TaskUid == 1).BudgetAtCompletion);
        Assert.Throws<OperationCanceledException>(() => document.AnalyzeEarnedValue(cancellationToken: new CancellationToken(true)));
    }
}
