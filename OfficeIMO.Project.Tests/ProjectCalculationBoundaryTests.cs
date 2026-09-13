namespace OfficeIMO.Project.Tests;

public sealed class ProjectCalculationBoundaryTests {
    private static readonly DateTime Monday = new(2026, 10, 5, 8, 0, 0);

    [Theory]
    [InlineData(50, false)]
    [InlineData(200, false)]
    [InlineData(50, true)]
    [InlineData(200, true)]
    public void WorkPerUseCostScalesByAllocationAcrossAnalysisAndScheduling(int percent, bool omitUnits) {
        using var document = ProjectDocument.Create();
        document.Calendar = document.Calendars.AddStandardWorkingWeek();
        document.Settings.StartDate = Monday;
        var task = document.Tasks.Add("Delivery"); task.Duration = ProjectDuration.WorkingDays(1);
        var resource = document.Resources.AddWork("Engineer"); resource.StandardRate = 100m; resource.CostPerUse = 25m;
        var assignment = document.Assignments.Add(task, resource, ProjectUnits.Percent(percent));
        if (omitUnits) { assignment.Units = null; resource.MaxUnits = ProjectUnits.Percent(percent); }
        assignment.Work = ProjectWork.Hours(8);
        decimal expected = 800m + 25m * percent / 100m;

        Assert.Equal(expected, Assert.Single(document.AnalyzeAssignments().Assignments).EstimatedCost);
        var schedule = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        schedule.Report.ThrowIfErrors();
        Assert.Equal(expected, Assert.Single(schedule.Assignments).Cost);
        document.ApplySchedule(schedule);
        Assert.Equal(expected, assignment.Cost); Assert.Equal(expected, task.Cost); Assert.Equal(expected, resource.Cost);
        document.CaptureBaseline(document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }));
        Assert.Equal(expected, Assert.Single(task.Baselines).Cost);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FixedMaterialMilestoneKeepsPointQuantityAndCostAcrossXml(bool completed) {
        using var document = ProjectDocument.Create();
        document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Monday;
        var task = document.Tasks.Add("Accept three parts"); task.Duration = ProjectDuration.WorkingDays(0); task.IsMilestone = true;
        var resource = document.Resources.AddMaterial("Parts"); resource.StandardRate = 10m; resource.CostPerUse = 25m;
        var assignment = document.Assignments.Add(task, resource, ProjectUnits.Fraction(3m));
        assignment.HasFixedRateUnits = true;
        if (completed) {
            assignment.Work = ProjectWork.Hours(3); assignment.ActualWork = ProjectWork.Hours(3);
            assignment.ActualStart = Monday; assignment.ActualFinish = Monday;
            task.ActualStart = Monday; task.ActualFinish = Monday;
            var actual = assignment.TimephasedData.Add(); actual.Type = 2; actual.Uid = assignment.Uid;
            actual.Start = Monday; actual.Finish = Monday; actual.Unit = 1; actual.Value = "PT3H0M0S";
        }

        var options = new ProjectScheduleOptions { CalculateAssignments = true };
        var schedule = document.CalculateSchedule(options); schedule.Report.ThrowIfErrors();
        var plan = Assert.Single(schedule.Assignments);
        Assert.Equal(3m, plan.MaterialQuantity); Assert.Equal(55m, plan.Cost);
        Assert.Equal(completed ? 55m : 0m, plan.ActualCost);
        var consumption = Assert.Single(plan.Intervals);
        Assert.Equal(Monday, consumption.Start); Assert.Equal(Monday, consumption.Finish);
        Assert.Equal(180m, consumption.Work.Minutes); Assert.Equal(0m, consumption.OvertimeWork.Minutes);
        Assert.Equal(completed, consumption.IsActual);
        document.ApplySchedule(schedule);
        Assert.Equal(0m, task.Work!.Value.Minutes); Assert.Equal(55m, task.Cost);
        using var xml = new MemoryStream(); document.Save(xml); xml.Position = 0;
        using var reopened = ProjectDocument.Load(xml);
        var repeated = reopened.CalculateSchedule(options); repeated.Report.ThrowIfErrors();
        var repeatedPlan = Assert.Single(repeated.Assignments);
        Assert.Equal(plan.MaterialQuantity, repeatedPlan.MaterialQuantity); Assert.Equal(plan.Cost, repeatedPlan.Cost);
        Assert.Equal(plan.ActualCost, repeatedPlan.ActualCost);
        reopened.CaptureBaseline(repeated);
        Assert.Equal(55m, Assert.Single(reopened.Tasks[0].Baselines).Cost);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EarnedValueIsZeroBeforeRecordedWorkAndUnknownInsideHistoricalProgress(bool physical) {
        using var document = ProjectDocument.Load(ProjectResourceCapacityTests.Fixture("progress"));
        var task = document.Tasks.GetByUid(1);
        task.EarnedValueMethod = physical ? ProjectEarnedValueMethod.PhysicalPercentComplete : ProjectEarnedValueMethod.PercentComplete;
        long revision = document.Revision;

        var before = document.AnalyzeEarnedValue(statusDate: Monday.AddMinutes(-1));
        Assert.Equal(0m, before.Tasks.Single(t => t.TaskUid == task.Uid).EarnedValue);
        var historical = document.AnalyzeEarnedValue(statusDate: Monday.AddHours(4));
        var value = historical.Tasks.Single(t => t.TaskUid == task.Uid);
        Assert.Null(value.EarnedValue); Assert.Null(value.CostPerformanceIndex); Assert.Null(value.EstimateAtCompletion);
        Assert.Contains(historical.Report.Diagnostics, d => d.Code == "PROJECT_EARNED_VALUE_INCOMPLETE");
        Assert.NotNull(document.AnalyzeEarnedValue().Tasks.Single(t => t.TaskUid == task.Uid).EarnedValue);
        Assert.Equal(revision, document.Revision);
    }

    [Fact]
    public void EarnedValueRecognizesCompletionBeforeALaterProjectStatusDate() {
        using var document = ProjectDocument.Load(ProjectResourceCapacityTests.Fixture("progress"));
        var task = document.Tasks.GetByUid(1);
        task.PercentComplete = 100; task.ActualFinish = Monday.AddHours(9);
        var result = document.AnalyzeEarnedValue(statusDate: Monday.AddDays(1)).Tasks.Single(t => t.TaskUid == task.Uid);
        Assert.Equal(result.BudgetAtCompletion, result.EarnedValue);
    }

    [Fact]
    public void IntervalLimitAllowsTheExactCountAndRejectsTheNextAssignment() {
        using var document = ProjectDocument.Create();
        document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Monday;
        var task = document.Tasks.Add("Delivery"); task.Duration = ProjectDuration.WorkingMinutes(60);
        for (int index = 0; index < 2; index++) {
            var resource = document.Resources.AddWork("Crew " + index); resource.StandardRate = 60m;
            document.Assignments.Add(task, resource, ProjectUnits.Percent(100));
        }
        var exact = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true, MaxIntervals = 4 });
        exact.Report.ThrowIfErrors();
        Assert.Equal(4, exact.Assignments.Sum(a => a.Intervals.Count + a.Costs.Count));
        var exceeded = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true, MaxIntervals = 3 });
        Assert.True(exceeded.Report.HasErrors);
        Assert.Throws<InvalidDataException>(() => document.ApplySchedule(exceeded));
    }
}
