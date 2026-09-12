namespace OfficeIMO.Project.Tests;

public sealed class ProjectCostIntegrityTests {
    private static readonly DateTime Monday = new(2026, 10, 5, 8, 0, 0);

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SummaryActualCostRequiresExplicitRecalculation(bool recalculate) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Monday;
        var summary = document.Tasks.AddSummary("Phase"); summary.FixedCost = 100; summary.ActualCost = 75;
        var child = summary.Children.Add("Work"); child.Duration = ProjectDuration.WorkingDays(1);
        child.ActualStart = Monday; child.ActualDuration = new ProjectDuration(4, ProjectDurationUnit.Hour); child.Stop = Monday.AddHours(4);
        var options = new ProjectScheduleOptions { CalculateAssignments = true, RecalculateActualCosts = recalculate };
        var schedule = document.CalculateSchedule(options); schedule.Report.ThrowIfErrors();
        var totals = schedule.Tasks.Single(t => t.TaskUid == summary.Uid).Calculation!;
        Assert.Equal(recalculate ? 50m : 75m, totals.ActualCost); Assert.Equal(recalculate ? 100m : 125m, totals.Cost);
        document.ApplySchedule(schedule); Assert.Equal(recalculate ? 50m : 75m, summary.ActualCost);
        var leveling = document.CalculateLeveling(new ProjectLevelingOptions { ScheduleOptions = options }); leveling.Report.ThrowIfErrors();
        document.ApplyLeveling(leveling); Assert.Equal(recalculate ? 50m : 75m, summary.ActualCost);
        using var copy = document.Clone();
        Assert.Equal(recalculate ? 50m : 75m, copy.CalculateSchedule(options).Tasks.Single(t => t.TaskUid == summary.Uid).Calculation!.ActualCost);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BaselineCannotStoreUnlocatedActualCostAdjustment(bool summaryAdjustment) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Monday;
        var summary = document.Tasks.AddSummary("Phase"); var task = summary.Children.Add("Work"); task.Duration = ProjectDuration.WorkingDays(1);
        task.ActualStart = Monday; task.ActualDuration = ProjectDuration.WorkingDays(.5m); task.Stop = Monday.AddHours(4);
        var owner = summaryAdjustment ? summary : task; owner.FixedCost = 100; owner.ActualCost = 75;
        var schedule = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); schedule.Report.ThrowIfErrors();
        string before = document.ToXml(); long revision = document.Revision;
        Assert.Throws<InvalidOperationException>(() => document.CaptureBaseline(schedule));
        Assert.Equal(revision, document.Revision); Assert.Equal(before, document.ToXml());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void StartedCostAssignmentRequiresActualAmountEvidence(bool explicitZero) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Monday;
        var task = document.Tasks.Add("Expense"); task.Duration = ProjectDuration.WorkingDays(1);
        var assignment = document.Assignments.Add(task, document.Resources.AddCost("Rental"));
        assignment.ActualStart = Monday; assignment.Stop = Monday.AddHours(1); assignment.RemainingCost = 100;
        if (explicitZero) assignment.ActualCost = 0;
        string before = document.ToXml();
        var schedule = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        if (explicitZero) {
            schedule.Report.ThrowIfErrors(); document.ApplySchedule(schedule);
            Assert.Equal(0m, assignment.ActualCost); Assert.Equal(100m, assignment.Cost);
        } else {
            Assert.True(schedule.Report.HasErrors); Assert.Throws<InvalidDataException>(() => document.ApplySchedule(schedule));
            Assert.Equal(before, document.ToXml());
        }
    }

    [Fact]
    public void MissingSummaryActualCostDoesNotHideRecordedChildCosts() {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        var summary = document.Tasks.AddSummary("Phase"); var baseline = summary.Baselines.Add(); baseline.Number = 0; baseline.Cost = 100;
        var child = summary.Children.Add("Work"); child.ActualStart = Monday; child.Stop = Monday.AddHours(1); child.ActualCost = 50;
        var result = document.AnalyzeEarnedValue(statusDate: Monday.AddHours(2)).Tasks.Single(t => t.TaskUid == summary.Uid);
        Assert.Null(result.ActualCost);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void PartialOrMissingSelectedRateCoverageBlocksApplicationAndBaselineCapture(bool missingTable, bool recalculateActuals) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Monday;
        var task = document.Tasks.Add("Work"); task.Duration = ProjectDuration.WorkingDays(1);
        var resource = document.Resources.AddWork("Engineer"); resource.StandardRate = 100;
        var assignment = document.Assignments.Add(task, resource); assignment.Cost = 100;
        if (missingTable) assignment.CostRateTable = ProjectCostRateTable.B;
        else { var rate = resource.Rates.Add(); rate.From = Monday; rate.To = Monday.AddHours(4); rate.StandardRate = 100; }
        var options = new ProjectScheduleOptions { CalculateAssignments = true, RecalculateActualCosts = recalculateActuals };
        string before = document.ToXml();
        var result = document.CalculateSchedule(options);
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "PROJECT_COST_RECALCULATION_INCOMPLETE");
        Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result));
        Assert.Throws<InvalidDataException>(() => document.CaptureBaseline(result));
        var leveling = document.CalculateLeveling(new ProjectLevelingOptions { ScheduleOptions = options });
        Assert.True(leveling.Report.HasErrors); Assert.Throws<InvalidDataException>(() => document.ApplyLeveling(leveling));
        Assert.Equal(before, document.ToXml());
    }

    [Theory]
    [InlineData("assignment remaining")]
    [InlineData("task remaining")]
    [InlineData("task actual")]
    [InlineData("resource actual")]
    public void UnknownCostComponentsCannotReplaceRecordedComponents(string owner) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Monday;
        var task = document.Tasks.Add("Work"); task.Duration = ProjectDuration.WorkingDays(1);
        var resource = document.Resources.AddWork("Engineer"); var assignment = document.Assignments.Add(task, resource);
        if (owner == "assignment remaining") assignment.RemainingCost = 100;
        if (owner == "task remaining") task.RemainingCost = 100;
        if (owner == "task actual") task.ActualCost = 100;
        if (owner == "resource actual") resource.ActualCost = 100;
        string before = document.ToXml();
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true, RecalculateActualCosts = true });
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "PROJECT_COST_RECALCULATION_INCOMPLETE");
        Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result)); Assert.Equal(before, document.ToXml());
    }

    [Fact]
    public void SummaryWithMixedDurationBasesDoesNotGuessPartialBaselineCosts() {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Monday;
        var summary = document.Tasks.AddSummary("Phase");
        var elapsed = summary.Children.Add("Cure"); elapsed.Duration = new ProjectDuration(24, ProjectDurationUnit.Hour, elapsed: true); elapsed.FixedCost = 240;
        var working = summary.Children.Add("Inspection"); working.Duration = ProjectDuration.WorkingDays(1); working.FixedCost = 80;
        document.CaptureBaseline(document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }));
        summary.PercentComplete = 50;
        var result = document.AnalyzeEarnedValue(statusDate: Monday.AddHours(12));
        var value = result.Tasks.Single(t => t.TaskUid == summary.Uid);
        Assert.Null(value.PlannedValue); Assert.Null(value.EarnedValue);
        Assert.Equal(120m, result.Tasks.Single(t => t.TaskUid == elapsed.Uid).PlannedValue);
        Assert.Equal(320m, document.AnalyzeEarnedValue(statusDate: Monday.AddDays(1)).Tasks.Single(t => t.TaskUid == summary.Uid).PlannedValue);
    }

    [Fact]
    public void ElapsedActualCostCurvesIncludeNonworkingHours() {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        var task = document.Tasks.Add("Cure"); task.Duration = new ProjectDuration(24, ProjectDurationUnit.Hour, elapsed: true);
        var baseline = task.Baselines.Add(); baseline.Number = 0; baseline.Cost = 240;
        var assignment = document.Assignments.Add(task, document.Resources.AddCost("Rental")); assignment.ActualCost = 240;
        var curve = assignment.TimephasedData.Add(); curve.Uid = assignment.Uid; curve.Type = 6;
        curve.Start = Monday; curve.Finish = Monday.AddDays(1); curve.Value = "24000";
        Assert.Equal(120m, Assert.Single(document.AnalyzeEarnedValue(statusDate: Monday.AddHours(12)).Tasks).ActualCost);
    }

    [Fact]
    public void StartedAssignmentWithAbsentActualCostRemainsUnknown() {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        var task = document.Tasks.Add("Work"); var baseline = task.Baselines.Add(); baseline.Number = 0; baseline.Cost = 100;
        var assignment = document.Assignments.Add(task, document.Resources.AddWork("Engineer")); assignment.ActualStart = Monday; assignment.Stop = Monday.AddHours(1);
        var result = document.AnalyzeEarnedValue(statusDate: Monday.AddHours(2));
        Assert.Null(Assert.Single(result.Tasks).ActualCost);
        assignment.ActualCost = 0;
        Assert.Equal(0m, Assert.Single(document.AnalyzeEarnedValue(statusDate: Monday.AddHours(2)).Tasks).ActualCost);
    }

    [Theory]
    [InlineData("assignment", false)]
    [InlineData("assignment", true)]
    [InlineData("task", false)]
    [InlineData("task", true)]
    [InlineData("resource", false)]
    [InlineData("resource", true)]
    [InlineData("summary", false)]
    [InlineData("summary", true)]
    public void MissingRatesCannotEraseEnteredCosts(string owner, bool material) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        document.Settings.StartDate = Monday;
        var summary = document.Tasks.AddSummary("Phase");
        var task = summary.Children.Add("Work"); task.Duration = ProjectDuration.WorkingDays(1);
        var resource = material ? document.Resources.AddMaterial("Parts") : document.Resources.AddWork("Engineer");
        var assignment = document.Assignments.Add(task, resource);
        if (owner == "assignment") assignment.Cost = 100;
        if (owner == "task") task.Cost = 100;
        if (owner == "resource") resource.Cost = 100;
        if (owner == "summary") summary.Cost = 100;
        string before = document.ToXml(); long revision = document.Revision;
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.True(result.Report.HasErrors);
        Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result));
        Assert.Equal(revision, document.Revision); Assert.Equal(before, document.ToXml());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MissingRatesCanStillCalculateWorkWhenNoCostIsBeingReplaced(bool material) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Monday;
        var task = document.Tasks.Add("Work"); task.Duration = ProjectDuration.WorkingDays(1);
        var resource = material ? document.Resources.AddMaterial("Parts") : document.Resources.AddWork("Engineer");
        var assignment = document.Assignments.Add(task, resource);
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); result.Report.ThrowIfErrors();
        Assert.Null(Assert.Single(result.Assignments).Cost);
        document.ApplySchedule(result); Assert.Null(assignment.Cost); Assert.NotNull(assignment.Work);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ElapsedBaselineCostsUseTheirCapturedTimeBasis(bool changeCurrentDuration) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Monday;
        var task = document.Tasks.Add("Cure"); task.Duration = new ProjectDuration(24, ProjectDurationUnit.Hour, elapsed: true); task.FixedCost = 240;
        document.CaptureBaseline(document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }));
        task.PercentComplete = 50;
        if (changeCurrentDuration) task.Duration = ProjectDuration.WorkingDays(1);
        var result = Assert.Single(document.AnalyzeEarnedValue(statusDate: Monday.AddHours(12)).Tasks);
        Assert.Equal(120m, result.PlannedValue); Assert.Equal(120m, result.EarnedValue);
        using var copy = document.Clone();
        Assert.Equal(120m, Assert.Single(copy.AnalyzeEarnedValue(statusDate: Monday.AddHours(12)).Tasks).PlannedValue);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MissingTaskActualCostDoesNotImplyZeroFixedCost(bool assigned) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        var task = document.Tasks.Add("Started work"); task.FixedCost = 100; task.FixedCostAccrual = ProjectCostAccrual.Start;
        task.ActualStart = Monday; task.Stop = Monday.AddHours(1); task.PercentComplete = 50;
        var baseline = task.Baselines.Add(); baseline.Number = 0; baseline.Cost = 200;
        if (assigned) {
            var assignment = document.Assignments.Add(task, document.Resources.AddWork("Engineer"));
            assignment.ActualCost = 50; assignment.ActualStart = Monday; assignment.Stop = task.Stop;
        }
        var result = document.AnalyzeEarnedValue(statusDate: Monday.AddHours(2));
        var value = Assert.Single(result.Tasks);
        Assert.Null(value.ActualCost); Assert.Null(value.CostPerformanceIndex); Assert.Null(value.EstimateAtCompletion);
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "PROJECT_EARNED_VALUE_INCOMPLETE");
        task.ActualCost = assigned ? 150m : 100m;
        Assert.Equal(task.ActualCost, Assert.Single(document.AnalyzeEarnedValue(statusDate: Monday.AddHours(2)).Tasks).ActualCost);
    }
}
