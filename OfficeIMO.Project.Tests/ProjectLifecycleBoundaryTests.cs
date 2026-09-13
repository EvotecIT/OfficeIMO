namespace OfficeIMO.Project.Tests;

public sealed class ProjectLifecycleBoundaryTests {
    private static readonly DateTime Monday = new(2026, 10, 5, 8, 0, 0);
    [Theory]
    [InlineData(false, false, false)]
    [InlineData(true, false, false)]
    [InlineData(false, true, true)]
    public void EarnedValueSeparatesFullActualTotalsFromStatusClipping(bool taskScalar, bool assignmentScalar, bool partial) {
        using var document = Create(); var task = document.Tasks.Add("Delivery");
        task.ActualStart = Monday; task.Stop = Monday.AddHours(partial ? 1 : 2); if (taskScalar) task.ActualCost = 100m;
        var baseline = task.Baselines.Add(); baseline.Number = 0; baseline.Cost = 100m;
        var assignment = document.Assignments.Add(task, document.Resources.AddWork("Engineer")); if (assignmentScalar) assignment.ActualCost = 100m;
        AddActualCost(assignment, 100m, Monday, Monday.AddHours(2));
        var value = Assert.Single(document.AnalyzeEarnedValue(statusDate: Monday.AddHours(partial ? 1 : 2)).Tasks);
        Assert.Equal(partial ? 50m : 100m, value.ActualCost);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void EnteredCostResourceCurvesSurviveCalculationAndApplication(bool scalar, bool recalculate) {
        using var document = Create(); var task = document.Tasks.Add("Delivery"); task.Duration = ProjectDuration.WorkingDays(1);
        var assignment = document.Assignments.Add(task, document.Resources.AddCost("Travel")); assignment.Cost = 500m;
        if (scalar) assignment.ActualCost = 100m;
        AddActualCost(assignment, 100m, Monday.AddHours(2), Monday.AddHours(2));
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true, RecalculateActualCosts = recalculate }); result.Report.ThrowIfErrors();
        var plan = Assert.Single(result.Assignments); Assert.Equal(100m, plan.ActualCost); Assert.Equal(400m, plan.RemainingCost);
        Assert.Equal(Monday.AddHours(2), Assert.Single(plan.Costs, c => c.IsActual).Start);
        document.ApplySchedule(result); Assert.Equal(100m, assignment.ActualCost);
        using var copy = document.Clone();
        Assert.Equal("10000", Assert.Single(copy.Assignments.Single().TimephasedData, c => c.Type == 6).Value);
    }

    [Fact]
    public void ContradictoryCostResourceActualsCannotBeApplied() {
        using var document = Create(); var task = document.Tasks.Add("Delivery"); task.Duration = ProjectDuration.WorkingDays(1);
        var assignment = document.Assignments.Add(task, document.Resources.AddCost("Travel")); assignment.Cost = 500m; assignment.ActualCost = 90m;
        AddActualCost(assignment, 100m, Monday, Monday);
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.True(result.Report.HasErrors); Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result));
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void DateOnlyApplicationCannotStrandEarlierWorkCurves(int curveType) {
        using var document = Create(); document.Settings.StartDate = Monday.AddDays(7);
        var task = document.Tasks.Add("Delivery"); task.Duration = ProjectDuration.WorkingDays(1);
        var assignment = document.Assignments.Add(task, document.Resources.AddWork("Engineer")); assignment.Work = ProjectWork.Hours(8);
        var curve = assignment.TimephasedData.Add(); curve.Uid = assignment.Uid; curve.Type = curveType; curve.Start = Monday; curve.Finish = Monday.AddHours(9); curve.Value = "PT8H";
        var result = document.CalculateSchedule(); Assert.True(result.Report.HasErrors);
        Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result)); Assert.Null(assignment.Start); Assert.Equal(Monday, curve.Start);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ProjectSummaryUsageAndBaselineIncludeRootAssignments(bool resourceFilter) {
        var fields = new[] { ProjectDataField.Uid, ProjectDataField.Name, ProjectDataField.Summary, ProjectDataField.DurationMinutes };
        var table = new ProjectDataTable(fields.Select(f => f.ToString()), new[] { new[] { "1", "Delivery", "false", "480" }, new[] { "0", "Project", "true", "480" } });
        using var document = ProjectDocument.ImportTables(new[] { new ProjectMappedTable(ProjectDataKind.Tasks, table, fields.Select(f => new ProjectDataColumn(f, f.ToString()))) }).Document;
        document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Monday;
        var resource = document.Resources.AddWork("Engineer"); resource.StandardRate = 100m;
        document.Assignments.Add(document.Tasks.GetByUid(1), resource);
        var schedule = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); schedule.Report.ThrowIfErrors();
        var view = document.CreateView(schedule, new ProjectViewOptions { Kind = ProjectViewKind.TaskUsage, TaskUids = new[] { 0 }, ResourceUids = resourceFilter ? new[] { resource.Uid } : null });
        var summary = Assert.Single(view.Rows); Assert.Equal(8m, summary.WorkHours); Assert.Equal(8m, summary.BucketWorkHours.Sum());
        document.CaptureBaseline(schedule);
        Assert.Equal(400m, document.AnalyzeEarnedValue(statusDate: Monday.AddHours(4)).Tasks.Single(t => t.TaskUid == 0).PlannedValue);
    }

    [Fact]
    public void FinishedTasksCannotSilentlyRevertToUnstartedDuration() {
        using var document = Create(); var task = document.Tasks.Add("Done"); task.Duration = ProjectDuration.WorkingDays(1);
        task.ActualStart = Monday; task.ActualFinish = Monday.AddHours(9); task.PercentComplete = 100;
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.True(result.Report.HasErrors); Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result)); Assert.Equal(100, task.PercentComplete);
    }

    private static ProjectDocument Create() { var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Monday; return document; }
    private static void AddActualCost(ProjectAssignment assignment, decimal amount, DateTime start, DateTime finish) {
        var curve = assignment.TimephasedData.Add(); curve.Uid = assignment.Uid; curve.Type = 6; curve.Start = start; curve.Finish = finish;
        curve.Value = (amount * 100m).ToString(System.Globalization.CultureInfo.InvariantCulture);
    }
}
