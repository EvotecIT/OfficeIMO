namespace OfficeIMO.Project.Tests;

public sealed class ProjectCostAssignmentBoundaryTests {
    private static readonly DateTime Monday = new(2026, 10, 5, 8, 0, 0);

    [Theory]
    [InlineData(false, false, 8)]
    [InlineData(true, false, 8)]
    [InlineData(false, true, 8)]
    [InlineData(true, true, 8)]
    [InlineData(false, false, 24)]
    [InlineData(true, false, 24)]
    [InlineData(false, true, 24)]
    [InlineData(true, true, 24)]
    public void MixedMaterialDurationChangesRequireAnExplicitConsumptionProjection(bool variable, bool backward, int workHours) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        document.Settings.StartDate = Monday;
        if (backward) { document.Settings.ScheduleFromStart = false; document.Settings.FinishDate = Monday.AddDays(3).AddHours(9); }
        var task = document.Tasks.Add("Delivery"); task.Type = ProjectTaskType.FixedUnits; task.Duration = ProjectDuration.WorkingDays(2);
        var engineer = document.Resources.AddWork("Engineer"); engineer.StandardRate = 100m;
        document.Assignments.Add(task, engineer, ProjectUnits.Percent(100)).Work = ProjectWork.Hours(workHours);
        var parts = document.Resources.AddMaterial("Parts"); parts.StandardRate = 10m;
        var material = document.Assignments.Add(task, parts, ProjectUnits.Fraction(3));
        material.HasFixedRateUnits = !variable; if (variable) material.MaterialRateScale = 3;
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.True(result.Report.HasErrors);
        Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result));
        Assert.Null(task.Start); Assert.Null(material.Start);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void CostAssignmentsFollowTheFinalTaskSpanAndBaseline(bool shortShifts, bool costFirst) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        document.Settings.StartDate = Monday;
        var task = document.Tasks.Add("Delivery"); task.Type = ProjectTaskType.FixedWork;
        task.Duration = ProjectDuration.WorkingDays(1);
        var labor = document.Resources.AddWork("Engineer"); labor.StandardRate = 100m;
        if (shortShifts) {
            labor.Calendar = document.Calendars.AddStandardWorkingWeek("Short shifts");
            for (int day = 1; day <= 5; day++) labor.Calendar.SetWorkingDay((DayOfWeek)day, ProjectWorkingTime.Hours(8, 12));
        }
        var expense = document.Resources.AddCost("Travel");
        ProjectAssignment cost;
        if (costFirst) { cost = document.Assignments.Add(task, expense); document.Assignments.Add(task, labor).Work = ProjectWork.Hours(16); }
        else { document.Assignments.Add(task, labor).Work = ProjectWork.Hours(16); cost = document.Assignments.Add(task, expense); }
        cost.Cost = 500m; cost.ActualCost = 100m;
        var schedule = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); schedule.Report.ThrowIfErrors();
        var taskPlan = Assert.Single(schedule.Tasks); var costPlan = schedule.Assignments.Single(a => a.AssignmentUid == cost.Uid);
        Assert.Equal(Monday.AddDays(shortShifts ? 3 : 1).AddHours(shortShifts ? 4 : 9), taskPlan.Finish);
        Assert.Equal(taskPlan.Start, costPlan.Start); Assert.Equal(taskPlan.Finish, costPlan.Finish);
        var remaining = Assert.Single(costPlan.Costs, c => !c.IsActual);
        Assert.Equal(taskPlan.Finish, remaining.Start); Assert.Equal(taskPlan.Finish, remaining.Finish); Assert.Equal(400m, remaining.Cost);
        Assert.Equal(taskPlan.Start, Assert.Single(costPlan.Costs, c => c.IsActual).Start);
        document.ApplySchedule(schedule);
        Assert.Equal(task.Finish, cost.Finish); Assert.Equal(2100m, task.Cost);
        var current = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); current.Report.ThrowIfErrors();
        document.CaptureBaseline(current);
        Assert.Equal(task.Finish, cost.Baselines.Single().Finish);
        using var reopened = document.Clone();
        Assert.Equal(task.Finish, reopened.Assignments.GetByUid(cost.Uid).Finish);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CostAssignmentsUseTheSelectedBackwardOrManualBounds(bool manual) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        document.Settings.StartDate = Monday;
        var task = document.Tasks.Add("Delivery"); task.Duration = ProjectDuration.WorkingDays(1);
        var labor = document.Resources.AddWork("Engineer"); labor.StandardRate = 100m;
        document.Assignments.Add(task, labor).Work = ProjectWork.Hours(16);
        var cost = document.Assignments.Add(task, document.Resources.AddCost("Travel")); cost.Cost = 500m;
        if (manual) { task.IsManual = true; task.Start = Monday; task.Finish = Monday.AddDays(2).AddHours(9); }
        else { document.Settings.ScheduleFromStart = false; document.Settings.FinishDate = Monday.AddDays(3).AddHours(9); }
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); result.Report.ThrowIfErrors();
        var taskPlan = Assert.Single(result.Tasks); var costPlan = result.Assignments.Single(a => a.AssignmentUid == cost.Uid);
        Assert.Equal(manual ? task.Start : Monday.AddDays(2), taskPlan.Start);
        Assert.Equal(manual ? task.Finish : document.Settings.FinishDate, taskPlan.Finish);
        Assert.Equal(taskPlan.Start, costPlan.Start); Assert.Equal(taskPlan.Finish, costPlan.Finish);
        Assert.Equal(taskPlan.Finish, Assert.Single(costPlan.Costs, c => !c.IsActual).Finish);
    }

    [Fact]
    public void SummaryAssignmentsCannotDisappearFromAnAssignmentCalculation() {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        document.Settings.StartDate = Monday;
        var summary = document.Tasks.AddSummary("Phase"); summary.Children.Add("Delivery").Duration = ProjectDuration.WorkingDays(1);
        document.Assignments.Add(summary, document.Resources.AddCost("Travel")).Cost = 500m;
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "PROJECT_SUMMARY_ASSIGNMENT_PROFILE" && d.Severity == ProjectDiagnosticSeverity.Error);
        Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result));
    }
}
