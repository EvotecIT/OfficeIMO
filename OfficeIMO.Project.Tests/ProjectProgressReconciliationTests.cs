namespace OfficeIMO.Project.Tests;

public sealed class ProjectProgressReconciliationTests {
    private static readonly DateTime Monday = new(2026, 10, 5, 8, 0, 0);
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FixedDurationSubtractsActualIntervalsBeforePlanningRemainingWork(bool twoAssignments) {
        using var document = Create(); var task = document.Tasks.Add("Delivery"); task.Type = ProjectTaskType.FixedDuration;
        task.Duration = ProjectDuration.WorkingHours(8);
        for (int index = 0; index < (twoAssignments ? 2 : 1); index++) {
            var resource = document.Resources.AddWork("Engineer " + index); resource.StandardRate = 10;
            var assignment = document.Assignments.Add(task, resource); assignment.Work = ProjectWork.Hours(8);
            assignment.ActualWork = ProjectWork.Hours(4); assignment.ActualStart = Monday;
        }
        for (int pass = 0; pass < 2; pass++) {
            var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); result.Report.ThrowIfErrors();
            var plan = Assert.Single(result.Tasks); Assert.Equal(ProjectDuration.WorkingHours(8), plan.Duration);
            Assert.Equal(Monday.AddHours(9), plan.Finish); Assert.Equal(4m, plan.Calculation!.ActualDuration.Value / 60m);
            document.ApplySchedule(result);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ContradictoryDurationComponentsCannotBeApplied(bool noActualDuration) {
        using var document = Create(); var task = document.Tasks.Add("Delivery"); task.Duration = ProjectDuration.WorkingHours(8);
        task.ActualStart = Monday;
        if (noActualDuration) { task.ActualFinish = Monday.AddHours(9); task.RemainingDuration = ProjectDuration.WorkingHours(0); task.PercentComplete = 100; }
        else { task.ActualDuration = ProjectDuration.WorkingHours(4); task.RemainingDuration = ProjectDuration.WorkingHours(8); }
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.True(result.Report.HasErrors); Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result));
        Assert.Equal(ProjectDuration.WorkingHours(8), task.Duration);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void ActualIntervalsCannotContradictTheirRecordedBounds(bool scalarStop, bool startConflict) {
        using var document = Create(); var task = document.Tasks.Add("Delivery"); task.Duration = ProjectDuration.WorkingHours(4);
        var resource = document.Resources.AddWork("Engineer"); resource.StandardRate = 10;
        var assignment = document.Assignments.Add(task, resource); assignment.Work = assignment.ActualWork = ProjectWork.Hours(4);
        assignment.ActualStart = Monday.AddHours(startConflict ? 1 : 0); assignment.ActualFinish = Monday.AddHours(startConflict ? 4 : 9);
        if (scalarStop) assignment.Stop = Monday.AddHours(4);
        else { var actual = assignment.TimephasedData.Add(); actual.Type = 2; actual.Start = Monday; actual.Finish = Monday.AddHours(4); actual.Value = "PT4H"; }
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.True(result.Report.HasErrors); Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result));
    }

    [Fact]
    public void WorkAssignmentsCannotStartBeforeRecordedTaskActualStart() {
        using var document = Create(); var task = document.Tasks.Add("Delivery"); task.Duration = ProjectDuration.WorkingHours(4); task.ActualStart = Monday.AddHours(2);
        var resource = document.Resources.AddWork("Engineer"); resource.StandardRate = 10;
        var assignment = document.Assignments.Add(task, resource); assignment.Work = assignment.ActualWork = ProjectWork.Hours(4); assignment.ActualStart = Monday;
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.True(result.Report.HasErrors); Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result));
    }

    [Fact]
    public void VariableMaterialUsesTheSameResolvedUnitsBeforeAndAfterApplication() {
        using var document = Create(); var task = document.Tasks.Add("Delivery"); task.Duration = ProjectDuration.WorkingDays(1);
        var resource = document.Resources.AddMaterial("Fuel"); resource.StandardRate = 10;
        var assignment = document.Assignments.Add(task, resource); assignment.Units = null; assignment.HasFixedRateUnits = false; assignment.MaterialRateScale = 3;
        for (int pass = 0; pass < 2; pass++) {
            var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); result.Report.ThrowIfErrors();
            var plan = Assert.Single(result.Assignments); Assert.Equal(1m, plan.Units.Value); Assert.Equal(1m, plan.MaterialQuantity); Assert.Equal(10m, plan.Cost);
            document.ApplySchedule(result);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CostActualsCannotFallOutsideTheFinalTaskSpan(bool curve) {
        using var document = Create(); var task = document.Tasks.Add("Delivery"); task.Duration = ProjectDuration.WorkingDays(1);
        var assignment = document.Assignments.Add(task, document.Resources.AddCost("Travel")); assignment.Cost = 100; assignment.ActualCost = 100;
        if (curve) { var point = assignment.TimephasedData.Add(); point.Type = 6; point.Start = Monday.AddDays(1); point.Finish = point.Start; point.Value = "10000"; }
        else { assignment.ActualStart = Monday.AddDays(1); assignment.ActualFinish = Monday.AddDays(1).AddHours(1); }
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.True(result.Report.HasErrors); Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result));
    }

    [Fact]
    public void EnteredCostActualDateAnchorsItsScalarCharge() {
        using var document = Create(); var task = document.Tasks.Add("Delivery"); task.Duration = ProjectDuration.WorkingDays(2);
        var assignment = document.Assignments.Add(task, document.Resources.AddCost("Travel")); assignment.Cost = 100; assignment.ActualCost = 100;
        assignment.ActualStart = Monday.AddDays(1); assignment.ActualFinish = assignment.ActualStart;
        for (int pass = 0; pass < 2; pass++) {
            var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); result.Report.ThrowIfErrors();
            var plan = Assert.Single(result.Assignments); Assert.Equal(assignment.ActualStart, Assert.Single(plan.Costs, c => c.IsActual).Start);
            document.ApplySchedule(result); Assert.True(assignment.Start <= assignment.ActualStart); Assert.True(assignment.Finish >= assignment.ActualFinish);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CostProgressCannotContradictRecordedCompletionOrCurveBounds(bool curve) {
        using var document = Create(); var task = document.Tasks.Add("Delivery"); task.Duration = ProjectDuration.WorkingDays(1);
        var assignment = document.Assignments.Add(task, document.Resources.AddCost("Travel")); assignment.Cost = 100;
        assignment.ActualStart = Monday.AddHours(2); assignment.ActualFinish = Monday.AddHours(3); assignment.ActualCost = curve ? 100 : 50;
        if (curve) { var point = assignment.TimephasedData.Add(); point.Type = 6; point.Start = Monday; point.Finish = Monday; point.Value = "10000"; }
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.True(result.Report.HasErrors); Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result));
    }

    private static ProjectDocument Create() { var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Monday; return document; }
}
