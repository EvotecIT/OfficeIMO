namespace OfficeIMO.Project.Tests;

public sealed class ProjectRemainingInputTests {
    private static readonly DateTime Monday = new(2026, 10, 5, 8, 0, 0);

    [Theory]
    [InlineData(0)]
    [InlineData(60)]
    public void FinishOnlyProgressCannotProduceAnInvertedRange(int actualMinutes) {
        using var document = Create(); document.Settings.StartDate = Monday.AddDays(7);
        var task = document.Tasks.Add("Completed"); task.Duration = ProjectDuration.WorkingMinutes(actualMinutes);
        task.ActualDuration = task.Duration; task.RemainingDuration = ProjectDuration.WorkingMinutes(0);
        task.ActualFinish = Monday; task.PercentComplete = 100; task.FixedCost = 100; task.FixedCostAccrual = ProjectCostAccrual.End;
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        if (actualMinutes != 0) { Assert.True(result.Report.HasErrors); Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result)); return; }
        result.Report.ThrowIfErrors(); var plan = Assert.Single(result.Tasks);
        Assert.Equal(Monday, plan.Start); Assert.Equal(Monday, plan.Finish);
        Assert.Equal(100m, plan.Calculation!.ActualCost);
        document.ApplySchedule(result); Assert.Equal(Monday, task.Start); Assert.Equal(100, task.PercentComplete);
    }

    [Theory]
    [InlineData(false, 2)]
    [InlineData(true, 2)]
    [InlineData(false, 4)]
    public void MaterialQuantityCannotContradictDeclaredFixedUnits(bool remainingOnly, int quantity) {
        using var document = Create(); var task = document.Tasks.Add("Consume parts"); task.Duration = ProjectDuration.WorkingDays(1);
        var resource = document.Resources.AddMaterial("Parts"); resource.StandardRate = 10;
        var assignment = document.Assignments.Add(task, resource, ProjectUnits.Fraction(3)); assignment.HasFixedRateUnits = true;
        if (remainingOnly) assignment.RemainingWork = ProjectWork.Hours(quantity); else assignment.Work = ProjectWork.Hours(quantity);
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.True(result.Report.HasErrors); Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result));
        Assert.Equal(3m, assignment.Units!.Value.Value);
    }

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(true, false, false)]
    [InlineData(false, true, false)]
    [InlineData(false, false, true)]
    [InlineData(true, true, false)]
    public void CostResourceRemainingInputSurvivesApplicationAndXml(bool total, bool actual, bool curve) {
        using var document = Create(); var task = document.Tasks.Add("Delivery"); task.Duration = ProjectDuration.WorkingDays(1);
        var assignment = document.Assignments.Add(task, document.Resources.AddCost("Travel")); assignment.RemainingCost = 400;
        if (total) assignment.Cost = 500; if (actual) assignment.ActualCost = 100;
        if (curve) { var item = assignment.TimephasedData.Add(); item.Uid = assignment.Uid; item.Type = 6; item.Start = Monday; item.Finish = Monday; item.Value = "10000"; }
        decimal expectedActual = total || actual || curve ? 100 : 0;
        for (int pass = 0; pass < 2; pass++) {
            var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); result.Report.ThrowIfErrors();
            var plan = Assert.Single(result.Assignments); Assert.Equal(expectedActual + 400, plan.Cost);
            Assert.Equal(expectedActual, plan.ActualCost); Assert.Equal(400m, plan.RemainingCost);
            Assert.Equal(400m, plan.Costs.Where(c => !c.IsActual).Sum(c => c.Cost)); document.ApplySchedule(result);
        }
        using var copy = document.Clone(); var repeated = copy.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); repeated.Report.ThrowIfErrors();
        Assert.Equal(400m, Assert.Single(repeated.Assignments).RemainingCost);
    }

    [Fact]
    public void ContradictoryCostResourceRemainingInputCannotBeApplied() {
        using var document = Create(); var task = document.Tasks.Add("Delivery"); task.Duration = ProjectDuration.WorkingDays(1);
        var assignment = document.Assignments.Add(task, document.Resources.AddCost("Travel"));
        assignment.Cost = 500; assignment.ActualCost = 100; assignment.RemainingCost = 450;
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.True(result.Report.HasErrors); Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result));
    }

    [Fact]
    public void CostResourceCannotDiscardEnteredWork() {
        using var document = Create(); var task = document.Tasks.Add("Delivery"); task.Duration = ProjectDuration.WorkingDays(1);
        var assignment = document.Assignments.Add(task, document.Resources.AddCost("Travel")); assignment.Work = ProjectWork.Hours(2);
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.True(result.Report.HasErrors); Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result));
    }

    [Fact]
    public void CriticalTotalsKeepUnassignedWorkAndSummaryFixedCostWithoutDoubleCounting() {
        using var document = Create(); var summary = document.Tasks.AddSummary("Phase"); summary.FixedCost = 20;
        var longer = summary.Children.Add("Critical"); longer.Duration = ProjectDuration.WorkingDays(2); longer.Work = ProjectWork.Hours(4); longer.FixedCost = 100;
        var shorter = summary.Children.Add("Other"); shorter.Duration = ProjectDuration.WorkingDays(1); shorter.Work = ProjectWork.Hours(3); shorter.FixedCost = 50;
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); result.Report.ThrowIfErrors();
        var view = document.CreateView(result, new ProjectViewOptions { CriticalOnly = true, TaskUids = new[] { summary.Uid } });
        Assert.Equal(4m, Assert.Single(view.Rows).WorkHours); Assert.Equal(120m, view.Rows.Single().Cost);
    }

    [Theory]
    [InlineData(ProjectViewKind.TaskUsage, false)]
    [InlineData(ProjectViewKind.Gantt, false)]
    [InlineData(ProjectViewKind.TaskUsage, true)]
    public void CriticalSummaryIncludesOnlyCriticalDescendantWorkAndCost(ProjectViewKind kind, bool resourceFilter) {
        using var document = Create(); var summary = document.Tasks.AddSummary("Delivery phase");
        var longer = summary.Children.Add("Critical delivery"); longer.Duration = ProjectDuration.WorkingDays(2); longer.FixedCost = 100;
        var shorter = summary.Children.Add("Optional delivery"); shorter.Duration = ProjectDuration.WorkingDays(1); shorter.FixedCost = 200;
        var resource = document.Resources.AddWork("Engineer"); resource.StandardRate = 10;
        document.Assignments.Add(longer, resource); document.Assignments.Add(shorter, resource);
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); result.Report.ThrowIfErrors();
        Assert.True(result.Tasks.Single(t => t.TaskUid == longer.Uid).IsCritical); Assert.False(result.Tasks.Single(t => t.TaskUid == shorter.Uid).IsCritical);
        var row = Assert.Single(document.CreateView(result, new ProjectViewOptions { Kind = kind, CriticalOnly = true,
            TaskUids = new[] { summary.Uid }, ResourceUids = resourceFilter ? new[] { resource.Uid } : null }).Rows);
        Assert.Equal(16m, row.WorkHours); Assert.Equal(resourceFilter ? 160m : 260m, row.Cost);
        Assert.Equal(16m, row.BucketWorkHours.Sum());
    }

    private static ProjectDocument Create() { var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Monday; return document; }
}
