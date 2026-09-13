namespace OfficeIMO.Project.Tests;

public sealed class ProjectAssignmentAnalysisTests {
    [Fact]
    public void NativeCostResourcesRetainEnteredCostsWithoutInterpretingRateTables() {
        using var document = ProjectDocument.Load(ProjectNativeTests.Fixture("resources.mpp"));
        long revision = document.Revision;
        var assignment = document.Assignments.Single(a => a.Resource?.Type == ProjectResourceType.Cost);
        var analysis = document.AnalyzeAssignments();
        var estimate = analysis.Assignments.Single(a => a.AssignmentUid == assignment.Uid);
        Assert.Equal(300m, estimate.EstimatedCost);
        Assert.Equal(assignment.Cost, estimate.StoredCost);
        Assert.Equal(revision, document.Revision);
        Assert.All(document.Assignments.Where(a => a.Resource?.Type == ProjectResourceType.Work || a.Resource?.Type == ProjectResourceType.Material),
            a => Assert.Null(analysis.Assignments.Single(e => e.AssignmentUid == a.Uid).EstimatedCost));
    }
    [Fact]
    public void UniformWorkEquationsHoldTheRequestedQuantityFixed() {
        var units = ProjectUnits.Percent(50);
        Assert.Equal(480, ProjectWorkEquation.Work(960, units).Minutes);
        Assert.Equal(960, ProjectWorkEquation.DurationMinutes(ProjectWork.Hours(8), units));
        Assert.Equal(units, ProjectWorkEquation.Units(ProjectWork.Hours(8), 960));
        Assert.Throws<ArgumentOutOfRangeException>(() => ProjectWorkEquation.DurationMinutes(ProjectWork.Hours(8), ProjectUnits.Percent(0)));
        Assert.Throws<ArgumentOutOfRangeException>(() => ProjectWorkEquation.Units(ProjectWork.Hours(8), 0));
        Assert.Equal(1125.75m, ProjectWorkEquation.WorkCost(ProjectWork.Hours(10), ProjectWork.Hours(2), 100, 150, 25.75m));
        Assert.Equal(35, ProjectWorkEquation.MaterialCost(3, 10, 5));
        Assert.Throws<ArgumentException>(() => ProjectWorkEquation.WorkCost(ProjectWork.Hours(1), ProjectWork.Hours(2), 100, 150));
    }
    [Fact]
    public void StoredResourceCachesAndAssignmentTotalsRemainSeparateObservations() {
        using var native = ProjectDocument.Load(ProjectNativeTests.Fixture("actuals.mpp"));
        using var reference = ProjectDocument.Load(ProjectNativeTests.Fixture("actuals.xml"));
        long revision = native.Revision;
        var result = native.AnalyzeAssignments();
        var totals = result.Resources.Single(r => r.ResourceUid == 1);
        Assert.Equal(5000m, totals.StoredResourceCost); Assert.Equal(0m, totals.StoredResourceActualCost);
        Assert.Equal(reference.Resources.GetByUid(1).Cost, totals.Cost);
        Assert.Equal(reference.Resources.GetByUid(1).ActualCost, totals.ActualCost);
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "PROJECT_RESOURCE_CACHED_TOTAL");
        Assert.Equal(revision, native.Revision); Assert.Equal(5000m, native.Resources.GetByUid(1).Cost);
    }
    [Fact]
    public void MissingAmountsStayUnknownAndCostResourcesDoNotInventWork() {
        using var document = ProjectDocument.Create();
        var task = document.Tasks.Add("Task"); task.Duration = ProjectDuration.WorkingDays(2);
        var work = document.Resources.AddWork("Engineer"); work.StandardRate = 100;
        var workAssignment = document.Assignments.Add(task, work, ProjectUnits.Percent(50));
        var cost = document.Resources.AddCost("Travel");
        var costAssignment = document.Assignments.Add(task, cost); costAssignment.Cost = 300;
        var analysis = document.AnalyzeAssignments();
        Assert.Equal(480, analysis.Assignments.Single(a => a.AssignmentUid == workAssignment.Uid).EstimatedWorkMinutes);
        Assert.Equal(800, analysis.Assignments.Single(a => a.AssignmentUid == workAssignment.Uid).EstimatedCost);
        Assert.Null(analysis.Resources.Single(r => r.ResourceUid == work.Uid).Cost);
        var estimate = analysis.Assignments.Single(a => a.AssignmentUid == costAssignment.Uid);
        Assert.Equal(300, estimate.EstimatedCost); Assert.Null(estimate.EstimatedWorkMinutes);
        Assert.Null(workAssignment.Work); Assert.Null(workAssignment.Cost);
    }
    [Fact]
    public void ActualRemainingAndOvertimeDisagreementsAreDiagnosedWithoutRepair() {
        using var document = ProjectDocument.Create();
        var task = document.Tasks.Add("Task"); var resource = document.Resources.AddWork("Engineer"); resource.StandardRate = 100;
        var assignment = document.Assignments.Add(task, resource);
        assignment.Work = ProjectWork.Hours(8); assignment.ActualWork = ProjectWork.Hours(5); assignment.RemainingWork = ProjectWork.Hours(4);
        assignment.OvertimeWork = ProjectWork.Hours(9); assignment.Cost = 100; assignment.ActualCost = 60; assignment.RemainingCost = 50;
        var result = document.AnalyzeAssignments();
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "PROJECT_WORK_BALANCE");
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "PROJECT_COST_BALANCE");
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "PROJECT_OVERTIME_BALANCE");
        Assert.Null(result.Assignments.Single().EstimatedCost);
        Assert.Equal(8 * 60, assignment.Work.Value.Minutes); Assert.Equal(100, assignment.Cost);
    }
}
