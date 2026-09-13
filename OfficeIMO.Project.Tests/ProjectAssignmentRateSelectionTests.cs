namespace OfficeIMO.Project.Tests;

public sealed class ProjectAssignmentRateSelectionTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void AlternateDatedRatesDoNotSuppressDefaultScalarEstimates(bool material, bool explicitDefault) {
        using var document = ProjectDocument.Create();
        var task = document.Tasks.Add("Task");
        var resource = material ? document.Resources.AddMaterial("Material") : document.Resources.AddWork("Work");
        resource.StandardRate = 100; resource.CostPerUse = 5;
        var rate = resource.Rates.Add(); rate.Table = ProjectCostRateTable.B; rate.From = new DateTime(2026, 1, 1); rate.To = new DateTime(2027, 1, 1); rate.StandardRate = 200;
        var assignment = document.Assignments.Add(task, resource, ProjectUnits.Fraction(2)); assignment.Work = ProjectWork.Hours(2);
        if (explicitDefault) assignment.CostRateTable = ProjectCostRateTable.A;
        if (material) assignment.HasFixedRateUnits = true;
        long revision = document.Revision;
        var analysis = document.AnalyzeAssignments();
        Assert.Equal(material ? 205m : 210m, Assert.Single(analysis.Assignments).EstimatedCost);
        Assert.DoesNotContain(analysis.Report.Diagnostics, d => d.Code == "PROJECT_RATE_ESTIMATE_UNSUPPORTED");
        Assert.Equal(revision, document.Revision); Assert.Null(assignment.Cost);
        using var copy = document.Clone();
        Assert.Equal(material ? 205m : 210m, Assert.Single(copy.AnalyzeAssignments().Assignments).EstimatedCost);
        assignment.CostRateTable = ProjectCostRateTable.B;
        Assert.Null(Assert.Single(document.AnalyzeAssignments().Assignments).EstimatedCost);
        assignment.CostRateTable = ProjectCostRateTable.A; rate.Table = null;
        Assert.Null(Assert.Single(document.AnalyzeAssignments().Assignments).EstimatedCost);
    }
}
