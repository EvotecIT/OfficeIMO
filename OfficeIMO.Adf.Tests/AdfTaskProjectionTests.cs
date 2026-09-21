using System;
using OfficeIMO.Adf;
using Xunit;

namespace OfficeIMO.Adf.Tests;

public sealed class AdfTaskProjectionTests {
    [Fact]
    public void TaskProjection_ReportsRegeneratedLocalIds() {
        var item = new AdfNode("taskItem") {
            Content = { AdfNode.TextNode("Ready") }
        }.SetAttribute("localId", "item-1").SetAttribute("state", "DONE");
        var list = new AdfNode("taskList") {
            Content = { item }
        }.SetAttribute("localId", "list-1");
        var document = new AdfDocument(new[] { list });

        Assert.True(document.Validate().IsValid);

        AdfConversionResult<string> markdown = AdfConverter.ToMarkdown(document);
        AdfConversionResult<AdfDocument> roundTrip = AdfConverter.FromMarkdown(markdown.Value);

        Assert.Equal("- [x] Ready", markdown.Value.Replace("\r\n", "\n"));
        Assert.False(markdown.Report.IsLossless);
        AdfConversionDiagnostic diagnostic = Assert.Single(
            markdown.Report.Diagnostics,
            item => item.Code == "ADF_TASK_LOCAL_IDS_REGENERATED");
        Assert.Equal("$.content[0]", diagnostic.Path);
        OfficeConversionFidelityDiagnostic fidelity = Assert.Single(markdown.Report.FidelityDiagnostics,
            item => item.Code == diagnostic.Code);
        Assert.Equal(OfficeConversionLossKind.Omission, fidelity.LossKind);
        Assert.Throws<InvalidOperationException>(markdown.Report.RequireNoLoss);
        AdfNode projectedList = Assert.Single(roundTrip.Value.Content);
        AdfNode projectedItem = Assert.Single(projectedList.Content);
        Assert.NotEqual("list-1", projectedList.GetStringAttribute("localId"));
        Assert.NotEqual("item-1", projectedItem.GetStringAttribute("localId"));
    }
}
