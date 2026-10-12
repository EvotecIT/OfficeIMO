using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioLocalDocumentFontTests {
    [Theory]
    [InlineData("docx-pdf")]
    [InlineData("doc-pdf")]
    [InlineData("xlsx-pdf")]
    [InlineData("pptx-pdf")]
    public void LocalOfficeExportsAllowDocumentFontsWithoutOtherHostResources(string routeId) {
        using var workbench = new ConversionWorkbenchViewModel(
            _ => Task.FromResult<IReadOnlyList<string>>([]), _ => Task.FromResult<string?>(null));
        var job = new ConversionJobViewModel("local-document", workbench.Routes.Single(route => route.Route.Id == routeId));
        OfficeWorkflowConversionOptions conversion = job.CreateConversionOptions();
        PdfResourcePolicy policy = conversion.Word?.ResourcePolicy ?? conversion.Excel?.ResourcePolicy ?? conversion.PowerPoint!.ResourcePolicy;
        Assert.True(policy.AllowSystemFontEmbedding);
        Assert.True(policy.AllowDocumentFontEmbedding);
        Assert.False(policy.AllowLocalFileAccess);
        Assert.False(policy.AllowRemoteResourceResolution);
        // Per-job opt-in must not mutate the policy used by other hosts.
        Assert.False(PdfResourcePolicy.CreateDefault().AllowDocumentFontEmbedding);
    }

    [Theory]
    [InlineData(OfficeWorkflowOutputProfile.Faithful, false)]
    [InlineData(OfficeWorkflowOutputProfile.Lightweight, false)]
    [InlineData(OfficeWorkflowOutputProfile.PrintReady, true)]
    [InlineData(OfficeWorkflowOutputProfile.TextOnly, false)]
    public void FontSelectionPreservesTheSelectedWordExportProfile(OfficeWorkflowOutputProfile profile, bool borders) {
        using var workbench = new ConversionWorkbenchViewModel(
            _ => Task.FromResult<IReadOnlyList<string>>([]), _ => Task.FromResult<string?>(null));
        var job = new ConversionJobViewModel("local.docx", workbench.Routes.Single(route => route.Route.Id == "docx-pdf")) {
            OutputProfile = profile
        };
        var options = job.CreateConversionOptions().Word!;
        Assert.Equal(borders, options.DefaultTableBorders);
        Assert.False(options.IncludePageNumbers);
        Assert.True(options.ResourcePolicy.AllowDocumentFontEmbedding);
    }
}
