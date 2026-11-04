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
    public void LocalOfficeExportsUseIsolatedBalancedFontPolicies(string routeId) {
        using var workbench = new ConversionWorkbenchViewModel(
            _ => Task.FromResult<IReadOnlyList<string>>([]), _ => Task.FromResult<string?>(null));
        var job = new ConversionJobViewModel("local-document", workbench.Routes.Single(route => route.Route.Id == routeId));
        OfficeConversionOptions conversion = job.CreateConversionOptions();
        PdfResourcePolicy policy = conversion.Word?.ResourcePolicy ?? conversion.Excel?.ResourcePolicy ?? conversion.PowerPoint!.ResourcePolicy;
        Assert.True(policy.AllowSystemFontEmbedding);
        Assert.True(policy.AllowDocumentFontEmbedding);
        Assert.False(policy.AllowLocalFileAccess);
        Assert.False(policy.AllowRemoteResourceResolution);
        // Mutating one job must not change another job or the shared default.
        policy.AllowDocumentFontEmbedding = false;
        Assert.True(PdfResourcePolicy.CreateDefault().AllowDocumentFontEmbedding);
        OfficeConversionOptions another = job.CreateConversionOptions();
        Assert.True((another.Word?.ResourcePolicy ?? another.Excel?.ResourcePolicy ?? another.PowerPoint!.ResourcePolicy).AllowDocumentFontEmbedding);
    }

    [Theory]
    [InlineData(OfficeConversionProfile.Faithful, false)]
    [InlineData(OfficeConversionProfile.Lightweight, false)]
    [InlineData(OfficeConversionProfile.PrintReady, true)]
    [InlineData(OfficeConversionProfile.TextOnly, false)]
    public void FontSelectionPreservesTheSelectedWordExportProfile(OfficeConversionProfile profile, bool borders) {
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
