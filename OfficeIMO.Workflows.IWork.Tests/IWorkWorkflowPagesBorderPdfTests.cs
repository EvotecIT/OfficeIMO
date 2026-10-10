using System.Security.Cryptography;
using OfficeIMO.Workflows.IWork;
using Xunit;

namespace OfficeIMO.Workflows.IWork.Tests;

public sealed partial class IWorkWorkflowTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Native_Pages_border_PDF_rejects_unqualified_positioning(bool allowPartial) {
        using var files = new Files("pages", "pdf");
        File.Copy(Path.Combine(AppContext.BaseDirectory, "Corpus", "pages-table-borders-v15.4.pages"), files.Input, overwrite: true);
        string sourceHash = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(files.Input)));
        var request = files.Request("pages-pdf");
        var settings = NativePdfSettings();
        settings.ConversionOptions.AllowPartialEditableReconstruction = allowPartial;
        request.RegisteredConversionSettings = settings;

        var result = await IWorkWorkflow.CreateRunner().RunAsync(request);

        Assert.False(result.Succeeded);
        Assert.Contains(allowPartial ? "complete editable reconstruction" : "DOCX", result.Summary);
        Assert.False(File.Exists(files.Output));
        Assert.Equal(sourceHash, Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(files.Input))));
        Assert.Empty(Directory.GetFiles(files.Root, "*.tmp"));
    }

    [Fact]
    public async Task Native_Pages_border_DOCX_retains_source_units_and_reports_placement_limit() {
        using var files = new Files("pages", "docx");
        File.Copy(Path.Combine(AppContext.BaseDirectory, "Corpus", "pages-table-borders-v15.4.pages"), files.Input, overwrite: true);
        string sourceHash = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(files.Input)));
        var result = await IWorkWorkflow.CreateRunner().RunAsync(files.Request("pages-docx"));

        Assert.True(result.Succeeded, result.Summary);
        var evidence = Assert.IsType<OfficeWorkflowConversionEvidence>(result.ConversionEvidence);
        Assert.Equal("True", evidence.Facts["partialEditableReconstruction"]);
        Assert.NotEqual("0", evidence.Facts["sourceUnitCount"]);
        Assert.Equal(evidence.Facts["sourceUnitCount"], evidence.Facts["reconstructedSourceUnitCount"]);
        Assert.Equal("0", evidence.Facts["omittedSourceUnitCount"]);
        Assert.Equal("0", evidence.Facts["unassessedSourceUnitCount"]);
        Assert.DoesNotContain(evidence.FidelityDiagnostics, d => d.Code == "IWORK_PAGES_DOCX_PRECISION");
        Assert.Contains(evidence.FidelityDiagnostics, d => d.Code == "IWORK_PAGES_TABLE_LAYOUT_APPROXIMATED");
        Assert.DoesNotContain(evidence.FidelityDiagnostics, d => d.Code == "IWORK_TABLE_BORDERS_UNSUPPORTED"
            || d.LossKind == OfficeConversionLossKind.Failure);
        Assert.Contains(result.Diagnostics, d => d.Code == "OutputReopened");
        Assert.Equal(sourceHash, Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(files.Input))));
        Assert.Empty(Directory.GetFiles(files.Root, "*.tmp"));
    }
}
