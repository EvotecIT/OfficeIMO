using System.Security.Cryptography;
using OfficeIMO.Pdf;
using OfficeIMO.Workflows.IWork;
using Xunit;

namespace OfficeIMO.Workflows.IWork.Tests;

public sealed partial class IWorkWorkflowTests {
    [Fact]
    public async Task Native_Pages_border_PDF_preserves_complete_table_content_and_fixed_layout() {
        using var files = new Files("pages", "pdf");
        File.Copy(Path.Combine(AppContext.BaseDirectory, "Corpus", "pages-table-borders-v15.4.pages"), files.Input, overwrite: true);
        string sourceHash = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(files.Input)));
        var request = files.Request("pages-pdf");
        var settings = NativePdfSettings();
        request.RegisteredConversionSettings = settings;

        var result = await IWorkWorkflow.CreateRunner().RunAsync(request);

        Assert.True(result.Succeeded, result.Summary);
        var evidence = Assert.IsType<OfficeWorkflowConversionEvidence>(result.ConversionEvidence);
        Assert.Equal("iwork-editable,pdf", evidence.Facts["conversionStages"]);
        Assert.Equal("False", evidence.Facts["partialEditableReconstruction"]);
        Assert.Equal("18", evidence.Facts["sourceUnitCount"]);
        Assert.Equal("18", evidence.Facts["reconstructedSourceUnitCount"]);
        Assert.Equal("0", evidence.Facts["omittedSourceUnitCount"]);
        Assert.Equal("0", evidence.Facts["unassessedSourceUnitCount"]);
        Assert.DoesNotContain(evidence.FidelityDiagnostics, d => d.Code == "IWORK_PAGES_TABLE_LAYOUT_APPROXIMATED"
            || d.LossKind is OfficeConversionLossKind.Omission or OfficeConversionLossKind.Failure);
        Assert.Contains(result.Diagnostics, d => d.Code == "OutputReopened");
        var pdf = PdfReadDocument.Open(File.ReadAllBytes(files.Output));
        string text = Assert.Single(pdf.Pages).ExtractText();
        int previous = -1;
        foreach (string value in new[] { "Corner", "One", "Two", "Three", "Row A", "Red edge", "Hidden edge",
            "Body", "Row B", "Green edge", "Body", "Footer", "End", "Last", "Tail" }) {
            int at = text.IndexOf(value, previous + 1, StringComparison.Ordinal);
            Assert.True(at > previous, "Missing or reordered table value: " + value);
            previous = at;
        }
        Assert.Equal(sourceHash, Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(files.Input))));
        Assert.Empty(Directory.GetFiles(files.Root, "*.tmp"));
    }

    [Fact]
    public async Task Native_Pages_border_DOCX_retains_source_units_under_the_default_complete_policy() {
        using var files = new Files("pages", "docx");
        File.Copy(Path.Combine(AppContext.BaseDirectory, "Corpus", "pages-table-borders-v15.4.pages"), files.Input, overwrite: true);
        string sourceHash = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(files.Input)));
        var request = files.Request("pages-docx");
        request.RegisteredConversionSettings = new IWorkWorkflowSettings();
        var result = await IWorkWorkflow.CreateRunner().RunAsync(request);

        Assert.True(result.Succeeded, result.Summary);
        var evidence = Assert.IsType<OfficeWorkflowConversionEvidence>(result.ConversionEvidence);
        Assert.Equal("False", evidence.Facts["partialEditableReconstruction"]);
        Assert.NotEqual("0", evidence.Facts["sourceUnitCount"]);
        Assert.Equal(evidence.Facts["sourceUnitCount"], evidence.Facts["reconstructedSourceUnitCount"]);
        Assert.Equal("0", evidence.Facts["omittedSourceUnitCount"]);
        Assert.Equal("0", evidence.Facts["unassessedSourceUnitCount"]);
        Assert.DoesNotContain(evidence.FidelityDiagnostics, d => d.Code == "IWORK_PAGES_DOCX_PRECISION");
        Assert.DoesNotContain(evidence.FidelityDiagnostics, d => d.Code == "IWORK_PAGES_TABLE_LAYOUT_APPROXIMATED");
        Assert.DoesNotContain(evidence.FidelityDiagnostics, d => d.Code == "IWORK_TABLE_BORDERS_UNSUPPORTED"
            || d.LossKind == OfficeConversionLossKind.Failure);
        Assert.Contains(result.Diagnostics, d => d.Code == "OutputReopened");
        Assert.Equal(sourceHash, Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(files.Input))));
        Assert.Empty(Directory.GetFiles(files.Root, "*.tmp"));
    }
}
