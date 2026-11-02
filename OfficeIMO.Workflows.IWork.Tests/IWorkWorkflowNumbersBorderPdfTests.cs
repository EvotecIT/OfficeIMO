using System.Security.Cryptography;
using OfficeIMO.Workflows.IWork;
using Xunit;

namespace OfficeIMO.Workflows.IWork.Tests;

public sealed partial class IWorkWorkflowTests {
    [Fact]
    public async Task Native_Numbers_PDF_rejects_padding_loss_and_explicit_partial_sources() {
        using var files = new Files("numbers", "pdf");
        File.Copy(Path.Combine(AppContext.BaseDirectory, "Corpus", "numbers-table-borders-v14.5.numbers"), files.Input, overwrite: true);
        string hash = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(files.Input)));
        OfficeWorkflowRequest request = files.Request("numbers-pdf");
        var settings = NativePdfSettings();
        request.RegisteredConversionSettings = settings;
        var runner = IWorkWorkflow.CreateRunner();
        var complete = await runner.RunAsync(request);
        Assert.False(complete.Succeeded);
        Assert.Contains("padding", complete.Summary, StringComparison.OrdinalIgnoreCase);
        Assert.False(File.Exists(files.Output));

        settings.ConversionOptions.AllowPartialEditableReconstruction = true;
        var result = await runner.RunAsync(request);
        Assert.False(result.Succeeded);
        Assert.Contains("complete editable reconstruction", result.Summary);
        Assert.False(File.Exists(files.Output));
        Assert.Equal(hash, Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(files.Input))));
        Assert.Empty(Directory.GetFiles(files.Root, "*.tmp"));
    }

    [Fact]
    public async Task Native_Numbers_border_XLSX_reports_padding_loss_without_losing_source_rows() {
        using var files = new Files("numbers", "xlsx");
        File.Copy(Path.Combine(AppContext.BaseDirectory, "Corpus", "numbers-table-borders-v14.5.numbers"), files.Input, overwrite: true);
        string hash = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(files.Input)));
        var result = await IWorkWorkflow.CreateRunner().RunAsync(files.Request("numbers-xlsx"));
        Assert.True(result.Succeeded, result.Summary);
        var evidence = Assert.IsType<OfficeWorkflowConversionEvidence>(result.ConversionEvidence);
        Assert.Equal("True", evidence.Facts["partialEditableReconstruction"]);
        Assert.Equal("3", evidence.Facts["sourceUnitCount"]);
        Assert.Equal("3", evidence.Facts["reconstructedSourceUnitCount"]);
        Assert.Equal("0", evidence.Facts["omittedSourceUnitCount"]);
        Assert.Equal("0", evidence.Facts["unassessedSourceUnitCount"]);
        Assert.Contains(evidence.FidelityDiagnostics, d => d.Code == "IWORK_NUMBERS_CELL_PADDING_OMITTED"
            && d.LossKind == OfficeConversionLossKind.Omission);
        Assert.DoesNotContain(evidence.FidelityDiagnostics, d => d.Code == "IWORK_TABLE_BORDERS_UNSUPPORTED"
            || d.LossKind == OfficeConversionLossKind.Failure);
        Assert.Contains(result.Diagnostics, d => d.Code == "conversion-output-reopened");
        Assert.Equal(hash, Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(files.Input))));
        Assert.Empty(Directory.GetFiles(files.Root, "*.tmp"));
    }
}
