using OfficeIMO.DjVu.Pdf;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows.Tests;

public sealed class DjVuPdfWorkflowTests {
    [Theory]
    [InlineData(".djvu")]
    [InlineData(".djv")]
    public async Task SharedRoutePublishesSourcePagesAndSearchableTextWithEvidence(string extension) {
        string root = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "OfficeIMO-DjVu-" + Guid.NewGuid().ToString("N"))).FullName;
        try {
            string input = Path.Combine(root, "source" + extension), output = Path.Combine(root, "result.pdf");
            byte[] source = File.ReadAllBytes(Fixture("reader-book.djvu")); File.WriteAllBytes(input, source);
            Assert.Equal("djvu-pdf", OfficeWorkflowCatalog.Find(extension, ".pdf")?.Id);
            var result = await OfficeWorkflow.Convert(input).To(output).RunAsync();
            Assert.True(result.Succeeded, result.Summary);
            var pdf = PdfReadDocument.Open(File.ReadAllBytes(output));
            Assert.Equal(4, pdf.Pages.Count);
            Assert.Contains("Zażółć", pdf.ExtractText());
            Assert.Equal((23.04D, 30.72D), pdf.Pages[1].GetPageSize());
            var evidence = Assert.IsType<OfficeWorkflowConversionEvidence>(result.ConversionEvidence);
            Assert.Equal("4", evidence.Facts["sourcePages"]);
            Assert.Contains(evidence.FidelityDiagnostics, d => d.Code == "djvu.pdf.stored-text-corrupt");
            Assert.Equal(source, File.ReadAllBytes(input));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task LosslessAcceptanceAndLimitsPreserveExistingDestination() {
        string root = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "OfficeIMO-DjVu-" + Guid.NewGuid().ToString("N"))).FullName;
        try {
            string input = Path.Combine(root, "small.djvu"), output = Path.Combine(root, "result.pdf");
            File.Copy(Fixture("noise-small.djvu"), input);
            byte[] sentinel = [1, 2, 3]; File.WriteAllBytes(output, sentinel);
            var request = new OfficeWorkflowRequest { Operation = OfficeWorkflowOperation.Convert, InputPath = input, OutputPath = output,
                ConversionRouteId = "djvu-pdf", ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                ConversionOptions = new() { RequireNoLoss = true, DjVu = new DjVuToPdfOptions() } };
            var result = await new OfficeWorkflowRunner().RunAsync(request);
            Assert.False(result.Succeeded);
            Assert.Contains(result.ConversionEvidence!.FidelityDiagnostics, d => d.Code == "djvu.render.short-iw44-edge");
            Assert.Equal(sentinel, File.ReadAllBytes(output));
            request.ConversionOptions.RequireNoLoss = false;
            request.ConversionOptions.DjVu!.MaxPdfBytes = 128;
            result = await new OfficeWorkflowRunner().RunAsync(request);
            Assert.False(result.Succeeded);
            Assert.Equal(sentinel, File.ReadAllBytes(output));
        } finally { Directory.Delete(root, true); }
    }

    private static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "Fixtures", "DjVu", name);
}
