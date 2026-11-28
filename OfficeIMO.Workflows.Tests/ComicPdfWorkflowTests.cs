using OfficeIMO.Archives;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows.Tests;

public sealed class ComicPdfWorkflowTests {
    [Theory]
    [InlineData("cb7")]
    [InlineData("cbr")]
    public async Task ComicConversionKeepsEachNativePageCanvasInPdf(string format) {
        string source = Path.Combine(AppContext.BaseDirectory, "Fixtures", "harbour." + format);
        DirectoryInfo directory = Directory.CreateTempSubdirectory("officeimo-comic-route-");
        try {
            string output = Path.Combine(directory.FullName, "comic.pdf");
            OfficeWorkflowResult result = await new OfficeWorkflowRunner().RunAsync(new OfficeWorkflowRequest {
                Operation = OfficeWorkflowOperation.Convert, InputPath = source, OutputPath = output,
                ConversionRouteId = $"ebook-{format}-pdf"
            });

            Assert.True(result.Succeeded, result.Summary);
            using PdfDocument pdf = PdfDocument.Load(output);
            Assert.Collection(pdf.Pages,
                cover => Assert.Equal((75D, 112.5D), cover.GetPageSize()),
                second => Assert.Equal((90D, 120D), second.GetPageSize()),
                third => Assert.Equal((108D, 135D), third.GetPageSize()));
        } finally {
            directory.Delete(recursive: true);
        }
    }
    [Fact]
    public async Task CompressedRarComicFailsWithTheEntryAndCodecDiagnosticBeforePublishingPdf() {
        string source = Path.Combine(AppContext.BaseDirectory, "Fixtures", "harbour-compressed.cbr");
        DirectoryInfo directory = Directory.CreateTempSubdirectory("officeimo-cbr-compressed-");
        try {
            string output = Path.Combine(directory.FullName, "comic.pdf");
            OfficeWorkflowResult result = await new OfficeWorkflowRunner().RunAsync(new OfficeWorkflowRequest {
                Operation = OfficeWorkflowOperation.Convert, InputPath = source, OutputPath = output,
                ConversionRouteId = "ebook-cbr-pdf"
            });

            Assert.False(result.Succeeded);
            OfficeWorkflowDiagnostic failure = Assert.Single(result.Diagnostics, item => item.Code == "WorkflowFailed");
            Assert.Equal(nameof(Rar5CompressionNotSupportedException), failure.Details["exceptionType"]);
            Assert.Contains("ComicInfo.xml", failure.Message, StringComparison.Ordinal);
            Assert.Contains("method 3", failure.Message, StringComparison.Ordinal);
            Assert.Contains("131072-byte dictionary", failure.Message, StringComparison.Ordinal);
            Assert.False(File.Exists(output));
        } finally {
            directory.Delete(recursive: true);
        }
    }

}
