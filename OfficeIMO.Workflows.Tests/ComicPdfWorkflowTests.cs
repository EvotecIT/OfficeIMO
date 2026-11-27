using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows.Tests;

public sealed class ComicPdfWorkflowTests {
    [Fact]
    public async Task ComicSevenZipConversionKeepsEachNativePageCanvasInPdf() {
        string source = Path.Combine(AppContext.BaseDirectory, "Fixtures", "harbour.cb7");
        DirectoryInfo directory = Directory.CreateTempSubdirectory("officeimo-comic-route-");
        try {
            string output = Path.Combine(directory.FullName, "comic.pdf");
            OfficeWorkflowResult result = await new OfficeWorkflowRunner().RunAsync(new OfficeWorkflowRequest {
                Operation = OfficeWorkflowOperation.Convert, InputPath = source, OutputPath = output,
                ConversionRouteId = "ebook-cb7-pdf"
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
}
