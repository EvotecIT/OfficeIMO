using OfficeIMO.Pdf;
using OfficeIMO.Word;

namespace OfficeIMO.Workflows.Tests;

public sealed class DocumentPdfWorkflowTests {
    [Theory]
    [InlineData(".doc", "doc-pdf")]
    [InlineData(".docx", "docx-pdf")]
    [InlineData(".txt", "txt-pdf")]
    public async Task DocumentRoutesPublishReopenedPdfWithContent(string extension, string route) {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-document-pdf-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string input = Path.Combine(directory, "source" + extension);
            if (extension == ".txt") File.WriteAllText(input, "# literal archive content");
            else {
                using var word = WordDocument.Create();
                word.AddParagraph("ArchiveContentMarker");
                word.Save(input);
            }
            var result = await OfficeWorkflow.RunAsync((new OfficeWorkflowRequest {
                Operation = OfficeWorkflowOperation.Convert,
                InputPath = input,
                OutputPath = Path.Combine(directory, "output.pdf")
,
                ConversionRouteId = OfficeWorkflowCatalog.Find(Path.GetExtension(input), Path.GetExtension(Path.Combine(directory, "output.pdf")), executableOnly: true)?.Id
            }));
            Assert.True(result.Succeeded, result.Summary);
            Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "conversion-output-reopened");
            Assert.Equal(route, OfficeWorkflowCatalog.Find(extension, ".pdf")?.Id);
            var pdf = PdfDocument.Load(result.OutputPath!);
            Assert.Single(pdf.Read().Pages);
            Assert.True(new FileInfo(result.OutputPath!).Length > 100);
        } finally { Directory.Delete(directory, true); }
    }
}