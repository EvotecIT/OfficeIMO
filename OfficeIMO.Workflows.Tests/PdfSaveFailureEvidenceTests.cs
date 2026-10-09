using OfficeIMO.Excel;
using OfficeIMO.OpenDocument;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.Workflows.Tests;

public sealed class PdfSaveFailureEvidenceTests {
    [Theory]
    [InlineData(".fodg", "odg-pdf")]
    [InlineData(".docx", "docx-pdf")]
    [InlineData(".xlsx", "xlsx-pdf")]
    [InlineData(".pptx", "pptx-pdf")]
    public async Task PdfEncodingFailureRetainsStructuredEvidenceAndPreservesFiles(string extension, string route) {
        string root = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "officeimo-pdf-failure-evidence-" + Guid.NewGuid().ToString("N"))).FullName;
        try {
            string input = Path.Combine(root, "source" + extension), output = Path.Combine(root, "result.pdf");
            const string text = "مرحبا";
            if (extension == ".docx") { using var word = WordDocument.Create(input); word.AddParagraph(text); word.Save(); }
            else if (extension == ".xlsx") { using var excel = ExcelDocument.Create(input); excel.AddWorksheet("Unicode").CellValue(1, 1, text); excel.Save(); }
            else if (extension == ".pptx") { using var ppt = PowerPointPresentation.Create(input); ppt.AddSlide().AddTextBox(text); ppt.Save(); }
            else {
                var drawing = OdgDocument.Create();
                drawing.AddPage("Unicode").Shapes.AddTextBox(OdfRect.FromCentimeters(1, 1, 12, 5), text);
                drawing.SaveFlatXml(input);
            }
            byte[] source = File.ReadAllBytes(input), destination = { 1, 2, 3 };
            File.WriteAllBytes(output, destination);
            // Portable hosts can deliberately disable font discovery and generated fallbacks.
            var options = new OfficeWorkflowConversionOptions();
            if (extension == ".docx") options.Word = new() { TextFallbacks = PdfTextFallbackFeatures.None, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() };
            else if (extension == ".xlsx") options.Excel = new() { TextFallbacks = PdfTextFallbackFeatures.None, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() };
            else if (extension == ".pptx") options.PowerPoint = new() { TextFallbacks = PdfTextFallbackFeatures.None, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() };
            var result = await new OfficeWorkflowRunner().RunAsync(new() {
                Operation = OfficeWorkflowOperation.Convert, InputPath = input, OutputPath = output,
                ConversionRouteId = route, ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                ConversionOptions = options
            });
            Assert.False(result.Succeeded);
            var evidence = Assert.IsType<OfficeWorkflowConversionEvidence>(result.ConversionEvidence);
            Assert.Contains(evidence.FidelityDiagnostics, finding => finding.Code == "unsupported-text-glyph");
            Assert.Contains(result.Diagnostics, finding => finding.Code == "unsupported-text-glyph");
            Assert.Equal(source, File.ReadAllBytes(input));
            Assert.Equal(destination, File.ReadAllBytes(output));
            Assert.Empty(Directory.GetFiles(root, ".*.tmp"));
        } finally { Directory.Delete(root, true); }
    }
}
