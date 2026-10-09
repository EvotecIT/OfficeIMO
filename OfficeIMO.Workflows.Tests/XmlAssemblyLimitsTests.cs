using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.Workflows.Tests;

public sealed class XmlAssemblyLimitsTests {
    [Theory]
    [InlineData(".docx")]
    [InlineData(".xlsx")]
    [InlineData(".pptx")]
    public async Task AssemblyNormalizationHonorsTheCapturedXmlBudget(string extension) {
        string root = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "officeimo-assembly-xml-" + Guid.NewGuid().ToString("N"))).FullName;
        try {
            string input = Path.Combine(root, "source" + extension), output = Path.Combine(root, "result.pdf");
            string marker = new('A', 1024);
            if (extension == ".docx") { using var word = WordDocument.Create(input); word.AddParagraph(marker); word.Save(); }
            else if (extension == ".xlsx") { using var excel = ExcelDocument.Create(input); excel.AddWorksheet("Budget").CellValue(1, 1, marker); excel.Save(); }
            else { using var ppt = PowerPointPresentation.Create(input); ppt.AddSlide().AddTextBox(marker); ppt.Save(); }
            byte[] originalSource = File.ReadAllBytes(input), originalOutput = { 1, 2, 3 };
            File.WriteAllBytes(output, originalOutput);
            var request = new PdfAssemblyRequest {
                Sources = [input], OutputPath = output, ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                Limits = new() { MaximumXmlCharactersInPart = 256 }
            };
            var rejected = await new OfficeWorkflowRunner().AssemblePdfAsync(request);
            Assert.False(rejected.Succeeded);
            Assert.Equal(originalOutput, File.ReadAllBytes(output));
            request.Limits.MaximumXmlCharactersInPart = 1024 * 1024;
            var accepted = await new OfficeWorkflowRunner().AssemblePdfAsync(request);
            Assert.True(accepted.Succeeded, accepted.Summary);
            Assert.Equal(originalSource, File.ReadAllBytes(input));
            Assert.Empty(Directory.GetFiles(root, ".*.tmp"));
        } finally { Directory.Delete(root, true); }
    }
}
