using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.Workflows.Tests;

public sealed class PdfAssemblyFailureEvidenceTests {
    [Theory]
    [InlineData(".docx", "docx-pdf")]
    [InlineData(".xlsx", "xlsx-pdf")]
    [InlineData(".pptx", "pptx-pdf")]
    [InlineData(".html", "html-pdf")]
    public async Task AssemblyNormalizationRetainsTheSameSaveFailureDiagnosticsAsConversion(string extension, string route) {
        string root = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "officeimo-assembly-failure-evidence-" + Guid.NewGuid().ToString("N"))).FullName;
        try {
            string input = Path.Combine(root, "source" + extension), output = Path.Combine(root, "result.pdf");
            const string text = "مرحبا";
            if (extension == ".docx") { using var word = WordDocument.Create(input); word.AddParagraph(text); word.Save(); }
            else if (extension == ".xlsx") { using var excel = ExcelDocument.Create(input); excel.AddWorksheet("Unicode").CellValue(1, 1, text); excel.Save(); }
            else if (extension == ".pptx") { using var ppt = PowerPointPresentation.Create(input); ppt.AddSlide().AddTextBox(text); ppt.Save(); }
            else File.WriteAllText(input, "<html><body><p>" + text + "</p></body></html>");
            byte[] source = File.ReadAllBytes(input), destination = { 1, 2, 3 };
            File.WriteAllBytes(output, destination);
            var runner = new OfficeWorkflowRunner();
            var conversion = await runner.RunAsync(new() {
                Operation = OfficeWorkflowOperation.Convert, InputPath = input, OutputPath = output,
                ConversionRouteId = route, ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                Limits = new() { MaximumOutputBytes = 1 }
            });
            Assert.False(conversion.Succeeded);
            var expected = Assert.Single(conversion.Diagnostics, d => d.Code == "WorkflowFailed");
            var assembly = await runner.AssemblePdfAsync(new() {
                Sources = [input], OutputPath = output, ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                Limits = new() { MaximumOutputBytes = 1 }
            });
            Assert.False(assembly.Succeeded);
            Assert.Equal(conversion.FailureKind, assembly.FailureKind);
            Assert.Equal(expected.Details["exceptionType"], Assert.Single(assembly.Diagnostics, d => d.Code == "PdfAssemblyFailed").Details["exceptionType"]);
            var evidence = Assert.IsType<OfficeWorkflowConversionEvidence>(conversion.ConversionEvidence);
            Assert.NotEmpty(evidence.FidelityDiagnostics);
            foreach (var finding in evidence.FidelityDiagnostics) {
                Assert.Contains(assembly.Diagnostics, d => d.Code == finding.Code && d.Message == finding.Message
                    && d.Details.TryGetValue("source", out string? sourceName) && sourceName == finding.Source);
            }
            Assert.Equal(source, File.ReadAllBytes(input));
            Assert.Equal(destination, File.ReadAllBytes(output));
            Assert.Empty(Directory.GetFiles(root, ".*.tmp"));
        } finally { Directory.Delete(root, true); }
    }
}
