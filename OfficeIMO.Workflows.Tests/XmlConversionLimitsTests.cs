using OfficeIMO.Excel;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.Workflows.Tests;

public sealed class XmlConversionLimitsTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task AsyncProviderCannotChangeCapturedXmlOrLossAcceptance(bool tightXml) {
        string directory = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "officeimo-xml-snapshot-" + Guid.NewGuid().ToString("N"))).FullName;
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var resume = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        try {
            string source = Path.Combine(directory, "source.fodg"), output = Path.Combine(directory, "result.pdf");
            CreateSource(source, ".fodg");
            byte[] bytes = File.ReadAllBytes(source), original = { 1, 2, 3 };
            File.WriteAllBytes(output, original);
            var request = new OfficeWorkflowRequest {
                Operation = OfficeWorkflowOperation.Convert, InputPath = "content://provider/source", OutputPath = output,
                ConversionRouteId = "odg-pdf", ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                Limits = new() { MaximumXmlCharactersInPart = tightXml ? 256 : 1024 * 1024 },
                ConversionOptions = new() { RequireNoLoss = !tightXml },
                InputStream = new("source.fodg", async token => {
                    entered.TrySetResult();
                    await resume.Task.WaitAsync(token);
                    return new MemoryStream(bytes, writable: false);
                })
            };
            Task<OfficeWorkflowResult> running = new OfficeWorkflowRunner().RunAsync(request);
            await entered.Task.WaitAsync(TimeSpan.FromSeconds(15));
            request.Limits.MaximumXmlCharactersInPart = 1024 * 1024;
            request.ConversionOptions.RequireNoLoss = false;
            resume.SetResult();
            var result = await running;
            Assert.False(result.Succeeded);
            Assert.Equal(original, File.ReadAllBytes(output));
            if (!tightXml) Assert.True(Assert.IsType<OfficeWorkflowConversionEvidence>(result.ConversionEvidence).HasLoss);
        } finally { resume.TrySetResult(); Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData(".odg")]
    [InlineData(".fodg")]
    [InlineData(".docx")]
    [InlineData(".xlsx")]
    [InlineData(".pptx")]
    public async Task PerPartXmlBudgetIsEnforcedBeforeReplacingAnExistingDestination(string extension) {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-xml-budget-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string input = Path.Combine(directory, "source" + extension), output = Path.Combine(directory, "result.pdf");
            CreateSource(input, extension);
            byte[] originalSource = File.ReadAllBytes(input), originalOutput = { 1, 2, 3 };
            File.WriteAllBytes(output, originalOutput);
            var request = new OfficeWorkflowRequest {
                Operation = OfficeWorkflowOperation.Convert, InputPath = input, OutputPath = output,
                ConversionRouteId = extension is ".odg" or ".fodg" ? "odg-pdf" : extension[1..] + "-pdf",
                ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                Limits = new() { MaximumXmlCharactersInPart = 256 }
            };
            var rejected = await new OfficeWorkflowRunner().RunAsync(request);
            Assert.False(rejected.Succeeded);
            Assert.Equal(originalOutput, File.ReadAllBytes(output));
            Assert.Equal(originalSource, File.ReadAllBytes(input));

            request.Limits.MaximumXmlCharactersInPart = 1024 * 1024;
            var accepted = await new OfficeWorkflowRunner().RunAsync(request);
            Assert.True(accepted.Succeeded, accepted.Summary);
            Assert.NotEqual(originalOutput, File.ReadAllBytes(output));
            Assert.Equal(originalSource, File.ReadAllBytes(input));
            Assert.Empty(Directory.GetFiles(directory, ".*.tmp"));
        } finally { Directory.Delete(directory, true); }
    }

    private static void CreateSource(string path, string extension) {
        string marker = new('A', 1024);
        switch (extension) {
            case ".odg":
            case ".fodg":
                var drawing = OdgDocument.Create();
                drawing.AddPage("Budget").Shapes.AddTextBox(OdfRect.FromCentimeters(1, 1, 12, 5), marker);
                if (extension == ".fodg") drawing.SaveFlatXml(path); else drawing.Save(path);
                break;
            case ".docx":
                using (var word = WordDocument.Create(path)) { word.AddParagraph(marker); word.Save(); }
                break;
            case ".xlsx":
                using (var excel = ExcelDocument.Create(path)) { excel.AddWorksheet("Budget").CellValue(1, 1, marker); excel.Save(); }
                break;
            case ".pptx":
                using (var presentation = PowerPointPresentation.Create(path)) { presentation.AddSlide().AddTextBox(marker); presentation.Save(); }
                break;
        }
    }
}
