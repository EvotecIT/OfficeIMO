using OfficeIMO.OpenDocument;

namespace OfficeIMO.Workflows.Tests;

public sealed class DrawBatchWorkflowTests {
    [Fact]
    public async Task CheckpointReuseRetainsLocatedLossesAndRejectsChangedXmlAcceptance() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-draw-batch-" + Guid.NewGuid().ToString("N"));
        string input = Directory.CreateDirectory(Path.Combine(root, "input")).FullName;
        try {
            CreateDraw(Path.Combine(input, "drawing.fodg"));
            File.WriteAllText(Path.Combine(input, "text.txt"), "Literal batch marker");
            var request = new OfficeConversionBatchRequest {
                InputDirectory = input, OutputDirectory = Path.Combine(root, "output"),
                CheckpointDirectory = Path.Combine(root, "state"), MaximumConcurrency = 1
            };
            var firstItems = new List<OfficeConversionBatchItemResult>();
            var first = await new OfficeWorkflowRunner().RunBatchAsync(request, new InlineProgress(firstItems));
            Assert.Equal(2, first.Completed);
            var original = Assert.Single(firstItems, item => item.InputPath.EndsWith(".fodg", StringComparison.Ordinal));
            var located = original.Diagnostics.Where(d => d.Severity == OfficeWorkflowDiagnosticSeverity.Warning && d.Details.ContainsKey("location")).ToArray();
            Assert.Contains(located, d => d.Details["location"].StartsWith("page:1:", StringComparison.Ordinal));

            var reusedItems = new List<OfficeConversionBatchItemResult>();
            var reused = await new OfficeWorkflowRunner().RunBatchAsync(request, new InlineProgress(reusedItems));
            Assert.Equal(2, reused.Reused);
            var restored = Assert.Single(reusedItems, item => item.InputPath.EndsWith(".fodg", StringComparison.Ordinal));
            foreach (var finding in located)
                Assert.Contains(restored.Diagnostics, d => d.Code == finding.Code && d.Details["location"] == finding.Details["location"]
                    && d.Details["source"] == finding.Details["source"] && d.Details["lossKind"] == finding.Details["lossKind"]);

            byte[] originalOutput = File.ReadAllBytes(Path.Combine(request.OutputDirectory, "drawing.fodg.pdf"));
            request.MaximumXmlCharactersInPart = 256;
            var changed = await new OfficeWorkflowRunner().RunBatchAsync(request);
            Assert.Equal(1, changed.Failed);
            Assert.Equal(1, changed.Reused);
            Assert.Equal(originalOutput, File.ReadAllBytes(Path.Combine(request.OutputDirectory, "drawing.fodg.pdf")));
        } finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task OrdinaryMixedBatchEnforcesDiagramAcceptanceWithoutApplyingItToLiteralText(bool strictLoss) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-draw-batch-" + Guid.NewGuid().ToString("N"));
        string input = Directory.CreateDirectory(Path.Combine(root, "input")).FullName;
        try {
            CreateDraw(Path.Combine(input, "drawing.fodg"));
            File.WriteAllText(Path.Combine(input, "text.txt"), "Literal batch marker");
            var request = new OfficeConversionBatchRequest {
                InputDirectory = input, OutputDirectory = Path.Combine(root, "output"), MaximumConcurrency = 1,
                MaximumXmlCharactersInPart = strictLoss ? 1024 * 1024 : 256,
                ConversionOptions = new() { RequireNoLoss = strictLoss }
            };
            var result = await new OfficeWorkflowRunner().RunBatchAsync(request);
            Assert.Equal(1, result.Completed);
            Assert.Equal(1, result.Failed);
            Assert.True(File.Exists(Path.Combine(request.OutputDirectory, "text.txt.pdf")));
            Assert.False(File.Exists(Path.Combine(request.OutputDirectory, "drawing.fodg.pdf")));
        } finally { Directory.Delete(root, true); }
    }

    private static void CreateDraw(string path) {
        var drawing = OdgDocument.Create();
        drawing.AddPage("Native capitalization").Shapes.AddTextBox(OdfRect.FromCentimeters(1, 1, 12, 5), string.Empty)
            .AddParagraph("native capitalization").TextTransform = OdfTextTransform.Capitalize;
        drawing.SaveFlatXml(path);
    }

    private sealed class InlineProgress(List<OfficeConversionBatchItemResult> items) : IProgress<OfficeConversionBatchItemResult> {
        public void Report(OfficeConversionBatchItemResult item) => items.Add(item);
    }
}
