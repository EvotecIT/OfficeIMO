using OfficeIMO.Pdf;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Pdf;

namespace OfficeIMO.Workflows.Tests;

public sealed class VisioPdfWorkflowTests {
    [Theory]
    [InlineData(".vsdx")]
    [InlineData(".vdx")]
    [InlineData(".vtx")]
    public async Task DiagramRoutesPublishPhysicalPagesAndRetainEveryConversionStage(string extension) {
        string root = CreateDirectory();
        try {
            string input = Path.Combine(root, "source" + extension), output = Path.Combine(root, "result.pdf");
            WriteSource(input);
            byte[] before = File.ReadAllBytes(input);
            var result = await OfficeWorkflow.Convert(input).To(output).RunAsync();
            Assert.True(result.Succeeded, result.Summary);
            Assert.Equal("visio-pdf", OfficeWorkflowCatalog.Find(extension, ".pdf")?.Id);
            var pdf = PdfReadDocument.Open(File.ReadAllBytes(output));
            Assert.Equal(2, pdf.Pages.Count);
            Assert.Equal((360D, 216D), pdf.Pages[0].GetPageSize());
            Assert.Equal((144D, 108D), pdf.Pages[1].GetPageSize());
            Assert.Contains("WorkflowCaption", pdf.ExtractText());
            var evidence = Assert.IsType<OfficeWorkflowConversionEvidence>(result.ConversionEvidence);
            Assert.Equal("2", evidence.Facts["sourcePages"]);
            Assert.Equal(extension[1..].ToUpperInvariant(), evidence.Facts["sourceFormat"]);
            Assert.Contains(result.Diagnostics, d => d.Code == "OutputReopened");
            if (extension != ".vsdx")
                Assert.Contains(evidence.FidelityDiagnostics, d => d.Code == "VDX_MODEL_PROFILE" && d.LossKind == OfficeConversionLossKind.Approximation);
            Assert.Equal(before, File.ReadAllBytes(input));
        } finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData(".vsdx", false)]
    [InlineData(".vsdx", true)]
    [InlineData(".vdx", false)]
    [InlineData(".vdx", true)]
    [InlineData(".vtx", false)]
    [InlineData(".vtx", true)]
    public async Task StrictSourceOrCombinedAcceptanceRetainsLossesAndPreservesDestination(string extension, bool sourceStrict) {
        string root = CreateDirectory();
        try {
            string input = Path.Combine(root, "source" + extension), output = Path.Combine(root, "result.pdf");
            WriteSource(input, hyperlink: true);
            byte[] before = File.ReadAllBytes(input), sentinel = [1, 2, 3];
            File.WriteAllBytes(output, sentinel);
            var result = await new OfficeWorkflowRunner().RunAsync(new() {
                Operation = OfficeWorkflowOperation.Convert, InputPath = input, OutputPath = output,
                ConversionRouteId = "visio-pdf", ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                ConversionOptions = new() {
                    RequireNoLoss = !sourceStrict,
                    Visio = new() { Mode = VisioPdfProjectionMode.DiagramPages,
                        DrawingOptions = new() { RequireNoLoss = sourceStrict } }
                }
            });
            Assert.False(result.Succeeded);
            var evidence = Assert.IsType<OfficeWorkflowConversionEvidence>(result.ConversionEvidence);
            Assert.True(evidence.HasLoss);
            Assert.Contains(evidence.FidelityDiagnostics, d => d.Code == "VISIO_DRAWING_METADATA" && d.Location!.StartsWith("page:1:First:shape:"));
            if (extension != ".vsdx") Assert.Contains(evidence.FidelityDiagnostics, d => d.Code == "VDX_MODEL_PROFILE");
            Assert.Equal(before, File.ReadAllBytes(input));
            Assert.Equal(sentinel, File.ReadAllBytes(output));
            Assert.Empty(Directory.GetFiles(root, ".*.tmp"));
        } finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData(".vsdx", "xml")]
    [InlineData(".vdx", "xml")]
    [InlineData(".vdx", "input")]
    [InlineData(".vdx", "output")]
    [InlineData(".vdx", "pages")]
    public async Task InputXmlProjectionAndOutputLimitsPreventPublication(string extension, string limit) {
        string root = CreateDirectory();
        try {
            string input = Path.Combine(root, "source" + extension), output = Path.Combine(root, "result.pdf");
            WriteSource(input);
            byte[] sentinel = [3, 2, 1]; File.WriteAllBytes(output, sentinel);
            var result = await new OfficeWorkflowRunner().RunAsync(new() {
                Operation = OfficeWorkflowOperation.Convert, InputPath = input, OutputPath = output,
                ConversionRouteId = "visio-pdf",
                ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                Limits = new() {
                    MaximumInputBytes = limit == "input" ? 16 : 1024 * 1024,
                    MaximumOutputBytes = limit == "output" ? 64 : 1024 * 1024,
                    MaximumXmlCharactersInPart = limit == "xml" ? 64 : 1024 * 1024
                },
                ConversionOptions = new() { Visio = new() { Mode = VisioPdfProjectionMode.DiagramPages,
                    DrawingOptions = new() { MaximumPages = limit == "pages" ? 1 : 2 } } }
            });
            Assert.False(result.Succeeded);
            Assert.Equal(sentinel, File.ReadAllBytes(output));
            if (limit == "output") {
                Assert.True(result.ConversionEvidence != null, result.Summary);
                var evidence = Assert.IsType<OfficeWorkflowConversionEvidence>(result.ConversionEvidence);
                Assert.Equal("2", evidence.Facts["sourcePages"]);
            }
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void DiagramSettingsAreDetachedAndMixedBatchSelectionRemainsRouteSpecific() {
        var options = new OfficeWorkflowConversionOptions { RequireNoLoss = true,
            Visio = new() { Mode = VisioPdfProjectionMode.DiagramPages,
                DrawingOptions = new() { MaximumPages = 3 }, PdfOptions = new() } };
        byte[] font = OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFont(['G']);
        options.Visio.DrawingOptions.Fonts.Add("Snapshot font", font);
        var copy = options.ForRoute("visio-pdf");
        options.Visio.DrawingOptions.MaximumPages = 30;
        options.Visio.DrawingOptions.Fonts.Add("Later font", font);
        Assert.Equal(3, copy.Visio!.DrawingOptions!.MaximumPages);
        Assert.Single(copy.Visio.DrawingOptions.Fonts.Faces);
        Assert.NotSame(options.Visio.PdfOptions, copy.Visio.PdfOptions);
        Assert.True(copy.RequireNoLoss);
        Assert.Null(options.ForRoute("txt-pdf").Visio);
        Assert.False(options.ForRoute("txt-pdf").RequireNoLoss);
        Assert.Throws<ArgumentException>(() => new OfficeWorkflowConversionOptions { Visio = new() }
            .Snapshot(OfficeWorkflowCatalog.FindExecutable("visio-pdf")!));
    }

    [Fact]
    public async Task CheckpointsRetainLegacyLossesAndInvalidateChangedXmlAcceptanceInMixedBatches() {
        string root = CreateDirectory();
        string input = Directory.CreateDirectory(Path.Combine(root, "input")).FullName;
        try {
            WriteSource(Path.Combine(input, "source.vdx"), hyperlink: true);
            File.WriteAllText(Path.Combine(input, "literal.txt"), "Literal marker");
            var request = new OfficeConversionBatchRequest {
                InputDirectory = input, OutputDirectory = Path.Combine(root, "output"),
                CheckpointDirectory = Path.Combine(root, "state"), MaximumConcurrency = 1,
                ConversionOptions = new() { Visio = new() { Mode = VisioPdfProjectionMode.DiagramPages,
                    DrawingOptions = new() { MaximumPages = 3 } } }
            };
            var firstItems = new List<OfficeConversionBatchItemResult>();
            var first = await new OfficeWorkflowRunner().RunBatchAsync(request, new InlineProgress(firstItems));
            Assert.Equal(2, first.Completed);
            var original = Assert.Single(firstItems, item => item.InputPath.EndsWith(".vdx", StringComparison.Ordinal));
            Assert.Contains(original.Diagnostics, d => d.Code == "VDX_MODEL_PROFILE");
            var reusedItems = new List<OfficeConversionBatchItemResult>();
            var reused = await new OfficeWorkflowRunner().RunBatchAsync(request, new InlineProgress(reusedItems));
            Assert.Equal(2, reused.Reused);
            var restored = Assert.Single(reusedItems, item => item.InputPath.EndsWith(".vdx", StringComparison.Ordinal));
            Assert.Contains(restored.Diagnostics, d => d.Code == "VDX_MODEL_PROFILE");
            byte[] before = File.ReadAllBytes(Path.Combine(request.OutputDirectory, "source.vdx.pdf"));
            request.MaximumXmlCharactersInPart = 64;
            var changed = await new OfficeWorkflowRunner().RunBatchAsync(request);
            Assert.Equal(1, changed.Failed);
            Assert.Equal(1, changed.Reused);
            Assert.Equal(before, File.ReadAllBytes(Path.Combine(request.OutputDirectory, "source.vdx.pdf")));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task RuntimeFontCallbacksAreAcceptedInOrdinaryBatchesAndRejectedForCheckpoints() {
        string root = CreateDirectory();
        try {
            string input = Path.Combine(root, "source.vdx"); WriteSource(input);
            var request = new OfficeConversionBatchRequest {
                InputPaths = [input], OutputDirectory = Path.Combine(root, "output"),
                CheckpointDirectory = Path.Combine(root, "state"), MaximumConcurrency = 1,
                ConversionOptions = new() { Visio = new() { Mode = VisioPdfProjectionMode.DiagramPages, DrawingOptions = new() } }
            };
            request.ConversionOptions.Visio!.DrawingOptions!.Fonts.FontVariationResolver = _ => new Dictionary<string, float>();
            var rejectedItems = new List<OfficeConversionBatchItemResult>();
            var rejected = await new OfficeWorkflowRunner().RunBatchAsync(request, new InlineProgress(rejectedItems));
            Assert.Equal(1, rejected.Failed);
            Assert.Contains("ordinary batch without checkpoints", Assert.Single(rejectedItems).Summary, StringComparison.Ordinal);
            Assert.False(File.Exists(Path.Combine(request.OutputDirectory, "source.vdx.pdf")));
            request.CheckpointDirectory = null;
            var ordinary = await new OfficeWorkflowRunner().RunBatchAsync(request);
            Assert.Equal(1, ordinary.Completed);
            Assert.True(File.Exists(Path.Combine(request.OutputDirectory, "source.vdx.pdf")));
        } finally { Directory.Delete(root, true); }
    }

    private static string CreateDirectory() => Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "officeimo-visio-workflow-" + Guid.NewGuid().ToString("N"))).FullName;

    private static void WriteSource(string path, bool hyperlink = false) {
        VisioDocument source = VisioDocument.Create(Path.GetExtension(path) == ".vtx" ? VisioPackageType.Template : VisioPackageType.Drawing);
        VisioPage first = source.AddPage("First"); first.Width = 5; first.Height = 3;
        VisioShape shape = first.AddRectangle(1, 1, 2, 1, "WorkflowCaption");
        if (hyperlink) shape.AddHyperlink("https://example.com");
        VisioPage second = source.AddPage("Blank"); second.Width = 2; second.Height = 1.5;
        if (Path.GetExtension(path) == ".vsdx") source.Save(path);
        else {
            // Generated modern PageSheet/theme cells have no legacy XML mapping.
            // This controlled fixture accepts those exporter omissions before exercising import/conversion.
            source.SaveLegacyXml(path, allowOmissions: true);
        }
    }

    private sealed class InlineProgress(List<OfficeConversionBatchItemResult> items) : IProgress<OfficeConversionBatchItemResult> {
        public void Report(OfficeConversionBatchItemResult item) => items.Add(item);
    }
}
