using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Publisher;
using OfficeIMO.TestAssets;

namespace OfficeIMO.Workflows.Tests;

public sealed class PublisherPdfWorkflowTests {
    [Fact]
    public async Task SerializationCancellationRetainsRecoveryEvidenceAndPreservesTheDestination() {
        string root = CreateDirectory();
        try {
            string input = CopySource(root), output = Path.Combine(root, "result.pdf");
            byte[] sentinel = [3, 2, 1]; File.WriteAllBytes(output, sentinel);
            using var cancellation = new CancellationTokenSource();
            var provider = new CancellingShaper(cancellation);
            byte[] font = ManagedTextShapingTestAssets.CreateFont(Enumerable.Range(32, 95).ToArray());
            var options = new PdfOptions { TextShapingProvider = provider };
            options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Times New Roman", font));
            options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Arial", font));
            var request = OfficeWorkflow.Convert(input).To(output).WithConversionOptions(new() { PublisherPdf = options }).Build();
            request.ConflictPolicy = OfficeWorkflowConflictPolicy.Replace;
            var result = await new OfficeWorkflowRunner().RunAsync(request, cancellationToken: cancellation.Token);
            Assert.True(provider.Called);
            Assert.Equal(OfficeWorkflowStatus.Cancelled, result.Status);
            Assert.Equal(sentinel, File.ReadAllBytes(output));
            Assert.Empty(Directory.GetFiles(root, ".*.tmp"));
            var evidence = Assert.IsType<OfficeWorkflowConversionEvidence>(result.ConversionEvidence);
            Assert.Equal("PUB", evidence.Facts["sourceFormat"]);
            Assert.Equal("2", evidence.Facts["sourcePages"]);
            Assert.Contains(evidence.FidelityDiagnostics, finding => finding.Code.StartsWith("PUB_", StringComparison.Ordinal));
            Assert.Contains(result.Diagnostics, finding => finding.Code.StartsWith("PUB_", StringComparison.Ordinal));
        } finally { Directory.Delete(root, true); }
    }
    [Fact]
    public async Task NativePagesAndRecoveryEvidenceReachThePublishedPdfAndPreview() {
        string root = CreateDirectory();
        try {
            string input = CopySource(root), output = Path.Combine(root, "result.pdf");
            byte[] before = File.ReadAllBytes(input);
            PublisherDocument source = PublisherDocument.Load(before);
            var result = await OfficeWorkflow.Convert(input).To(output).RunAsync();
            Assert.True(result.Succeeded, result.Summary);
            Assert.Equal("publisher-pdf", OfficeWorkflowCatalog.Find(".pub", ".pdf", executableOnly: true)?.Id);
            PdfReadDocument pdf = PdfReadDocument.Open(File.ReadAllBytes(output));
            Assert.Equal(source.Pages.Count, pdf.Pages.Count);
            var size = pdf.Pages[0].GetPageSize();
            Assert.Equal(source.Pages[0].Width, size.Width, 3); Assert.Equal(source.Pages[0].Height, size.Height, 3);
            Assert.Contains("This is some text", pdf.ExtractText());
            OfficeWorkflowConversionEvidence evidence = Assert.IsType<OfficeWorkflowConversionEvidence>(result.ConversionEvidence);
            Assert.Equal("PUB", evidence.Facts["sourceFormat"]); Assert.Equal("2", evidence.Facts["sourcePages"]);
            Assert.All(source.ReadReport.FidelityDiagnostics, finding => Assert.Contains(evidence.FidelityDiagnostics,
                retained => retained.Code == finding.Code && retained.LossKind == finding.LossKind && retained.Location == finding.Location));
            Assert.Contains(result.Diagnostics, finding => finding.Code == "OutputReopened");
            Assert.Equal(before, File.ReadAllBytes(input));
            OfficeWorkflowDocumentPreview preview = OfficeWorkflowRunner.PreviewDocument(before, ".pub");
            Assert.Equal(2, preview.Pages.Count); Assert.All(preview.Pages, page => Assert.NotEmpty(page.Bytes!));
            Assert.Contains(preview.Diagnostics, finding => finding.Code.StartsWith("PUB_", StringComparison.Ordinal));
        } finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData("strict")]
    [InlineData("pages")]
    [InlineData("input")]
    [InlineData("records")]
    [InlineData("output")]
    public async Task StrictAcceptanceAndNativeResourceLimitsPreserveAnExistingDestination(string limit) {
        string root = CreateDirectory();
        try {
            string input = CopySource(root), output = Path.Combine(root, "result.pdf");
            byte[] sentinel = [3, 2, 1]; File.WriteAllBytes(output, sentinel);
            var read = new PublisherReadOptions { MaximumPages = limit == "pages" ? 1 : 1024 };
            if (limit == "records") read.Limits.MaxRecords = 16;
            var request = OfficeWorkflow.Convert(input).To(output).WithConversionOptions(new() {
                RequireNoLoss = limit == "strict", PublisherRead = read
            }).Build();
            request.ConflictPolicy = OfficeWorkflowConflictPolicy.Replace;
            request.Limits = new() {
                MaximumInputBytes = limit == "input" ? 32 : 1024 * 1024,
                MaximumOutputBytes = limit == "output" ? 64 : 1024 * 1024
            };
            var result = await new OfficeWorkflowRunner().RunAsync(request);
            Assert.False(result.Succeeded); Assert.Equal(sentinel, File.ReadAllBytes(output));
            Assert.Empty(Directory.GetFiles(root, ".*.tmp"));
            if (limit is "strict" or "output") {
                var evidence = Assert.IsType<OfficeWorkflowConversionEvidence>(result.ConversionEvidence);
                Assert.True(evidence.HasLoss); Assert.Equal("2", evidence.Facts["sourcePages"]);
                Assert.Contains(evidence.FidelityDiagnostics, finding => finding.Code.StartsWith("PUB_", StringComparison.Ordinal));
            }
        } finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData("Sample2000.pub")]
    [InlineData("Sample98.pub")]
    public async Task EarlierGenerationsAreRejectedWithoutPublishingSalvagedText(string file) {
        string root = CreateDirectory();
        try {
            string input = CopySource(root, file), output = Path.Combine(root, "result.pdf");
            var result = await OfficeWorkflow.Convert(input).To(output).RunAsync();
            Assert.False(result.Succeeded); Assert.False(File.Exists(output));
            Assert.Equal(OfficeWorkflowFailureKind.UnsupportedInput, result.FailureKind);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void PublisherSettingsAreDetachedAndSelectedOnlyForTheirNativeRoute() {
        var options = new OfficeWorkflowConversionOptions {
            PublisherRead = new() { MaximumPages = 3 }, PublisherPdf = new(), RequireNoLoss = true
        };
        var copy = options.ForRoute("publisher-pdf");
        options.PublisherRead.MaximumPages = 30; options.PublisherRead.Limits.MaxRecords = 32;
        Assert.Equal(3, copy.PublisherRead!.MaximumPages); Assert.Equal(1_000_000, copy.PublisherRead.Limits.MaxRecords);
        Assert.NotSame(options.PublisherPdf, copy.PublisherPdf); Assert.True(copy.RequireNoLoss);
        var text = options.ForRoute("txt-pdf");
        Assert.Null(text.PublisherRead); Assert.Null(text.PublisherPdf); Assert.False(text.RequireNoLoss);
        Assert.Throws<ArgumentException>(() => options.Snapshot(OfficeWorkflowCatalog.FindExecutable("txt-pdf")!));
    }

    [Fact]
    public async Task CheckpointsCarryRecoveryLossAndRevalidateChangedNativeAcceptance() {
        string root = CreateDirectory();
        try {
            string input = CopySource(root); string text = Path.Combine(root, "literal.txt"); File.WriteAllText(text, "Plain text");
            var request = new OfficeConversionBatchRequest {
                InputPaths = [input, text], OutputDirectory = Path.Combine(root, "output"),
                CheckpointDirectory = Path.Combine(root, "state"), MaximumConcurrency = 1,
                ConversionOptions = new() { PublisherRead = new() { MaximumPages = 8 } }
            };
            var initialItems = new List<OfficeConversionBatchItemResult>();
            Assert.Equal(2, (await new OfficeWorkflowRunner().RunBatchAsync(request, new InlineProgress(initialItems))).Completed);
            Assert.Contains(Assert.Single(initialItems, item => item.InputPath == input).Diagnostics, finding => finding.Code.StartsWith("PUB_", StringComparison.Ordinal));
            var reusedItems = new List<OfficeConversionBatchItemResult>();
            Assert.Equal(2, (await new OfficeWorkflowRunner().RunBatchAsync(request, new InlineProgress(reusedItems))).Reused);
            Assert.Contains(Assert.Single(reusedItems, item => item.InputPath == input).Diagnostics, finding => finding.Code.StartsWith("PUB_", StringComparison.Ordinal));
            string output = Path.Combine(request.OutputDirectory, "source.pub.pdf"); byte[] before = File.ReadAllBytes(output);
            request.ConversionOptions.PublisherRead!.MaximumPages = 1;
            var changed = await new OfficeWorkflowRunner().RunBatchAsync(request);
            Assert.Equal(1, changed.Failed); Assert.Equal(1, changed.Reused); Assert.Equal(before, File.ReadAllBytes(output));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task ApplicationImageCodecsRequireAnOrdinaryBatchWithoutCheckpoints() {
        string root = CreateDirectory();
        try {
            var request = new OfficeConversionBatchRequest {
                InputPaths = [CopySource(root)], OutputDirectory = Path.Combine(root, "output"),
                CheckpointDirectory = Path.Combine(root, "state"), MaximumConcurrency = 1,
                ConversionOptions = new() { PublisherRead = new() { ImageCodec = new ApplicationCodec() } }
            };
            var items = new List<OfficeConversionBatchItemResult>();
            Assert.Equal(1, (await new OfficeWorkflowRunner().RunBatchAsync(request, new InlineProgress(items))).Failed);
            Assert.Contains("ordinary batch without checkpoints", Assert.Single(items).Summary, StringComparison.Ordinal);
            request.CheckpointDirectory = null;
            Assert.Equal(1, (await new OfficeWorkflowRunner().RunBatchAsync(request)).Completed);
        } finally { Directory.Delete(root, true); }
    }

    private static string CopySource(string root, string file = "Sample.pub") {
        string path = Path.Combine(root, "source.pub");
        File.Copy(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Publisher", file), path);
        return path;
    }
    private static string CreateDirectory() => Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "officeimo-publisher-workflow-" + Guid.NewGuid().ToString("N"))).FullName;
    private sealed class InlineProgress(List<OfficeConversionBatchItemResult> items) : IProgress<OfficeConversionBatchItemResult> {
        public void Report(OfficeConversionBatchItemResult item) => items.Add(item);
    }
    private sealed class ApplicationCodec : IOfficeRasterImageCodec {
        public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) { image = null; return false; }
    }
    private sealed class CancellingShaper(CancellationTokenSource source) : IOfficeTextShapingProvider {
        internal bool Called { get; private set; }
        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
            if (request.CancellationToken.CanBeCanceled) {
                Called = true;
                source.Cancel();
                request.CancellationToken.ThrowIfCancellationRequested();
            }
            return null;
        }
    }
}
