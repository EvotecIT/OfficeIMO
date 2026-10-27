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
            var options = new PdfCreateOptions { TextShapingProvider = provider };
            options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Times New Roman", font));
            options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Arial", font));
            var request = ((new OfficeWorkflowRequest {
                Operation = OfficeWorkflowOperation.Convert,
                InputPath = input,
                OutputPath = output,
                ConversionOptions = (new OfficeConversionOptions() {
                    RegisteredSettings = new OfficeIMO.Workflows.OfficeWorkflowRouteSettings { PublisherPdf = options }
                }

    ).Clone()
,
                ConversionRouteId = OfficeWorkflowCatalog.Find(Path.GetExtension(input), Path.GetExtension(output), executableOnly: true)?.Id
            }));
            request.ConflictPolicy = OfficePublicationPolicy.Replace;
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
            var result = await OfficeWorkflow.RunAsync((new OfficeWorkflowRequest {
                Operation = OfficeWorkflowOperation.Convert,
                InputPath = input,
                OutputPath = output
,
                ConversionRouteId = OfficeWorkflowCatalog.Find(Path.GetExtension(input), Path.GetExtension(output), executableOnly: true)?.Id
            }));
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
            Assert.Contains(result.Diagnostics, finding => finding.Code == "conversion-output-reopened");
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
            var request = ((new OfficeWorkflowRequest {
                Operation = OfficeWorkflowOperation.Convert,
                InputPath = input,
                OutputPath = output,
                ConversionOptions = (new OfficeConversionOptions() {
                    FailOnLoss = limit == "strict",
                    RegisteredSettings = new OfficeIMO.Workflows.OfficeWorkflowRouteSettings { PublisherRead = read }
                }

    ).Clone()
,
                ConversionRouteId = OfficeWorkflowCatalog.Find(Path.GetExtension(input), Path.GetExtension(output), executableOnly: true)?.Id
            }));
            request.ConflictPolicy = OfficePublicationPolicy.Replace;
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
            var result = await OfficeWorkflow.RunAsync((new OfficeWorkflowRequest {
                Operation = OfficeWorkflowOperation.Convert,
                InputPath = input,
                OutputPath = output
,
                ConversionRouteId = OfficeWorkflowCatalog.Find(Path.GetExtension(input), Path.GetExtension(output), executableOnly: true)?.Id
            }));
            Assert.False(result.Succeeded); Assert.False(File.Exists(output));
            Assert.Equal(OfficeWorkflowFailureKind.UnsupportedInput, result.FailureKind);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void PublisherSettingsAreDetachedAndSelectedOnlyForTheirNativeRoute() {
        var options = new OfficeConversionOptions {
            FailOnLoss = true,
            RegisteredSettings = new OfficeWorkflowRouteSettings {
                PublisherRead = new() { MaximumPages = 3 },
                PublisherPdf = new()
            }
        };
        OfficeConversionOptions copy = options.ForRoute("publisher-pdf");
        var native = Assert.IsType<OfficeWorkflowRouteSettings>(options.RegisteredSettings);
        var copied = Assert.IsType<OfficeWorkflowRouteSettings>(copy.RegisteredSettings);
        native.PublisherRead!.MaximumPages = 30; native.PublisherRead.Limits.MaxRecords = 32;
        Assert.Equal(3, copied.PublisherRead!.MaximumPages); Assert.Equal(1_000_000, copied.PublisherRead.Limits.MaxRecords);
        Assert.NotSame(native.PublisherPdf, copied.PublisherPdf); Assert.True(copy.FailOnLoss);
        OfficeConversionOptions text = options.ForRoute("txt-pdf"); Assert.Null(text.RegisteredSettings); Assert.True(text.FailOnLoss);
        Assert.Throws<ArgumentException>(() => OfficeWorkflowRenderingOptions.Capture(options).Snapshot(OfficeWorkflowCatalog.FindExecutable("txt-pdf")!));
    }

    [Fact]
    public async Task CheckpointsCarryRecoveryLossAndRevalidateChangedNativeAcceptance() {
        string root = CreateDirectory();
        try {
            string input = CopySource(root); string text = Path.Combine(root, "literal.txt"); File.WriteAllText(text, "Plain text");
            var request = new OfficeBatchOptions {
                InputPaths = [input, text],
                OutputDirectory = Path.Combine(root, "output"),
                ResumeFrom = Path.Combine(Path.Combine(root, "state"), "plan.json"),
                MaxParallelism = 1,
                ConversionOptions = new() {
                    RegisteredSettings = new OfficeWorkflowRouteSettings { PublisherRead = new() { MaximumPages = 8 } }
                }
            };
            var initialItems = new List<OfficeBatchItemReport>();
            Assert.Equal(2, (await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), request, new InlineProgress(initialItems))).Completed);
            Assert.Contains(Assert.Single(initialItems, item => item.InputPath == input).Diagnostics, finding => finding.Code.StartsWith("PUB_", StringComparison.Ordinal));
            var reusedItems = new List<OfficeBatchItemReport>();
            Assert.Equal(2, (await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), request, new InlineProgress(reusedItems))).Reused);
            Assert.Contains(Assert.Single(reusedItems, item => item.InputPath == input).Diagnostics, finding => finding.Code.StartsWith("PUB_", StringComparison.Ordinal));
            string output = Path.Combine(request.OutputDirectory, "source.pub.pdf"); byte[] before = File.ReadAllBytes(output);
            ((OfficeIMO.Workflows.OfficeWorkflowRouteSettings)(request.ConversionOptions).RegisteredSettings!).PublisherRead!.MaximumPages = 1;
            var changed = await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), request);
            Assert.Equal(1, changed.Failed); Assert.Equal(1, changed.Reused); Assert.Equal(before, File.ReadAllBytes(output));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task ApplicationImageCodecsRequireAnOrdinaryBatchWithoutCheckpoints() {
        string root = CreateDirectory();
        try {
            var request = new OfficeBatchOptions {
                InputPaths = [CopySource(root)],
                OutputDirectory = Path.Combine(root, "output"),
                ResumeFrom = Path.Combine(Path.Combine(root, "state"), "plan.json"),
                MaxParallelism = 1,
                ConversionOptions = new() {
                    RegisteredSettings = new OfficeWorkflowRouteSettings { PublisherRead = new() { ImageCodec = new ApplicationCodec() } }
                }
            };
            var items = new List<OfficeBatchItemReport>();
            Assert.Equal(1, (await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), request, new InlineProgress(items))).Failed);
            Assert.Contains("ordinary batch without checkpoints", Assert.Single(items).Summary, StringComparison.Ordinal);
            request.ResumeFrom = null;
            Assert.Equal(1, (await OfficeBatchExecutor.RunAsync(new OfficeConversionExecutor(), request)).Completed);
        } finally { Directory.Delete(root, true); }
    }

    private static string CopySource(string root, string file = "Sample.pub") {
        string path = Path.Combine(root, "source.pub");
        File.Copy(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Publisher", file), path);
        return path;
    }
    private static string CreateDirectory() => Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "officeimo-publisher-workflow-" + Guid.NewGuid().ToString("N"))).FullName;
    private sealed class InlineProgress(List<OfficeBatchItemReport> items) : IProgress<OfficeBatchItemReport> {
        public void Report(OfficeBatchItemReport item) => items.Add(item);
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