using OfficeIMO.Epub;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookPublishingBatchTests {
    [Fact]
    public async Task ReviewedProjectsExportAndCheckpointReuseVerifiesTheEpub() {
        using var scope = new Scope();
        var project = BookProject.Create("Publication"); project.CreateRevision("Editorial baseline");
        byte[] source = project.ToProjectBytes(); File.WriteAllBytes(scope.Source, source);
        var request = scope.Request with { CheckpointDirectory = scope.Checkpoint };
        var runner = new OfficeWorkflowRunner();
        var first = await runner.RunBatchAsync(request);
        Assert.Equal(1, first.Completed); Assert.Equal(0, first.Failed);
        Assert.Equal(project.Export().Bytes, File.ReadAllBytes(scope.Output));
        Assert.Equal("Publication", EpubPublication.Load(new MemoryStream(File.ReadAllBytes(scope.Output))).Title);
        Assert.Equal(source, File.ReadAllBytes(scope.Source));
        var second = await runner.RunBatchAsync(request);
        Assert.Equal(1, second.Reused); Assert.Equal(1, second.Completed);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportReviewCannotBeBypassedAndAcceptedLossRemainsVisible(bool acknowledge) {
        using var scope = new Scope();
        var project = BookProject.FromImport(EpubManuscript.ImportHtml(HtmlConversionDocument.Parse("<title>Book</title><h1>One</h1><p>Text</p><script>ignored()</script>")));
        if (acknowledge) project.AcknowledgeImportLoss();
        File.WriteAllBytes(scope.Source, project.ToProjectBytes());
        var items = new List<OfficeConversionBatchItemResult>();
        var result = await new OfficeWorkflowRunner().RunBatchAsync(scope.Request, new Capture(items));
        Assert.Equal(acknowledge ? 1 : 0, result.Completed);
        Assert.Equal(acknowledge ? 0 : 1, result.Failed);
        Assert.Equal(acknowledge, File.Exists(scope.Output));
        if (acknowledge) Assert.Contains(Assert.Single(items).Diagnostics, item => item.Code == "EPUB_IMPORT_ACTIVE_CONTENT_OMITTED" && item.Severity == OfficeWorkflowDiagnosticSeverity.Warning);
    }

    [Fact]
    public async Task OutputBoundsAndConflictsLeaveNoPartialPublication() {
        using var scope = new Scope(); File.WriteAllBytes(scope.Source, BookProject.Create("Book").ToProjectBytes());
        var runner = new OfficeWorkflowRunner();
        Assert.Equal(1, (await runner.RunBatchAsync(scope.Request with { MaximumOutputBytes = 16 })).Failed);
        Assert.False(File.Exists(scope.Output));
        Assert.Equal(1, (await runner.RunBatchAsync(scope.Request)).Completed);
        byte[] before = File.ReadAllBytes(scope.Output);
        Assert.Equal(1, (await runner.RunBatchAsync(scope.Request)).Failed);
        Assert.Equal(before, File.ReadAllBytes(scope.Output));
    }

    [Fact]
    public async Task CancellationAtPublicationCanResumeThroughTheExistingCheckpoint() {
        using var scope = new Scope(); File.WriteAllBytes(scope.Source, BookProject.Create("Book").ToProjectBytes());
        using var cancellation = new CancellationTokenSource();
        var request = scope.Request with { CheckpointDirectory = scope.Checkpoint };
        var runner = new OfficeWorkflowRunner();
        var interrupted = await runner.RunBatchAsync(request, cancellationToken: cancellation.Token,
            publicationGuard: new CancelPublication(scope.Output, cancellation));
        Assert.True(interrupted.Cancelled); Assert.False(File.Exists(scope.Output));
        Assert.Single(Directory.GetFiles(Path.GetDirectoryName(scope.Output)!, "*.conversion.epub"));
        var resumed = await runner.RunBatchAsync(request);
        Assert.Equal(1, resumed.Completed); Assert.Equal(0, resumed.Failed);
        Assert.Equal("Book", EpubPublication.Load(new MemoryStream(File.ReadAllBytes(scope.Output))).Title);
    }

    private sealed class Capture(List<OfficeConversionBatchItemResult> items) : IProgress<OfficeConversionBatchItemResult> {
        public void Report(OfficeConversionBatchItemResult value) { lock (items) items.Add(value); }
    }
    private sealed class CancelPublication(string output, CancellationTokenSource cancellation) : IOfficeWorkflowPublicationGuard {
        private int _checks;
        public ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken token) {
            if (!isDirectory && path == output && Interlocked.Increment(ref _checks) == 2) cancellation.Cancel();
            return ValueTask.FromResult(true);
        }
    }
    private sealed class Scope : IDisposable {
        private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-book-batch-" + Guid.NewGuid().ToString("N"));
        public Scope() => Directory.CreateDirectory(Input);
        public string Input => Path.Combine(_root, "input");
        public string Source => Path.Combine(Input, "book.oibook");
        public string Output => Path.Combine(_root, "output", "book.oibook.epub");
        public string Checkpoint => Path.Combine(_root, "checkpoint");
        public OfficeConversionBatchRequest Request => new() { InputDirectory = Input, OutputDirectory = Path.GetDirectoryName(Output)!, TargetExtension = ".epub", MaximumConcurrency = 1 };
        public void Dispose() => Directory.Delete(_root, true);
    }
}
