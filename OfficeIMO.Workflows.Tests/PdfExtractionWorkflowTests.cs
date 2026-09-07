using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows.Tests;

public sealed class PdfExtractionWorkflowTests {
    [Fact]
    public void CanonicalExtractionEnforcesBudgetBeforeReturningAnOversizedDocument() {
        var document = PdfDocument.Load(CreatePdf());
        Assert.Throws<InvalidDataException>(() => document.Pages.Extract(Enumerable.Repeat(1, 100000).ToArray(), 10));
        Assert.Throws<InvalidDataException>(() => document.Pages.Split(PdfPageRange.From(1, 3), 10));
        var extracted = document.Pages.Extract(new[] { 3, 1, 3 }, 100000);
        Assert.Equal(new[] { 202D, 200D, 202D }, extracted.Inspect().Pages.Select(page => page.Width));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SourceReplacementAndSourceDestinationAreRejected(bool replaceSource) {
        await InDirectory(async root => {
            string source = Path.Combine(root, "source.pdf");
            byte[] original = CreatePdf();
            File.WriteAllBytes(source, original);
            var progress = new InlineProgress(value => {
                if (replaceSource && value.Stage == "execute") File.WriteAllText(source, "external replacement");
            });
            var result = await OfficeWorkflow.ExtractPages(source, 1)
                .To(replaceSource ? Path.Combine(root, "output.pdf") : source)
                .OnConflict(OfficeWorkflowConflictPolicy.Replace).RunAsync(progress: progress);
            Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
            Assert.Null(result.OutputPath);
            if (replaceSource) Assert.Equal("external replacement", File.ReadAllText(source));
            else Assert.Equal(original, File.ReadAllBytes(source));
            Assert.Single(Directory.GetFiles(root));
        });
    }

    [Fact]
    public async Task OrderedRepeatedPagesPublishWithoutChangingSourceOrExistingDestination() {
        await InDirectory(async root => {
            string source = Path.Combine(root, "source.pdf");
            byte[] original = CreatePdf();
            File.WriteAllBytes(source, original);
            string destination = Path.Combine(root, "output.pdf");
            File.WriteAllText(destination, "keep");
            int[] pages = [3, 1, 3];
            var builder = OfficeWorkflow.ExtractPages(source, pages).To(destination);
            pages[0] = 2;
            var request = builder.Build();
            request.PageNumbers![0] = 2;
            var result = await builder.RunAsync();
            Assert.True(result.Succeeded, result.Summary);
            Assert.NotEqual(destination, result.OutputPath);
            Assert.Equal("keep", File.ReadAllText(destination));
            Assert.Equal(original, File.ReadAllBytes(source));
            Assert.Equal(new[] { 202D, 200D, 202D }, PdfDocument.Load(File.ReadAllBytes(result.OutputPath!))
                .Inspect().Pages.Select(page => page.Width));
        });
    }

    [Theory]
    [InlineData(0)]
    [InlineData(4)]
    [InlineData(100001)]
    public async Task InvalidSelectionsCannotReplaceDestination(int invalid) {
        await InDirectory(async root => {
            string source = Path.Combine(root, "source.pdf"), output = Path.Combine(root, "output.pdf");
            File.WriteAllBytes(source, CreatePdf());
            File.WriteAllText(output, "keep");
            var result = await new OfficeWorkflowRunner().RunAsync(new() {
                Operation = OfficeWorkflowOperation.ExtractPages, InputPath = source, OutputPath = output,
                ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                PageNumbers = invalid == 100001 ? Enumerable.Repeat(1, invalid).ToArray() : [invalid]
            });
            Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
            Assert.Null(result.OutputPath);
            Assert.Equal("keep", File.ReadAllText(output));
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task BudgetOrCancellationPreservesDestination(bool cancel) {
        await InDirectory(async root => {
            string source = Path.Combine(root, "source.pdf"), output = Path.Combine(root, "output.pdf");
            File.WriteAllBytes(source, CreatePdf());
            File.WriteAllText(output, "keep");
            using var cts = new CancellationTokenSource();
            if (cancel) cts.Cancel();
            var result = await OfficeWorkflow.ExtractPages(source, 1, 2).To(output)
                .OnConflict(OfficeWorkflowConflictPolicy.Replace).WithLimits(1000000, 10)
                .RunAsync(cancellationToken: cts.Token);
            Assert.Equal(cancel ? OfficeWorkflowStatus.Cancelled : OfficeWorkflowStatus.Failed, result.Status);
            Assert.Equal("keep", File.ReadAllText(output));
            Assert.Equal(2, Directory.GetFiles(root).Length);
        });
    }

    [Fact]
    public async Task ProviderAwaitCannotChangeCapturedPageOrder() {
        await InDirectory(async root => {
            var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var resume = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var request = new OfficeWorkflowRequest {
                Operation = OfficeWorkflowOperation.ExtractPages, InputPath = "urn:test:source",
                OutputPath = Path.Combine(root, "extracted.pdf"), PageNumbers = [3, 1],
                InputStream = new("source.pdf", async token => {
                    entered.TrySetResult();
                    await resume.Task.WaitAsync(token);
                    return new MemoryStream(CreatePdf(), writable: false);
                })
            };
            var running = new OfficeWorkflowRunner().RunAsync(request);
            await entered.Task;
            request.PageNumbers[0] = 2;
            resume.SetResult();
            var result = await running;
            Assert.True(result.Succeeded, result.Summary);
            Assert.Equal(new[] { 202D, 200D }, PdfDocument.Load(File.ReadAllBytes(result.OutputPath!))
                .Inspect().Pages.Select(page => page.Width));
        });
    }

    private static byte[] CreatePdf() => PdfDocument.Create(document => {
        for (int index = 0; index < 3; index++) { int width = 200 + index; document.Page(page => page.Size(width, 350)); }
    }).ToBytes();
    private sealed class InlineProgress(Action<OfficeWorkflowProgress> report) : IProgress<OfficeWorkflowProgress> {
        public void Report(OfficeWorkflowProgress value) => report(value);
    }
    private static async Task InDirectory(Func<string, Task> action) {
        string root = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "officeimo-extract-tests-" + Guid.NewGuid().ToString("N"))).FullName;
        try { await action(root); } finally { Directory.Delete(root, true); }
    }
}
