using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows.Tests;

public sealed class PdfSplitWorkflowTests {
    [Fact]
    public void PreviewPlanningHandlesLargePageNumbersWithoutOverflowAndRejectsExcessParts() {
        var plan = PdfSplitPlan.Create(int.MaxValue, int.MaxValue - 1);
        Assert.Equal(2, plan.Parts.Count);
        Assert.Equal(new PdfSplitPart("part-001.pdf", 1, int.MaxValue - 1), plan.Parts[0]);
        Assert.Equal(new PdfSplitPart("part-002.pdf", int.MaxValue, 1), plan.Parts[1]);
        plan.Validate(int.MaxValue, maximumParts: 1000);
        Assert.Throws<InvalidOperationException>(() => PdfSplitPlan.Create(int.MaxValue, 1));
    }

    [Fact]
    public void BookmarkPlansStartPartsAtEachTitleAndKeepLeadingPages() {
        var plan = PdfSplitPlan.FromStarts(10, [new PdfSplitStart(6, "Results: Q3/Q4"), new PdfSplitStart(3, "Intro"), new PdfSplitStart(3, "Duplicate")]);
        Assert.Equal([
            new PdfSplitPart("part-001.pdf", 1, 2),
            new PdfSplitPart("002-Intro.pdf", 3, 3),
            new PdfSplitPart("003-Results Q3 Q4.pdf", 6, 5)], plan.Parts);
        Assert.Throws<ArgumentException>(() => PdfSplitPlan.FromStarts(10, [new PdfSplitStart(11, "Outside")]));
    }

    [Fact]
    public async Task ExplicitPlansProduceOneVerifiedFilePerPart() {
        string root = NewRoot();
        try {
            string source = Path.Combine(root, "source.pdf");
            File.WriteAllBytes(source, CreatePdf());
            var result = await new OfficeWorkflowRunner().SplitPdfAsync(new() {
                InputPath = source, OutputDirectory = Path.Combine(root, "parts"),
                Plan = PdfSplitPlan.FromStarts(5, [new PdfSplitStart(1, "Cover"), new PdfSplitStart(2, "Body")])
            });
            Assert.Equal(OfficeWorkflowStatus.Completed, result.Status);
            Assert.Equal([1, 4], result.Files.Select(file => file.PageCount));
            Assert.Equal(["001-Cover.pdf", "002-Body.pdf"], result.Files.Select(file => Path.GetFileName(file.Path)));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void BookmarkPlanRejectsAnyInvalidStartAndKeepsSurrogatePairsIntact() {
        Assert.Throws<ArgumentException>(() => PdfSplitPlan.FromStarts(10,
            [new PdfSplitStart(1, "Valid"), new PdfSplitStart(11, "Invalid")]));
        Assert.Throws<ArgumentException>(() => PdfSplitPlan.FromStarts(10,
            [new PdfSplitStart(0, "Invalid"), new PdfSplitStart(2, "Valid")]));
        var plan = PdfSplitPlan.FromStarts(1, [new PdfSplitStart(1, new string('a', 59) + "😀more")]);
        Assert.Equal($"001-{new string('a', 59)}.pdf", Assert.Single(plan.Parts).Name);
    }

    [Fact]
    public async Task ExplicitPlanRejectsUncoveredSourcePagesBeforePublication() {
        string root = NewRoot();
        try {
            string source = Path.Combine(root, "source.pdf");
            File.WriteAllBytes(source, CreatePdf());
            string output = Path.Combine(root, "parts");
            var result = await new OfficeWorkflowRunner().SplitPdfAsync(new() {
                InputPath = source, OutputDirectory = output,
                Plan = PdfSplitPlan.FromStarts(3, [new PdfSplitStart(1, "Cover"), new PdfSplitStart(2, "Body")])
            });

            Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
            Assert.Empty(result.Files);
            Assert.False(Directory.Exists(output));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task InterruptedLocalReplacementReportsBothPreservedLocations() {
        string root = NewRoot();
        try {
            string source = Path.Combine(root, "source.pdf");
            File.WriteAllBytes(source, CreatePdf());
            string output = Directory.CreateDirectory(Path.Combine(root, "parts")).FullName;
            string recovery = Directory.CreateDirectory(output + ".officeimo-recovery-" + new string('3', 32)).FullName;
            OfficeWorkflowRunner.CreateDirectoryPublicationOwnershipMarker(output, new string('3', 32));
            File.WriteAllText(Path.Combine(output, "current.txt"), "current");
            File.WriteAllText(Path.Combine(recovery, "previous.txt"), "previous");
            var result = await new OfficeWorkflowRunner().SplitPdfAsync(new() {
                InputPath = source, OutputDirectory = output, ConflictPolicy = OfficeWorkflowConflictPolicy.Replace
            });
            Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
            var diagnostic = Assert.Single(result.Diagnostics, item => item.Code == "PdfSplitFailed");
            Assert.Equal(output, diagnostic.Details["destination"]);
            Assert.Contains(recovery, diagnostic.Details["recoveryPaths"]);
            Assert.Equal("current", File.ReadAllText(Path.Combine(output, "current.txt")));
            Assert.Equal("previous", File.ReadAllText(Path.Combine(recovery, "previous.txt")));
        } finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task GenerationStopsBetweenPartsForBudgetOrCancellation(bool cancelDuringGeneration) {
        string root = NewRoot();
        try {
            string source = Path.Combine(root, "source.pdf");
            File.WriteAllBytes(source, CreatePdf());
            using var cancellation = new CancellationTokenSource();
            var request = new PdfSplitWorkflowRequest { InputPath = source, OutputDirectory = Path.Combine(root, "parts") };
            if (!cancelDuringGeneration) request.Limits.MaximumOutputBytes = 10;
            int started = 0;
            var progress = new ProgressInline(item => {
                if (item.Stage == "split") {
                    started++;
                    if (cancelDuringGeneration && started == 2) cancellation.Cancel();
                }
            });
            var result = await new OfficeWorkflowRunner().SplitPdfAsync(request, progress, cancellation.Token);
            Assert.Equal(cancelDuringGeneration ? OfficeWorkflowStatus.Cancelled : OfficeWorkflowStatus.Failed, result.Status);
            Assert.Equal(cancelDuringGeneration ? 2 : 1, started);
            Assert.Empty(result.Files);
            Assert.False(Directory.Exists(request.OutputDirectory));
            Assert.Empty(Directory.GetDirectories(root, ".officeimo-split.*.tmp"));
        } finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData(1, 5)]
    [InlineData(2, 3)]
    [InlineData(int.MaxValue, 1)]
    public async Task LocalOutputsPreserveConsecutivePagesAndRenameConflictingFolders(int perPart, int expectedCount) {
        string root = NewRoot();
        try {
            string source = Path.Combine(root, "source.pdf");
            File.WriteAllBytes(source, CreatePdf());
            string output = Directory.CreateDirectory(Path.Combine(root, "parts")).FullName;
            File.WriteAllText(Path.Combine(output, "keep.txt"), "retained");
            var result = await new OfficeWorkflowRunner().SplitPdfAsync(new() {
                InputPath = source, OutputDirectory = output, PagesPerDocument = perPart
            });
            Assert.True(result.Succeeded, result.Summary);
            Assert.Equal(expectedCount, result.Files.Count);
            Assert.Equal("retained", File.ReadAllText(Path.Combine(output, "keep.txt")));
            int pageIndex = 0;
            foreach (var file in result.Files) {
                Assert.NotEqual(output, Path.GetDirectoryName(file.Path));
                Assert.Equal(pageIndex + 1, file.FirstSourcePage);
                var info = PdfDocument.Load(File.ReadAllBytes(file.Path)).Inspect();
                Assert.Equal(file.PageCount, info.PageCount);
                foreach (var page in info.Pages) Assert.Equal(200D + pageIndex++, page.Width);
            }
            Assert.Equal(5, pageIndex);
        } finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData("source-change")]
    [InlineData("part-limit")]
    [InlineData("byte-limit")]
    [InlineData("source-folder")]
    [InlineData("cancel")]
    public async Task InvalidOrChangedInputsDoNotPublish(string failure) {
        string root = NewRoot();
        try {
            string source = Path.Combine(root, "source.pdf");
            File.WriteAllBytes(source, CreatePdf());
            using var cancel = new CancellationTokenSource();
            string output = Path.Combine(root, "parts");
            var request = new PdfSplitWorkflowRequest { InputPath = source, OutputDirectory = output,
                ConflictPolicy = OfficeWorkflowConflictPolicy.Fail };
            if (failure == "part-limit") request.MaximumParts = 2;
            if (failure == "byte-limit") request.Limits.MaximumOutputBytes = 10;
            if (failure == "source-folder") { request.OutputDirectory = root; request.ConflictPolicy = OfficeWorkflowConflictPolicy.Replace; }
            var progress = new ProgressInline(item => {
                if (item.Stage != "publish") return;
                if (failure == "source-change") File.WriteAllBytes(source, CreatePdf(1));
                if (failure == "cancel") cancel.Cancel();
            });
            var result = await new OfficeWorkflowRunner().SplitPdfAsync(request, progress, cancel.Token);
            Assert.False(result.Succeeded);
            Assert.Empty(result.Files);
            Assert.False(Directory.Exists(output));
            Assert.True(File.Exists(source));
        } finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ProviderOutputsAreVerifiedAndPartialFailureRetainsRecovery(bool failSecond) {
        string root = NewRoot();
        try {
            var bytes = new Dictionary<string, byte[]>();
            var recovery = new OfficeWorkflowOutputRecoveryStore(Path.Combine(root, "recovery"));
            int writes = 0;
            var request = new PdfSplitWorkflowRequest {
                InputPath = "content://split/input", InputStream = new("input.pdf", _ => Task.FromResult<Stream>(new MemoryStream(CreatePdf()))),
                OutputDirectory = "content://split/output", PagesPerDocument = 2, ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                DirectoryOutput = new((name, _) => {
                    string path = "content://split/output/" + name;
                    return Task.FromResult(new OfficeWorkflowDirectoryOutputFile(path, new(name,
                        _ => Task.FromResult<Stream>(new MemoryStream(bytes.TryGetValue(path, out var value) ? value : throw new FileNotFoundException())),
                        _ => { writes++; return Task.FromResult<Stream>(new CommitStream(value => bytes[path] = failSecond && writes == 2 ? [] : value)); }, recovery)));
                })
            };
            var result = await new OfficeWorkflowRunner().SplitPdfAsync(request);
            Assert.Equal(failSecond ? OfficeWorkflowStatus.Unconfirmed : OfficeWorkflowStatus.Completed, result.Status);
            Assert.Equal(failSecond ? 1 : 3, result.Files.Count);
            foreach (var file in result.Files) Assert.Equal(file.PageCount, PdfDocument.Load(bytes[file.Path]).Inspect().PageCount);
            if (failSecond) {
                var retained = Assert.Single(result.OutputRecoveries);
                await recovery.VerifyAsync(retained);
                Assert.Equal(2, PdfDocument.Load(File.ReadAllBytes(retained.FilePath)).Inspect().PageCount);
            } else Assert.Empty(result.OutputRecoveries);
        } finally { Directory.Delete(root, true); }
    }

    private static string NewRoot() => Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "officeimo-split-tests-" + Guid.NewGuid().ToString("N"))).FullName;
    private static byte[] CreatePdf(int count = 5) => PdfDocument.Create(document => {
        for (int index = 0; index < count; index++) { int width = 200 + index; document.Page(page => page.Size(width, 350)); }
    }).ToBytes();
    private sealed class ProgressInline(Action<OfficeWorkflowProgress> report) : IProgress<OfficeWorkflowProgress> {
        public void Report(OfficeWorkflowProgress value) => report(value);
    }
    private sealed class CommitStream(Action<byte[]> commit) : MemoryStream {
        private bool _closed;
        protected override void Dispose(bool disposing) {
            if (!_closed) { _closed = true; commit(ToArray()); }
            base.Dispose(disposing);
        }
    }
}
