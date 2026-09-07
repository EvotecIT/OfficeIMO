using System.Runtime.CompilerServices;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class WorkflowDirectoryInputTests {
    [Fact]
    public async Task RelativeHtmlResourcesArePreservedWithoutAddingAnImagePage() {
        using var fixture = new Fixture();
        fixture.Entries.Add(new("page", "provider://folder/page", null));
        fixture.Entries.Add(File("page/source.html", Encoding.UTF8.GetBytes("<html><head><link rel=\"stylesheet\" href=\"style.css\"></head><body><p>Folder resource proof</p><img src=\"pixel.png\"></body></html>")));
        fixture.Entries.Add(File("page/style.css", Encoding.UTF8.GetBytes("p { color: #123456; }")));
        fixture.Entries.Add(File("page/pixel.png", Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=")));
        var result = await new OfficeWorkflowRunner().AssemblePdfAsync(fixture.Request());
        Assert.True(result.Succeeded, result.Summary);
        Assert.Equal(1, result.SourceCount);
        Assert.Equal(1, PdfDocument.Load(result.OutputPath!).Inspect().PageCount);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Message.Contains("not found", StringComparison.OrdinalIgnoreCase));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task LocalMembersKeepTheirAccessScopesAndRejectPhysicalReplacement(bool replace) {
        using var fixture = new Fixture();
        byte[] bytes = PdfDocument.Create(document => document.Page(page => page.Size(300, 300))).ToBytes();
        var file = new ScopedWorkflowFile(Path.GetDirectoryName(fixture.Output)!, "scoped.pdf", bytes);
        fixture.Entries.Add(new("scoped.pdf", file.Path, new("scoped.pdf", file.OpenRead)));
        var request = fixture.Request();
        int hostCalls = 0;
        request.PublicationGuard = new Guard(() => { hostCalls++; Assert.True(System.IO.File.Exists(file.Path)); return true; });
        var result = await new OfficeWorkflowRunner().AssemblePdfAsync(request, new Progress(update => {
            if (!replace || update.Stage != "normalize") return;
            System.IO.File.Move(file.BackingPath, file.BackingPath + ".original");
            System.IO.File.WriteAllBytes(file.BackingPath, bytes);
        }));
        Assert.True(result.Status == (replace ? OfficeWorkflowStatus.Failed : OfficeWorkflowStatus.Completed), result.Summary);
        Assert.Equal(replace ? 0 : 1, hostCalls);
        Assert.Equal(file.Opens, file.Closes);
        Assert.Equal(bytes, System.IO.File.ReadAllBytes(file.BackingPath));
        if (replace) Assert.Contains("replaced", result.Summary);
    }

    [Fact]
    public async Task PortableNameCollisionsAndCancelledTraversalDoNotPublish() {
        using var fixture = new Fixture();
        fixture.Entries.Add(File("one.pdf", []));
        fixture.Entries.Add(File("ONE.pdf", []));
        var conflict = await new OfficeWorkflowRunner().AssemblePdfAsync(fixture.Request());
        Assert.Equal(OfficeWorkflowStatus.Failed, conflict.Status);
        Assert.Contains("conflicting", conflict.Summary);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var cancelled = await new OfficeWorkflowRunner().AssemblePdfAsync(fixture.Request(), cancellationToken: cancellation.Token);
        Assert.Equal(OfficeWorkflowStatus.Cancelled, cancelled.Status);
        Assert.False(System.IO.File.Exists(fixture.Output));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PreservesNestedDocumentsAndReopensOutput(bool recursive) {
        using var fixture = new Fixture();
        fixture.Entries.Add(File("first.pdf", PdfDocument.Create(document => document.Page(page => page.Size(300, 300))).ToBytes()));
        fixture.Entries.Add(new("nested", "provider://folder/nested", null));
        fixture.Entries.Add(File("nested/second.html", Encoding.UTF8.GetBytes("<html><body>Nested document</body></html>")));
        var request = fixture.Request();
        request.Options.IncludeSubdirectories = recursive;
        var result = await new OfficeWorkflowRunner().AssemblePdfAsync(request);
        Assert.True(result.Succeeded, result.Summary);
        Assert.Equal(recursive ? 2 : 1, result.SourceCount);
        Assert.Equal(recursive ? 2 : 1, PdfDocument.Load(result.OutputPath!).Inspect().PageCount);
        Assert.True(fixture.Enumerations >= 3);
    }

    [Theory]
    [InlineData("../escape.pdf")]
    [InlineData("/root.pdf")]
    [InlineData("nested\\escape.pdf")]
    [InlineData("nested/missing-parent.pdf")]
    [InlineData("NUL.pdf")]
    [InlineData("name.pdf:stream")]
    public async Task RejectsUnsafeMemberWithoutOpeningIt(string path) {
        using var fixture = new Fixture();
        int opens = 0;
        fixture.Entries.Add(new(path, "provider://folder/member", new("member.pdf", _ => { opens++; throw new IOException(); })));
        var result = await new OfficeWorkflowRunner().AssemblePdfAsync(fixture.Request());
        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.Equal(0, opens);
        Assert.False(System.IO.File.Exists(fixture.Output));
    }

    [Theory]
    [InlineData("membership")]
    [InlineData("bytes")]
    [InlineData("host-membership")]
    [InlineData("host-bytes")]
    [InlineData("new-item")]
    public async Task ChangedFolderCannotPublish(string change) {
        using var fixture = new Fixture();
        byte[] original = PdfDocument.Create(document => document.Page(page => page.Size(300, 300))).ToBytes();
        byte[] bytes = original;
        fixture.Entries.Add(new("first.pdf", "provider://folder/first.pdf", new("first.pdf", _ => Task.FromResult<Stream>(new MemoryStream(bytes)))));
        void Mutate() {
            if (change.EndsWith("membership", StringComparison.Ordinal)) fixture.Entries.Add(File("added.pdf", original));
            else if (change == "new-item") fixture.Entries[0] = File("first.pdf", [.. original, 32]);
            else bytes = [.. original, 32];
        }
        var request = fixture.Request();
        request.PublicationGuard = new Guard(() => { if (change.StartsWith("host-", StringComparison.Ordinal)) Mutate(); return true; });
        var result = await new OfficeWorkflowRunner().AssemblePdfAsync(request,
            new Progress(update => { if (update.Stage == "normalize" && !change.StartsWith("host-", StringComparison.Ordinal)) Mutate(); }));
        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.Contains("changed", result.Summary, StringComparison.OrdinalIgnoreCase);
        Assert.False(System.IO.File.Exists(fixture.Output));
    }

    [Fact]
    public async Task EntryAndByteLimitsApplyBeforePublication() {
        using var fixture = new Fixture();
        byte[] bytes = PdfDocument.Create(document => document.Page(page => page.Size(300, 300))).ToBytes();
        fixture.Entries.Add(File("first.pdf", bytes));
        fixture.Entries.Add(File("second.pdf", bytes));
        var count = fixture.Request();
        count.Options.MaximumDiscoveredEntries = 1;
        var countResult = await new OfficeWorkflowRunner().AssemblePdfAsync(count);
        Assert.Equal(OfficeWorkflowStatus.Failed, countResult.Status);
        Assert.Contains("entry limit", countResult.Summary);
        var size = fixture.Request();
        size.Limits.MaximumInputBytes = bytes.Length;
        var sizeResult = await new OfficeWorkflowRunner().AssemblePdfAsync(size);
        Assert.Equal(OfficeWorkflowStatus.Failed, sizeResult.Status);
        Assert.Contains("input limit", sizeResult.Summary);
        Assert.False(System.IO.File.Exists(fixture.Output));
    }

    private static OfficeWorkflowDirectoryEntry File(string relative, byte[] bytes) => new(relative,
        "provider://folder/" + relative, new(Path.GetFileName(relative), _ => Task.FromResult<Stream>(new MemoryStream(bytes))));

    private sealed class Fixture : IDisposable {
        internal readonly List<OfficeWorkflowDirectoryEntry> Entries = [];
        private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-directory-test-" + Guid.NewGuid().ToString("N"));
        internal string Output => Path.Combine(_root, "assembled.pdf");
        internal int Enumerations;
        internal Fixture() => Directory.CreateDirectory(_root);
        internal PdfAssemblyRequest Request() => new() {
            Sources = ["provider://folder/root"], OutputPath = Output,
            SourceDirectories = new Dictionary<string, OfficeWorkflowDirectoryInput> { ["provider://folder/root"] = new(Enumerate) }
        };
        private async IAsyncEnumerable<OfficeWorkflowDirectoryEntry> Enumerate(OfficeWorkflowDirectoryReadOptions options,
            [EnumeratorCancellation] CancellationToken token) {
            Enumerations++;
            await Task.CompletedTask;
            foreach (var entry in Entries.ToArray()) {
                token.ThrowIfCancellationRequested();
                if (options.IncludeSubdirectories || !entry.RelativePath.Contains('/')) yield return entry;
            }
        }
        public void Dispose() => Directory.Delete(_root, true);
    }

    private sealed class Progress(Action<OfficeWorkflowProgress> callback) : IProgress<OfficeWorkflowProgress> {
        public void Report(OfficeWorkflowProgress value) => callback(value);
    }
    private sealed class Guard(Func<bool> callback) : IOfficeWorkflowPublicationGuard {
        public ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken token) => ValueTask.FromResult(callback());
    }
}
