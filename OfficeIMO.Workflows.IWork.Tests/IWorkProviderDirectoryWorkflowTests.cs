using System.IO.Compression;
using System.Runtime.CompilerServices;
using OfficeIMO.IWork;
using OfficeIMO.Workflows.IWork;
using Xunit;

namespace OfficeIMO.Workflows.IWork.Tests;

public sealed class IWorkProviderDirectoryWorkflowTests {
    [Theory]
    [InlineData("pages", "pages-docx", "docx")]
    [InlineData("numbers", "numbers-xlsx", "xlsx")]
    [InlineData("key", "keynote-pptx", "pptx")]
    public async Task Provider_packages_convert_from_opaque_locations(string kind, string route, string target) {
        using var package = new Package(kind, target);
        var request = package.Request(route);
        if (kind == "key") request.RegisteredConversionSettings = new IWorkWorkflowSettings {
            ConversionOptions = new() { AllowPartialEditableReconstruction = true }
        };
        var result = await IWorkWorkflow.CreateRunner().RunAsync(request);
        Assert.True(result.Succeeded, result.Summary);
        Assert.Contains(result.Diagnostics, d => d.Code == "OutputReopened");
        var capture = Assert.Single(result.Diagnostics, d => d.Code == "SourceSnapshot");
        Assert.Equal("source." + kind, capture.Details["sourceName"]);
        Assert.Equal("DirectoryPackage", capture.Details["snapshotKind"]);
        Assert.True(package.Enumerations >= 3);
        Assert.True(package.RootChecks >= 4);
    }

    [Theory]
    [InlineData("add")]
    [InlineData("remove")]
    [InlineData("content")]
    [InlineData("root")]
    public async Task Provider_changes_during_host_authorization_preserve_destination(string mutation) {
        using var package = new Package("pages", "docx");
        File.WriteAllText(package.Output, "keep");
        var request = package.Request("pages-docx");
        request.ConflictPolicy = OfficeWorkflowConflictPolicy.Replace;
        request.PublicationGuard = new Guard(() => {
            switch (mutation) {
                case "add": package.Files.Add("new.txt", [42]); break;
                case "remove": package.Files.Remove(package.Files.Keys.First()); break;
                case "content": package.Files[package.Files.Keys.First()] = [42]; break;
                case "root": package.RootChanged = true; break;
            }
            return true;
        });
        var result = await IWorkWorkflow.CreateRunner().RunAsync(request);
        Assert.False(result.Succeeded);
        Assert.Equal("keep", File.ReadAllText(package.Output));
    }

    [Theory]
    [InlineData("bytes")]
    [InlineData("entries")]
    [InlineData("guard")]
    [InlineData("cancel")]
    public async Task Provider_limits_and_access_denial_do_not_publish(string failure) {
        using var package = new Package("numbers", "xlsx");
        var request = package.Request("numbers-xlsx", failure == "entries" ? 1 : 10000);
        if (failure == "bytes") request.Limits.MaximumInputBytes = 1;
        if (failure == "guard") package.RootChanged = true;
        using var cancellation = new CancellationTokenSource();
        if (failure == "cancel") cancellation.Cancel();
        var result = await IWorkWorkflow.CreateRunner().RunAsync(request, cancellationToken: cancellation.Token);
        Assert.False(result.Succeeded);
        Assert.False(File.Exists(package.Output));
    }

    [Theory]
    [InlineData("stream")]
    [InlineData("output")]
    [InlineData("route")]
    public async Task Invalid_requests_do_not_open_provider(string failure) {
        using var package = new Package("numbers", "xlsx");
        var request = package.Request("numbers-xlsx");
        if (failure == "stream") request.InputStream = new("source.numbers", _ => throw new InvalidOperationException());
        if (failure == "output") request.OutputPath = null;
        if (failure == "route") request.ConversionRouteId = "docx-pdf";
        var result = await IWorkWorkflow.CreateRunner().RunAsync(request);
        Assert.Equal(OfficeWorkflowFailureKind.ValidationFailed, result.FailureKind);
        Assert.Equal(0, package.Enumerations);
        Assert.Equal(0, package.RootChecks);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Batch_rejects_output_inside_another_selected_provider_package(bool opaque) {
        using var first = new Package("pages", "docx");
        using var second = new Package("pages", "docx");
        string selected = opaque ? "provider://other/package" : Path.Combine(Path.GetDirectoryName(second.Output)!, "selected.pages");
        if (!opaque) Directory.CreateDirectory(selected);
        var firstRequest = first.Request("pages-docx");
        firstRequest.OutputPath = selected + "/nested/result.docx";
        if (opaque) firstRequest.OutputStream = new("result.docx", _ => throw new FileNotFoundException(), _ => throw new InvalidOperationException("Must not write"),
            new OfficeWorkflowOutputRecoveryStore(Path.Combine(Path.GetDirectoryName(first.Output)!, "recovery")));
        firstRequest.ConflictPolicy = OfficeWorkflowConflictPolicy.Replace;
        var secondRequest = second.Request("pages-docx");
        secondRequest.InputPath = selected;
        secondRequest.InputDirectoryPackage = new("source.pages", secondRequest.InputDirectoryPackage!.Directory,
            new PathGuard(selected));
        var results = await IWorkWorkflow.CreateRunner().RunBatchAsync([firstRequest, secondRequest]);
        Assert.False(results[0].Succeeded);
        Assert.Equal(0, first.Enumerations);
        Assert.True(results[1].Succeeded, results[1].Summary);
        if (!opaque) Assert.False(Directory.Exists(Path.Combine(selected, "nested")));
    }

    [Fact]
    public async Task Batch_does_not_capture_unscoped_provider_root_identity() {
        using var package = new Package("pages", "docx");
        string root = Path.Combine(Path.GetDirectoryName(package.Output)!, "scoped.pages");
        var request = package.Request("pages-docx");
        request.InputPath = root;
        // Models a root which becomes visible only when provider access is first acquired.
        request.InputDirectoryPackage = new("source.pages", request.InputDirectoryPackage!.Directory,
            new Guard(() => { Directory.CreateDirectory(root); return true; }));
        var result = Assert.Single(await IWorkWorkflow.CreateRunner().RunBatchAsync([request]));
        Assert.True(result.Succeeded, result.Summary);
    }

    private sealed class PathGuard(string root) : IOfficeWorkflowPublicationGuard {
        public ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken token) =>
            ValueTask.FromResult(path != root && !path.StartsWith(root + "/", StringComparison.Ordinal));
    }

    private sealed class Guard(Func<bool> check) : IOfficeWorkflowPublicationGuard {
        public ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken token) {
            token.ThrowIfCancellationRequested();
            return ValueTask.FromResult(check());
        }
    }

    private sealed class Package : IDisposable {
        private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-provider-package-" + Guid.NewGuid().ToString("N"));
        private readonly string _kind;
        internal Dictionary<string, byte[]> Files { get; } = new(StringComparer.Ordinal);
        internal string Output { get; }
        internal int Enumerations { get; private set; }
        internal int RootChecks { get; private set; }
        internal bool RootChanged { get; set; }
        internal Package(string kind, string target) {
            _kind = kind;
            Directory.CreateDirectory(_root);
            Output = Path.Combine(_root, "result." + target);
            using var zip = ZipFile.OpenRead(Path.Combine(AppContext.BaseDirectory, "Corpus", kind == "key" ? "tabledeck.key" : "simple." + kind));
            foreach (var entry in zip.Entries.Where(e => !string.IsNullOrEmpty(e.Name))) {
                using var stream = entry.Open();
                using var bytes = new MemoryStream();
                stream.CopyTo(bytes);
                Files.Add(entry.FullName, bytes.ToArray());
            }
        }
        internal OfficeWorkflowRequest Request(string route, int maximumEntries = 10000) => new() {
            Operation = OfficeWorkflowOperation.Convert, ConversionRouteId = route,
            InputPath = "provider://selected/package", OutputPath = Output,
            InputDirectoryPackage = new("source." + _kind, new(Enumerate), new Guard(() => { RootChecks++; return !RootChanged; }), maximumEntries)
        };
        private async IAsyncEnumerable<OfficeWorkflowDirectoryEntry> Enumerate(OfficeWorkflowDirectoryReadOptions options,
            [EnumeratorCancellation] CancellationToken token) {
            Enumerations++;
            var directories = new HashSet<string>(StringComparer.Ordinal);
            foreach (var pair in Files.OrderBy(p => p.Key, StringComparer.Ordinal)) {
                token.ThrowIfCancellationRequested();
                string[] segments = pair.Key.Split('/');
                for (int index = 1; index < segments.Length; index++) {
                    string directory = string.Join("/", segments.Take(index));
                    if (directories.Add(directory)) yield return new(directory, "provider://selected/package/" + directory, null);
                }
                yield return new(pair.Key, "provider://selected/package/" + pair.Key,
                    new(segments[^1], _ => Task.FromResult<Stream>(new MemoryStream(Files[pair.Key], writable: false))));
            }
            await Task.CompletedTask;
        }
        public void Dispose() => Directory.Delete(_root, true);
    }
}
