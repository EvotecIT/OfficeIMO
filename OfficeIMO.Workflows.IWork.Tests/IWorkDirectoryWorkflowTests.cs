using System.IO.Compression;
using System.Runtime.InteropServices;
using OfficeIMO.IWork;
using OfficeIMO.Workflows.IWork;
using Xunit;

namespace OfficeIMO.Workflows.IWork.Tests;

public sealed class IWorkDirectoryWorkflowTests {
    [Theory]
    [InlineData("pages", "pages-docx", "docx")]
    [InlineData("numbers", "numbers-xlsx", "xlsx")]
    [InlineData("key", "keynote-pptx", "pptx")]
    public async Task Directory_packages_convert_with_verified_transport_snapshot(string kind, string route, string target) {
        using var bundle = new Bundle(kind, target);
        var request = bundle.Request(route);
        request.InputPath += Path.DirectorySeparatorChar;
        if (kind == "key") request.RegisteredConversionSettings = new IWorkWorkflowSettings {
            ConversionOptions = new IWorkConversionOptions { AllowPartialEditableReconstruction = true, RequireCompleteVisualCoverage = true }
        };
        var result = await IWorkWorkflow.CreateRunner().RunAsync(request);
        Assert.True(result.Succeeded, result.Summary);
        Assert.True(File.Exists(bundle.Output));
        Assert.Contains(result.Diagnostics, item => item.Code == "OutputReopened");
        var snapshot = Assert.Single(result.Diagnostics, item => item.Code == "SourceSnapshot");
        Assert.Equal("DirectoryPackage", snapshot.Details["snapshotKind"]);
        Assert.Equal(64, snapshot.Details["sha256"].Length);
        Assert.NotNull(result.ConversionEvidence);
        if (kind == "key") {
            Assert.Equal("True", result.ConversionEvidence!.Facts["partialEditableReconstruction"]);
            Assert.Contains(result.ConversionEvidence.FidelityDiagnostics, d => d.Code == "IWORK_KEYNOTE_PARAGRAPH_PAGINATION_OMITTED");
        }
        Assert.True(result.InputBytes > Directory.GetFiles(bundle.Input, "*", SearchOption.AllDirectories).Sum(path => new FileInfo(path).Length));
    }

    [Theory]
    [InlineData("add")]
    [InlineData("remove")]
    [InlineData("change")]
    [InlineData("replace")]
    public async Task Membership_content_and_root_replacement_prevent_publication(string mutation) {
        using var bundle = new Bundle("pages", "docx");
        string entry = Directory.GetFiles(bundle.Input, "*", SearchOption.AllDirectories)[0];
        var request = bundle.Request("pages-docx");
        request.PublicationGuard = new ActionGuard(() => {
            switch (mutation) {
                case "add": File.WriteAllText(Path.Combine(bundle.Input, "new-entry.txt"), "changed"); break;
                case "remove": File.Delete(entry); break;
                case "change": File.AppendAllText(entry, "changed"); break;
                case "replace":
                    string old = bundle.Input + ".old";
                    Directory.Move(bundle.Input, old);
                    Directory.CreateDirectory(bundle.Input);
                    foreach (string file in Directory.GetFiles(old, "*", SearchOption.AllDirectories)) {
                        string output = Path.Combine(bundle.Input, Path.GetRelativePath(old, file));
                        Directory.CreateDirectory(Path.GetDirectoryName(output)!);
                        File.Copy(file, output);
                    }
                    break;
            }
        });
        var result = await IWorkWorkflow.CreateRunner().RunAsync(request);
        Assert.False(result.Succeeded);
        Assert.False(File.Exists(bundle.Output));
        Assert.Contains(result.Diagnostics, item => item.Message.Contains("source", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public async Task Native_member_rename_with_unchanged_normalized_archive_path_prevents_publication() {
        if (OperatingSystem.IsWindows()) return;
        using var bundle = new Bundle("numbers", "xlsx");
        string member = Path.Combine(bundle.Input, "Data\\extra.bin");
        File.WriteAllBytes(member, [1, 2, 3, 4]);
        var request = bundle.Request("numbers-xlsx");
        request.PublicationGuard = new ActionGuard(() => {
            string directory = Path.Combine(bundle.Input, "Data");
            Directory.CreateDirectory(directory);
            File.Move(member, Path.Combine(directory, "extra.bin"));
        });
        var result = await IWorkWorkflow.CreateRunner().RunAsync(request);
        Assert.False(result.Succeeded);
        Assert.False(File.Exists(bundle.Output));
        Assert.Contains(result.Diagnostics, item => item.Message.Contains("membership", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public async Task Nested_index_archive_is_not_duplicated_by_transport_capture() {
        using var bundle = new Bundle("numbers", "xlsx");
        string index = Path.Combine(bundle.Input, "Index");
        ZipFile.CreateFromDirectory(index, Path.Combine(bundle.Input, "Index.zip"));
        Directory.Delete(index, recursive: true);
        IWorkSourceDocument source = IWorkSourceDocument.Open(bundle.Input);
        var result = await IWorkWorkflow.CreateRunner().RunAsync(bundle.Request("numbers-xlsx"));
        Assert.True(result.Succeeded, result.Summary);
        Assert.Equal(source.Records.Count.ToString(), result.ConversionEvidence!.Facts["totalRecordCount"]);
        Assert.Equal("EditableReconstruction", result.ConversionEvidence.Facts["projectionKind"]);
    }

    [Fact]
    public async Task Destination_inside_source_package_is_rejected_without_changes() {
        using var bundle = new Bundle("pages", "docx");
        string existing = Path.Combine(bundle.Input, "protected.docx");
        File.WriteAllText(existing, "keep");
        var request = bundle.Request("pages-docx");
        request.OutputPath = existing;
        request.ConflictPolicy = OfficeWorkflowConflictPolicy.Replace;
        var result = await IWorkWorkflow.CreateRunner().RunAsync(request);
        Assert.Equal(OfficeWorkflowFailureKind.ValidationFailed, result.FailureKind);
        Assert.Equal("keep", File.ReadAllText(existing));
        Assert.DoesNotContain(result.Diagnostics, item => item.Code == "SourceSnapshot");
    }

    [Fact]
    public async Task Directory_read_bounds_and_cancellation_leave_no_output() {
        using var bundle = new Bundle("numbers", "xlsx");
        var request = bundle.Request("numbers-xlsx");
        request.RegisteredConversionSettings = new IWorkWorkflowSettings { ReadOptions = new() { MaximumPackageBytes = 16 } };
        var bounded = await IWorkWorkflow.CreateRunner().RunAsync(request);
        Assert.False(bounded.Succeeded);
        Assert.False(File.Exists(bundle.Output));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var cancelled = await IWorkWorkflow.CreateRunner().RunAsync(bundle.Request("numbers-xlsx"), cancellationToken: cancellation.Token);
        Assert.Equal(OfficeWorkflowStatus.Cancelled, cancelled.Status);
        Assert.False(File.Exists(bundle.Output));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Provider_output_hard_link_to_bundle_member_is_rejected_before_write(bool literalBackslash) {
        if (literalBackslash && OperatingSystem.IsWindows()) return;
        using var bundle = new Bundle("numbers", "xlsx");
        string member = literalBackslash ? Path.Combine(bundle.Input, "Data\\extra.bin")
            : Directory.GetFiles(bundle.Input, "*", SearchOption.AllDirectories)[0];
        if (literalBackslash) File.WriteAllBytes(member, [1, 2, 3, 4]);
        byte[] original = File.ReadAllBytes(member);
        bool linked = OperatingSystem.IsWindows() ? CreateHardLink(bundle.Output, member, IntPtr.Zero) : Link(member, bundle.Output) == 0;
        Assert.True(linked, "The test filesystem must support hard links.");
        var request = bundle.Request("numbers-xlsx");
        request.ConflictPolicy = OfficeWorkflowConflictPolicy.Replace;
        bool openedWrite = false;
        var store = new OfficeWorkflowOutputRecoveryStore(Path.Combine(bundle.Root, "recovery"));
        request.OutputStream = new OfficeWorkflowStreamOutput("result.xlsx", _ => Task.FromResult<Stream>(File.OpenRead(bundle.Output)),
            _ => { openedWrite = true; return Task.FromResult<Stream>(File.Create(bundle.Output)); }, store);
        var result = await IWorkWorkflow.CreateRunner().RunAsync(request);
        Assert.False(result.Succeeded);
        Assert.False(openedWrite);
        Assert.Equal(original, File.ReadAllBytes(member));
        Assert.Equal(original, File.ReadAllBytes(bundle.Output));
        foreach (var recovery in store.GetRecoveries()) store.Discard(recovery);
    }

    [Fact]
    public async Task Batch_cannot_publish_inside_another_local_package() {
        using var first = new Bundle("pages", "docx");
        using var second = new Bundle("pages", "docx");
        var request = first.Request("pages-docx");
        request.OutputPath = Path.Combine(second.Input, "nested", "result.docx");
        var results = await IWorkWorkflow.CreateRunner().RunBatchAsync([request, second.Request("pages-docx")]);
        Assert.False(results[0].Succeeded);
        Assert.True(results[1].Succeeded, results[1].Summary);
        Assert.False(Directory.Exists(Path.Combine(second.Input, "nested")));
    }

    [DllImport("libc", EntryPoint = "link", SetLastError = true)]
    private static extern int Link(string source, string destination);
    [DllImport("kernel32.dll", EntryPoint = "CreateHardLinkW", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CreateHardLink(string destination, string source, IntPtr reserved);

    private sealed class ActionGuard(Action mutation) : IOfficeWorkflowPublicationGuard {
        private bool _changed;
        public ValueTask<bool> CanPublishAsync(string outputPath, bool isDirectory, CancellationToken token) {
            if (!_changed) { _changed = true; mutation(); }
            return ValueTask.FromResult(true);
        }
    }

    private sealed class Bundle : IDisposable {
        public Bundle(string kind, string target) {
            Root = Path.Combine(Path.GetTempPath(), "officeimo-iwork-directory-" + Guid.NewGuid().ToString("N"));
            Input = Path.Combine(Root, "source." + kind);
            Output = Path.Combine(Root, "result." + target);
            Directory.CreateDirectory(Input);
            ZipFile.ExtractToDirectory(Path.Combine(AppContext.BaseDirectory, "Corpus", kind == "key" ? "tabledeck.key" : "simple." + kind), Input);
        }
        public string Root { get; }
        public string Input { get; }
        public string Output { get; }
        public OfficeWorkflowRequest Request(string route) => new() { Operation = OfficeWorkflowOperation.Convert,
            InputPath = Input, OutputPath = Output, ConversionRouteId = route };
        public void Dispose() => Directory.Delete(Root, recursive: true);
    }
}
