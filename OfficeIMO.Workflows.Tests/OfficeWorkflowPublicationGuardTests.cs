using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows.Tests;

public sealed class OfficeWorkflowPublicationGuardTests {
    [Theory]
    [InlineData("convert", OfficeWorkflowConflictPolicy.Replace)]
    [InlineData("assembly", OfficeWorkflowConflictPolicy.Replace)]
    [InlineData("images", OfficeWorkflowConflictPolicy.Replace)]
    [InlineData("convert", OfficeWorkflowConflictPolicy.Fail)]
    [InlineData("assembly", OfficeWorkflowConflictPolicy.Fail)]
    [InlineData("images", OfficeWorkflowConflictPolicy.Fail)]
    public async Task DeniedPublicationPreservesExistingOutputAndCleansStaging(string operation, OfficeWorkflowConflictPolicy policy) {
        using var scope = new Scope();
        string target = Path.Combine(scope.Root, operation == "images" ? "output" : "output.pdf");
        string sentinel = operation == "images" ? Path.Combine(Directory.CreateDirectory(target).FullName, "previous.txt") : target;
        File.WriteAllText(sentinel, "previous output");
        var guard = new Guard((path, directory, token) => {
            Assert.Equal(target, path);
            Assert.Equal(operation == "images", directory);
            return ValueTask.FromResult(false);
        });

        var result = await RunAsync(scope, operation, target, policy, guard);

        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.Null(result.Path);
        Assert.Equal("previous output", File.ReadAllText(sentinel));
        Assert.Equal(1, guard.Calls);
        Assert.DoesNotContain(Directory.EnumerateFileSystemEntries(scope.Root), path => Path.GetFileName(path).StartsWith('.'));
        Assert.DoesNotContain(Directory.EnumerateFileSystemEntries(scope.Root), path => path.Contains("officeimo-recovery-"));
    }

    [Theory]
    [InlineData("convert")]
    [InlineData("assembly")]
    [InlineData("images")]
    public async Task RenameChecksActualCandidateEvenWhenOwnedPathDoesNotExist(string operation) {
        using var scope = new Scope();
        string target = Path.Combine(scope.Root, operation == "images" ? "output" : "output.pdf");
        var guard = new Guard((path, _, _) => ValueTask.FromResult(path != target));
        var result = await RunAsync(scope, operation, target, OfficeWorkflowConflictPolicy.Rename, guard);
        Assert.Equal(OfficeWorkflowStatus.Completed, result.Status);
        Assert.Equal(operation == "images" ? Path.Combine(scope.Root, "output (1)") : Path.Combine(scope.Root, "output (1).pdf"), result.Path);
        Assert.False(File.Exists(target));
        Assert.False(Directory.Exists(target));
        Assert.Equal(2, guard.Calls);
        if (operation != "images") {
            var pdf = PdfDocument.Load(File.ReadAllBytes(result.Path!));
            Assert.True(pdf.Inspect().PageCount > 0);
        } else {
            Assert.NotEmpty(Directory.GetFiles(result.Path!, "*.svg"));
        }
    }

    [Theory]
    [InlineData("convert", false)]
    [InlineData("convert", true)]
    [InlineData("assembly", false)]
    [InlineData("assembly", true)]
    [InlineData("images", false)]
    [InlineData("images", true)]
    public async Task CancellationOrGuardFailureCannotPublish(string operation, bool throws) {
        using var scope = new Scope();
        using var cancellation = new CancellationTokenSource();
        string target = Path.Combine(scope.Root, operation == "images" ? "output" : "output.pdf");
        var guard = new Guard(async (_, _, _) => {
            await Task.Yield();
            if (throws) throw new IOException("Ownership unavailable");
            cancellation.Cancel();
            return true;
        });
        var result = await RunAsync(scope, operation, target, OfficeWorkflowConflictPolicy.Rename, guard, cancellation.Token);
        Assert.Equal(throws ? OfficeWorkflowStatus.Failed : OfficeWorkflowStatus.Cancelled, result.Status);
        Assert.False(File.Exists(target));
        Assert.False(Directory.Exists(target));
        Assert.Null(result.Path);
    }

    private static async Task<(OfficeWorkflowStatus Status, string? Path)> RunAsync(Scope scope, string operation,
        string target, OfficeWorkflowConflictPolicy policy, IOfficeWorkflowPublicationGuard guard, CancellationToken token = default) {
        var runner = new OfficeWorkflowRunner();
        if (operation == "assembly") {
            var result = await runner.AssemblePdfAsync(new PdfAssemblyRequest {
                Sources = [scope.Pdf], OutputPath = target, ConflictPolicy = policy, PublicationGuard = guard
            }, cancellationToken: token);
            return (result.Status, result.OutputPath);
        }
        if (operation == "images") {
            var result = await runner.ExportPdfPagesAsync(new PdfPageImageExportRequest {
                InputPath = scope.Pdf, OutputDirectory = target, Format = OfficeImageExportFormat.Svg,
                ConflictPolicy = policy, PublicationGuard = guard
            }, cancellationToken: token);
            return (result.Status, result.OutputDirectory);
        }
        var conversion = await runner.RunAsync(new OfficeWorkflowRequest {
            InputPath = scope.Html, OutputPath = target, ConversionRouteId = "html-pdf",
            ConflictPolicy = policy, PublicationGuard = guard
        }, cancellationToken: token);
        return (conversion.Status, conversion.OutputPath);
    }

    private sealed class Guard(Func<string, bool, CancellationToken, ValueTask<bool>> check) : IOfficeWorkflowPublicationGuard {
        public int Calls { get; private set; }
        public ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken cancellationToken) {
            Calls++;
            return check(path, isDirectory, cancellationToken);
        }
    }

    private sealed class Scope : IDisposable {
        public string Root { get; } = Path.Combine(Path.GetTempPath(), "officeimo-publication-" + Guid.NewGuid().ToString("N"));
        public string Pdf => Path.Combine(Root, "source.pdf");
        public string Html => Path.Combine(Root, "source.html");
        public Scope() {
            Directory.CreateDirectory(Root);
            File.WriteAllText(Html, "<html><body><p>Publication acceptance</p></body></html>");
            PdfDocument.Create(document => document.Page(page => page.Size(200, 300))).Save(Pdf);
        }
        public void Dispose() => Directory.Delete(Root, recursive: true);
    }
}
