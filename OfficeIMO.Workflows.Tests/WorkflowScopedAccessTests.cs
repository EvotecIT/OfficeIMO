using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class WorkflowScopedAccessTests {
    [Theory]
    [InlineData("convert", false)]
    [InlineData("convert", true)]
    [InlineData("optimize", false)]
    [InlineData("optimize", true)]
    [InlineData("inspect", false)]
    [InlineData("inspect", true)]
    [InlineData("compare", false)]
    [InlineData("compare", true)]
    [InlineData("assemble", false)]
    [InlineData("assemble", true)]
    [InlineData("pages", false)]
    [InlineData("pages", true)]
    public async Task ProviderPublicationUsesActiveAccessAndRejectsReplacedSources(string operation, bool replaceSource) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-workflow-scopes-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            byte[] pdf = PdfDocument.Create(document => document.Page(page => page.Size(300, 300))).ToBytes();
            byte[] original = operation == "convert" ? Encoding.UTF8.GetBytes("<html><body>Scoped conversion</body></html>") : pdf;
            string inputName = operation == "convert" ? "source.html" : "source.pdf";
            var input = new ScopedWorkflowFile(root, inputName, original);
            var comparison = operation == "compare" ? new ScopedWorkflowFile(root, "comparison.pdf", pdf) : null;
            string outputName = operation == "compare" ? "comparison.html" : "output.pdf";
            var output = new ScopedWorkflowFile(root, outputName, pdf);
            var store = new OfficeWorkflowOutputRecoveryStore(Path.Combine(root, "recovery"));
            var sourceStream = new OfficeWorkflowStreamInput(inputName, input.OpenRead);
            var outputStream = new OfficeWorkflowStreamOutput(outputName, output.OpenRead, output.OpenWrite, store);
            bool replaced = false;
            int hostCalls = 0;
            var progress = new InlineProgress(update => {
                if (replaceSource && !replaced && update.Stage is "execute" or "normalize" or "render") {
                    Assert.False(File.Exists(input.Path));
                    File.Move(input.BackingPath, input.BackingPath + ".original");
                    File.WriteAllBytes(input.BackingPath, original);
                    replaced = true;
                }
            });
            var guard = new Guard(() => {
                Assert.True(File.Exists(input.Path));
                if (comparison is not null) Assert.True(File.Exists(comparison.Path));
                if (operation != "pages") Assert.True(File.Exists(output.Path));
                hostCalls++;
                return true;
            });
            var runner = new OfficeWorkflowRunner();
            OfficeWorkflowStatus status;
            string? published;
            string summary;
            bool hasHealthReport = false;
            string pageDirectory = Path.Combine(root, "images");
            if (operation == "assemble") {
                var result = await runner.AssemblePdfAsync(new() {
                    Sources = [input.Path], SourceStreams = new Dictionary<string, OfficeWorkflowStreamInput> { [input.Path] = sourceStream },
                    OutputPath = output.Path, OutputStream = outputStream, ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                    PublicationGuard = guard
                }, progress);
                status = result.Status; published = result.OutputPath; summary = result.Summary;
            } else if (operation == "pages") {
                var result = await runner.ExportPdfPagesAsync(new() {
                    InputPath = input.Path, InputStream = sourceStream, OutputDirectory = pageDirectory,
                    MaximumDimension = 128, PublicationGuard = guard
                }, progress);
                status = result.Status; published = result.OutputDirectory; summary = result.Summary;
            } else {
                var result = await runner.RunAsync(new() {
                    InputPath = input.Path, InputStream = sourceStream,
                    ComparisonPath = comparison?.Path, ComparisonStream = comparison is null ? null : new("comparison.pdf", comparison.OpenRead),
                    OutputPath = operation == "inspect" ? null : output.Path,
                    OutputStream = operation == "inspect" ? null : outputStream, ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                    Operation = operation == "convert" ? OfficeWorkflowOperation.Convert : operation == "compare"
                        ? OfficeWorkflowOperation.Compare : operation == "inspect" ? OfficeWorkflowOperation.Inspect : OfficeWorkflowOperation.Optimize,
                    ConversionRouteId = operation == "convert" ? "html-pdf" : null, PublicationGuard = guard
                }, progress);
                status = result.Status; published = result.OutputPath; summary = result.Summary;
                hasHealthReport = result.HealthReport is not null;
            }
            Assert.Equal(replaceSource, replaced);
            Assert.True(status == (replaceSource ? OfficeWorkflowStatus.Failed : OfficeWorkflowStatus.Completed), summary);
            Assert.Equal(replaceSource || operation == "inspect" ? 0 : 1, hostCalls);
            Assert.False(File.Exists(input.Path));
            Assert.False(File.Exists(output.Path));
            Assert.Equal(input.Opens, input.Closes);
            Assert.Equal(output.Opens, output.Closes);
            Assert.Equal(original, File.ReadAllBytes(input.BackingPath));
            if (replaceSource) {
                Assert.Contains("replaced", summary, StringComparison.OrdinalIgnoreCase);
                Assert.Null(published);
                Assert.Equal(0, output.Writes);
                Assert.Equal(pdf, File.ReadAllBytes(output.BackingPath));
                Assert.False(Directory.Exists(pageDirectory));
            } else if (operation == "inspect") {
                Assert.True(hasHealthReport);
                Assert.Null(published);
                Assert.Equal(0, output.Writes);
            } else if (operation == "pages") {
                Assert.Single(Directory.GetFiles(pageDirectory));
            } else {
                Assert.Equal(1, output.Writes);
                Assert.NotNull(published);
                if (operation == "compare") Assert.Contains("<html", File.ReadAllText(output.BackingPath), StringComparison.OrdinalIgnoreCase);
                else Assert.Equal(1, PdfDocument.Load(output.BackingPath).Inspect().PageCount);
            }
            Assert.Empty(store.GetRecoveries());
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public async Task LocalProviderOutputCanCreateASelectedNewFile() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-workflow-new-output-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string source = Path.Combine(root, "source.pdf");
            string output = Path.Combine(root, "new.pdf");
            PdfDocument.Create(document => document.Page(page => page.Size(300, 300))).Save(source);
            var store = new OfficeWorkflowOutputRecoveryStore(Path.Combine(root, "recovery"));
            var result = await new OfficeWorkflowRunner().RunAsync(new() {
                InputPath = source, Operation = OfficeWorkflowOperation.Optimize, OutputPath = output,
                OutputStream = new("new.pdf", _ => Task.FromResult<Stream>(File.OpenRead(output)),
                    _ => Task.FromResult<Stream>(File.Create(output)), store),
                ConflictPolicy = OfficeWorkflowConflictPolicy.Replace
            });
            Assert.True(result.Status == OfficeWorkflowStatus.Completed, result.Summary);
            Assert.Equal(output, result.OutputPath);
            Assert.Equal(1, PdfDocument.Load(output).Inspect().PageCount);
            Assert.Empty(store.GetRecoveries());
        } finally { Directory.Delete(root, recursive: true); }
    }

    private sealed class Guard(Func<bool> check) : IOfficeWorkflowPublicationGuard {
        public ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken token) => new(check());
    }
    private sealed class InlineProgress(Action<OfficeWorkflowProgress> report) : IProgress<OfficeWorkflowProgress> {
        public void Report(OfficeWorkflowProgress value) => report(value);
    }
}
