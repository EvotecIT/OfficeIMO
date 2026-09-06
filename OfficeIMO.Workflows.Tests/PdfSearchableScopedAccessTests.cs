using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class PdfSearchableScopedAccessTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task LocalProviderPathsAreInspectedOnlyWhileTheirStreamsAreOpen(bool replaceSource) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-ocr-scopes-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            byte[] pdf = PdfDocument.Create(document => document.Page(page => page.Size(300, 300))).ToBytes();
            var input = new ScopedWorkflowFile(root, "source.pdf", pdf);
            var output = new ScopedWorkflowFile(root, "output.pdf", pdf);
            var store = new OfficeWorkflowOutputRecoveryStore(Path.Combine(root, "recovery"));
            int hostCalls = 0;
            var request = new PdfSearchableWorkflowRequest {
                InputPath = input.Path, InputStream = new("source.pdf", input.OpenRead),
                OutputPath = output.Path, OutputStream = new("output.pdf", output.OpenRead, output.OpenWrite, store),
                ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                PublicationGuard = new Guard(() => {
                    Assert.True(File.Exists(input.Path));
                    Assert.True(File.Exists(output.Path));
                    hostCalls++;
                    return true;
                })
            };
            var engine = new DelegateOcrEngine("fixture", (_, _) => {
                Assert.False(File.Exists(input.Path)); // Recognition must not retain the provider read stream.
                if (replaceSource) {
                    File.Move(input.BackingPath, input.BackingPath + ".original");
                    File.WriteAllBytes(input.BackingPath, pdf);
                }
                return Task.FromResult(new OcrResult { Provider = "fixture" });
            });
            var result = await new OfficeWorkflowRunner().MakePdfSearchableAsync(request, engine);
            Assert.Equal(replaceSource ? OfficeWorkflowStatus.Failed : OfficeWorkflowStatus.Completed, result.Status);
            Assert.Equal(replaceSource ? 0 : 1, hostCalls);
            Assert.Equal(replaceSource ? 0 : 1, output.Writes);
            Assert.False(File.Exists(input.Path));
            Assert.False(File.Exists(output.Path));
            Assert.Equal(input.Opens, input.Closes);
            Assert.Equal(output.Opens, output.Closes);
            Assert.Equal(pdf, File.ReadAllBytes(input.BackingPath));
            if (replaceSource) Assert.Equal(pdf, File.ReadAllBytes(output.BackingPath));
            else Assert.Equal(1, PdfDocument.Load(output.BackingPath).Inspect().PageCount);
            Assert.Empty(store.GetRecoveries());
        } finally { Directory.Delete(root, recursive: true); }
    }

    private sealed class Guard(Func<bool> check) : IOfficeWorkflowPublicationGuard {
        public ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken token) => new(check());
    }
}
