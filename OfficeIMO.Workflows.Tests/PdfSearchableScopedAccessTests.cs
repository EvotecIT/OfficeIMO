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
            var input = new ScopedFile(root, "source.pdf", pdf);
            var output = new ScopedFile(root, "output.pdf", pdf);
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

    // Models a provider whose filesystem name is accessible only during the returned stream's lifetime.
    private sealed class ScopedFile {
        internal ScopedFile(string root, string name, byte[] bytes) {
            Path = System.IO.Path.Combine(root, name);
            BackingPath = Path + ".outside-scope";
            File.WriteAllBytes(BackingPath, bytes);
        }
        internal string Path { get; }
        internal string BackingPath { get; }
        internal int Opens { get; private set; }
        internal int Closes { get; private set; }
        internal int Writes { get; private set; }
        internal Task<Stream> OpenRead(CancellationToken token) => Open(false, token);
        internal Task<Stream> OpenWrite(CancellationToken token) => Open(true, token);
        private Task<Stream> Open(bool write, CancellationToken token) {
            token.ThrowIfCancellationRequested();
            File.Move(BackingPath, Path);
            Opens++;
            if (write) Writes++;
            return Task.FromResult<Stream>(new ScopedStream(Path, write, () => {
                File.Move(Path, BackingPath);
                Closes++;
            }));
        }
    }

    private sealed class ScopedStream(string path, bool write, Action close)
        : FileStream(path, write ? FileMode.Create : FileMode.Open, write ? FileAccess.Write : FileAccess.Read, FileShare.Read) {
        private bool _closed;
        private void CloseScope() { if (!_closed) { _closed = true; close(); } }
        protected override void Dispose(bool disposing) { base.Dispose(disposing); if (disposing) CloseScope(); }
        public override async ValueTask DisposeAsync() { await base.DisposeAsync(); CloseScope(); }
    }

    private sealed class Guard(Func<bool> check) : IOfficeWorkflowPublicationGuard {
        public ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken token) => new(check());
    }
}
