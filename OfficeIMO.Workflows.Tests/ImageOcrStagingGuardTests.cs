using System.Text;
using OfficeIMO.Ocr;

namespace OfficeIMO.Workflows.Tests {
    public sealed class ImageOcrStagingGuardTests {
        [Fact]
        public async Task LocalStagingDenialDoesNotCreateAnUnadmittedDirectory() {
            using var scope = new Scope();
            string destination = Path.Combine(scope.Root, "unadmitted", "recognized.txt");
            var guard = new DestinationGuard(Path.Combine(scope.Root, "allowed"), destination, false);

            ImageOcrWorkflowResult result = await RunAsync(scope.Source, destination, guard);

            Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
            Assert.Null(result.OutputPath);
            Assert.False(Directory.Exists(Path.GetDirectoryName(destination)));
            Assert.False(File.Exists(destination));
            Assert.Equal(0, guard.PublicationChecks);
            Assert.Equal(scope.Image, File.ReadAllBytes(scope.Source));
        }

        [Fact]
        public async Task AdmittedLocalStagingStillRequiresFinalPublicationPermission() {
            using var scope = new Scope();
            string destination = Path.Combine(scope.Root, "recognized.txt");
            File.WriteAllText(destination, "Previous output");
            var guard = new DestinationGuard(scope.Root, destination, false);

            ImageOcrWorkflowResult result = await RunAsync(scope.Source, destination, guard);

            Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
            Assert.Null(result.OutputPath);
            Assert.True(guard.StagingChecks > 0);
            Assert.True(guard.PublicationChecks > 0);
            Assert.Equal("Previous output", File.ReadAllText(destination));
            Assert.Empty(Directory.GetFiles(scope.Root, ".image-ocr-*.tmp"));
            Assert.Equal(scope.Image, File.ReadAllBytes(scope.Source));
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public async Task ProviderTextUsesFinalPublicationPolicyWithoutLocalStagingAdmission(bool allowPublication) {
            using var scope = new Scope();
            const string destination = "content://scans/recognized";
            byte[] original = Encoding.UTF8.GetBytes("Previous output");
            byte[] stored = original;
            int writes = 0;
            var guard = new DestinationGuard(Path.Combine(scope.Root, "allowed"), destination, allowPublication);
            var recovery = new OfficeWorkflowOutputRecoveryStore(Path.Combine(scope.Root, "recovery"));
            var output = new OfficeWorkflowStreamOutput("recognized.txt",
                _ => Task.FromResult<Stream>(new MemoryStream(stored)),
                _ => {
                    writes++;
                    return Task.FromResult<Stream>(new DestinationStream(bytes => stored = bytes));
                }, recovery);

            ImageOcrWorkflowResult result = await RunAsync(scope.Source, destination, guard, output);

            Assert.Equal(allowPublication ? OfficeWorkflowStatus.Completed : OfficeWorkflowStatus.Failed, result.Status);
            Assert.Equal(0, guard.StagingChecks);
            Assert.True(guard.PublicationChecks > 0);
            Assert.Equal(allowPublication ? 1 : 0, writes);
            Assert.Equal(allowPublication ? "Recognized" : "Previous output", Encoding.UTF8.GetString(stored));
            Assert.Equal(allowPublication ? destination : null, result.OutputPath);
            Assert.Empty(recovery.GetRecoveries());
            Assert.False(Directory.Exists(guard.AllowedLocalRoot));
            Assert.Equal(scope.Image, File.ReadAllBytes(scope.Source));
        }

        private static Task<ImageOcrWorkflowResult> RunAsync(string source, string destination,
            DestinationGuard guard, OfficeWorkflowStreamOutput? output = null) => new OfficeWorkflowRunner().RecognizeImageAsync(
                new ImageOcrWorkflowRequest {
                    InputPath = source,
                    OutputPath = destination,
                    ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                    PublicationGuard = guard,
                    OutputStream = output
                }, new DelegateOcrEngine("fixture", (_, _) => Task.FromResult(new OcrResult { Text = "Recognized" })));

        private sealed class DestinationGuard(string allowedLocalRoot, string destination, bool allowPublication)
            : IOfficeWorkflowPublicationGuard, IOfficeWorkflowStagingGuard {
            internal string AllowedLocalRoot { get; } = allowedLocalRoot;
            internal int StagingChecks { get; private set; }
            internal int PublicationChecks { get; private set; }

            public ValueTask EnsureStagingDirectoryAllowedAsync(string directory, CancellationToken cancellationToken) {
                cancellationToken.ThrowIfCancellationRequested();
                StagingChecks++;
                if (!string.Equals(Path.GetFullPath(directory), Path.GetFullPath(AllowedLocalRoot), StringComparison.Ordinal))
                    throw new UnauthorizedAccessException("Only the configured local output root is admitted for staging.");
                return ValueTask.CompletedTask;
            }

            public ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken cancellationToken) {
                cancellationToken.ThrowIfCancellationRequested();
                PublicationChecks++;
                return ValueTask.FromResult(allowPublication && string.Equals(path, destination, StringComparison.Ordinal));
            }
        }

        private sealed class DestinationStream(Action<byte[]> commit) : MemoryStream {
            private bool _closed;
            protected override void Dispose(bool disposing) {
                if (disposing && !_closed) {
                    _closed = true;
                    commit(ToArray());
                }
                base.Dispose(disposing);
            }
        }

        private sealed class Scope : IDisposable {
            internal string Root { get; } = Path.Combine(Path.GetTempPath(), "officeimo-image-ocr-staging-" + Guid.NewGuid().ToString("N"));
            internal byte[] Image { get; } = OcrSessionWorkflowTests.Png();
            internal string Source => Path.Combine(Root, "scan.png");
            internal Scope() {
                Directory.CreateDirectory(Root);
                File.WriteAllBytes(Source, Image);
            }
            public void Dispose() => Directory.Delete(Root, recursive: true);
        }
    }
}
