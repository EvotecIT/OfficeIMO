using OfficeIMO.Pdf;
using OfficeIMO.Workflows;

namespace OfficeIMO.Workflows.Tests {
    public sealed class ProviderStagingGuardTests {
        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public async Task ProviderOutputAppliesPublicationPolicyWithoutLocalDestinationStagingChecks(bool allowPublication) {
            using var scope = new Scope();
            string source = Path.Combine(scope.Root, "source.pdf");
            byte[] sourceBytes = PdfDocument.Create(document => document.Page(page => page.Size(240, 320))).ToBytes();
            File.WriteAllBytes(source, sourceBytes);
            byte[] original = PdfDocument.Create(document => document.Page(page => page.Size(120, 160).Margin(0))).ToBytes();
            byte[] destination = original;
            int writes = 0;
            const string providerLocation = "content://provider/selected";
            var guard = new DestinationGuard(Path.Combine(scope.Root, "local-output"), providerLocation, allowPublication);
            var recovery = new OfficeWorkflowOutputRecoveryStore(Path.Combine(scope.Root, "recovery"));

            OfficeWorkflowResult result = await new OfficeWorkflowRunner().RunAsync(new OfficeWorkflowRequest {
                InputPath = source,
                Operation = OfficeWorkflowOperation.Optimize,
                OutputPath = providerLocation,
                ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                PublicationGuard = guard,
                OutputStream = new OfficeWorkflowStreamOutput("selected.pdf",
                    _ => Task.FromResult<Stream>(new MemoryStream(destination)),
                    _ => {
                        writes++;
                        return Task.FromResult<Stream>(new DestinationStream(bytes => destination = bytes));
                    }, recovery)
            });

            Assert.True(guard.PublicationChecks > 0, result.Summary);
            Assert.Equal(allowPublication ? OfficeWorkflowStatus.Completed : OfficeWorkflowStatus.Failed, result.Status);
            Assert.Equal(allowPublication ? 1 : 0, writes);
            if (allowPublication) {
                Assert.Equal(providerLocation, result.OutputPath);
                Assert.Equal(1, PdfDocument.Load(destination).Inspect().PageCount);
                Assert.Equal(240D, PdfDocument.Load(destination).GetPageLayouts()[0].Width);
            } else {
                Assert.Equal(original, destination);
                Assert.Null(result.OutputPath);
            }
            Assert.Equal(sourceBytes, File.ReadAllBytes(source));
            Assert.Empty(recovery.GetRecoveries());
            Assert.False(Directory.Exists(guard.AllowedLocalRoot));
        }

        [Fact]
        public async Task LocalDestinationStagingPolicyRejectsBeforeCreatingAnUnadmittedDirectory() {
            using var scope = new Scope();
            string source = Path.Combine(scope.Root, "source.pdf");
            PdfDocument.Create(document => document.Page(page => page.Size(240, 320))).Save(source);
            string unadmitted = Path.Combine(scope.Root, "unadmitted", "output.pdf");
            var guard = new DestinationGuard(Path.Combine(scope.Root, "allowed"), "content://provider/selected", true);

            OfficeWorkflowResult result = await new OfficeWorkflowRunner().RunAsync(new OfficeWorkflowRequest {
                InputPath = source,
                Operation = OfficeWorkflowOperation.Optimize,
                OutputPath = unadmitted,
                PublicationGuard = guard
            });

            Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
            Assert.Equal(OfficeWorkflowFailureKind.OutputFailed, result.FailureKind);
            Assert.False(Directory.Exists(Path.GetDirectoryName(unadmitted)));
            Assert.False(File.Exists(unadmitted));
        }

        private sealed class DestinationGuard(string allowedLocalRoot, string providerLocation, bool allowPublication)
            : IOfficeWorkflowPublicationGuard, IOfficeWorkflowStagingGuard {
            internal string AllowedLocalRoot { get; } = allowedLocalRoot;
            internal int PublicationChecks { get; private set; }

            public ValueTask EnsureStagingDirectoryAllowedAsync(string directory, CancellationToken cancellationToken) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!string.Equals(Path.GetFullPath(directory), Path.GetFullPath(AllowedLocalRoot), StringComparison.Ordinal)) {
                    throw new UnauthorizedAccessException("Only the configured local output root is admitted for destination staging.");
                }
                return ValueTask.CompletedTask;
            }

            public ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken cancellationToken) {
                cancellationToken.ThrowIfCancellationRequested();
                PublicationChecks++;
                return ValueTask.FromResult(allowPublication && string.Equals(path, providerLocation, StringComparison.Ordinal));
            }
        }

        private sealed class DestinationStream(Action<byte[]> commit) : MemoryStream {
            private bool _closed;
            protected override void Dispose(bool disposing) {
                if (!_closed) {
                    _closed = true;
                    commit(ToArray());
                }
                base.Dispose(disposing);
            }
        }

        private sealed class Scope : IDisposable {
            internal string Root { get; } = Path.Combine(Path.GetTempPath(), "officeimo-provider-staging-" + Guid.NewGuid().ToString("N"));
            internal Scope() => Directory.CreateDirectory(Root);
            public void Dispose() => Directory.Delete(Root, recursive: true);
        }
    }
}
