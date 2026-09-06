using OfficeIMO.Pdf;
using OfficeIMO.Security;

namespace OfficeIMO.Workflows.Tests;

public sealed class PdfProtectionWorkflowTests {
    [Theory]
    [InlineData("decrypt")]
    [InlineData("reencrypt-input")]
    [InlineData("encrypt-verification")]
    [InlineData("reencrypt-verification")]
    public void CancellationDuringCryptographicInspectionStopsFurtherParsing(string operation) {
        byte[] plain = PdfDocument.Create(document => {
            for (int index = 0; index < 20; index++) document.Page(page => page.Size(200, 300));
        }).ToBytes();
        byte[] encrypted = PdfDocument.Load(plain).Security.Encrypt(new("reader") {
            OwnerPassword = "owner", Algorithm = PdfStandardEncryptionAlgorithm.Aes128
        }).Pdf;
        using var cts = new CancellationTokenSource();
        var provider = new CancellingAesProvider(cts);
        var inputOptions = new PdfLoadOptions { AesCryptographyProvider = provider };
        var outputOptions = new PdfStandardEncryptionOptions("new-reader") {
            Algorithm = PdfStandardEncryptionAlgorithm.Aes128,
            AesCryptographyProvider = operation.EndsWith("verification", StringComparison.Ordinal) ? provider : null
        };
        Assert.Throws<OperationCanceledException>(() => {
            switch (operation) {
                case "decrypt": PdfSecurityEditor.Decrypt(encrypted, "owner", inputOptions, cancellationToken: cts.Token); break;
                case "reencrypt-input": PdfSecurityEditor.Reencrypt(encrypted, "owner", outputOptions, inputOptions, cancellationToken: cts.Token); break;
                case "encrypt-verification": PdfSecurityEditor.Encrypt(plain, outputOptions, cancellationToken: cts.Token); break;
                default: PdfSecurityEditor.Reencrypt(encrypted, "owner", outputOptions, cancellationToken: cts.Token); break;
            }
        });
        Assert.InRange(provider.DecryptOperations, 1, 2);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ProviderPublicationKeepsTheVerifiedProtectedCopyWhenWritingFails(bool failWrite) {
        await InDirectory(async root => {
            var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var resume = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            byte[] output = [];
            var recovery = new OfficeWorkflowOutputRecoveryStore(Path.Combine(root, "recovery"));
            var request = new OfficeWorkflowRequest {
                Operation = OfficeWorkflowOperation.ProtectPdf, InputPath = "content://provider/source",
                InputStream = new("source.pdf", async token => {
                    entered.TrySetResult(); await resume.Task.WaitAsync(token);
                    return new MemoryStream(CreatePdf(), writable: false);
                }), OutputPath = "content://provider/protected", ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                OutputEncryption = new("selected-reader") { OwnerPassword = "selected-owner" },
                OutputStream = new("protected.pdf", _ => Task.FromResult<Stream>(new MemoryStream(output)), _ => {
                    if (failWrite) throw new IOException("Provider denied write.");
                    return Task.FromResult<Stream>(new CommitStream(bytes => output = bytes));
                }, recovery)
            };
            Task<OfficeWorkflowResult> pending = new OfficeWorkflowRunner().RunAsync(request);
            await entered.Task;
            request.OutputEncryption.UserPassword = "later-reader"; request.OutputEncryption.OwnerPassword = "later-owner";
            resume.SetResult(); var result = await pending;
            Assert.Equal(failWrite ? OfficeWorkflowStatus.Unconfirmed : OfficeWorkflowStatus.Completed, result.Status);
            if (failWrite) {
                Assert.Null(result.OutputPath); Assert.NotNull(result.Recovery);
                await recovery.VerifyAsync(result.Recovery); output = File.ReadAllBytes(result.Recovery.FilePath);
            } else Assert.Empty(recovery.GetRecoveries());
            Assert.Equal(2, PdfDocument.Load(output, new PdfLoadOptions { Password = "selected-reader" }).Inspect().PageCount);
        });
    }

    [Theory]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes256)]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes128)]
    [InlineData(PdfStandardEncryptionAlgorithm.LegacyRc4)]
    public async Task SupportedEncryptionAlgorithmsProduceVerifiedOutputs(PdfStandardEncryptionAlgorithm algorithm) {
        await InDirectory(async root => {
            string source = Path.Combine(root, "source.pdf"); File.WriteAllBytes(source, CreatePdf());
            var result = await OfficeWorkflow.ProtectPdf(source, new("reader") { Algorithm = algorithm })
                .To(Path.Combine(root, "protected.pdf")).RunAsync();
            Assert.True(result.Succeeded, result.Summary);
            Assert.Equal(2, PdfDocument.Load(result.OutputPath!, new PdfLoadOptions { Password = "reader" }).Inspect().PageCount);
        });
    }

    [Fact]
    public async Task ProtectReplaceProtectionAndRemoveProtectionPreserveContentAndSources() {
        await InDirectory(async root => {
            string source = Path.Combine(root, "source.pdf"), protectedPath = Path.Combine(root, "protected.pdf");
            byte[] original = CreatePdf(); File.WriteAllBytes(source, original);
            var settings = new PdfStandardEncryptionOptions("reader-secret") {
                OwnerPassword = "owner-secret", AllowedPermissions = PdfStandardPermissions.Print, EncryptMetadata = false
            };
            var builder = OfficeWorkflow.ProtectPdf(source, settings).To(protectedPath);
            settings.UserPassword = "mutated";
            var built = builder.Build(); built.OutputEncryption!.OwnerPassword = "other";
            var protectedResult = await builder.RunAsync();
            Assert.True(protectedResult.Succeeded, protectedResult.Summary);
            Assert.True(protectedResult.HealthReport!.Verified);
            var info = PdfDocument.Load(protectedPath, new PdfLoadOptions { Password = "owner-secret" }).Inspect();
            Assert.Equal(2, info.PageCount); Assert.Equal(new[] { 200D, 210D }, info.Pages.Select(page => page.Width));
            Assert.Equal(settings.Permissions, info.Security.EncryptionPermissions);
            Assert.False(info.Security.EncryptMetadata); Assert.Equal(original, File.ReadAllBytes(source));
            string changed = Path.Combine(root, "changed.pdf");
            var reprotected = await OfficeWorkflow.ProtectPdf(protectedPath, new("new-reader") { OwnerPassword = "new-owner" })
                .WithPdfPassword("reader-secret").WithPdfOwnerPassword("owner-secret").To(changed).RunAsync();
            Assert.True(reprotected.Succeeded, reprotected.Summary);
            Assert.Equal(2, PdfDocument.Load(changed, new PdfLoadOptions { Password = "new-reader" }).Inspect().PageCount);
            var removed = await OfficeWorkflow.RemovePdfProtection(changed, "new-owner").To(Path.Combine(root, "clear.pdf")).RunAsync();
            Assert.True(removed.Succeeded, removed.Summary);
            Assert.False(PdfDocument.Load(removed.OutputPath!).Inspect().Security.HasEncryption);
            string serialized = System.Text.Json.JsonSerializer.Serialize(protectedResult);
            Assert.DoesNotContain("reader-secret", serialized); Assert.DoesNotContain("owner-secret", serialized);
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task UserPasswordCannotAuthorizeRemovingOrReplacingProtection(bool replace) {
        await InDirectory(async root => {
            string source = Path.Combine(root, "source.pdf"), destination = Path.Combine(root, "output.pdf");
            byte[] bytes = PdfDocument.Load(CreatePdf()).Security.Encrypt(new("reader") { OwnerPassword = "owner" }).Pdf;
            File.WriteAllBytes(source, bytes); File.WriteAllText(destination, "keep");
            var result = await new OfficeWorkflowRunner().RunAsync(new() {
                Operation = replace ? OfficeWorkflowOperation.ProtectPdf : OfficeWorkflowOperation.RemovePdfProtection,
                InputPath = source, PdfPassword = "reader", PdfOwnerPassword = "reader", OutputPath = destination,
                OutputEncryption = replace ? new("new-reader") : null, ConflictPolicy = OfficeWorkflowConflictPolicy.Replace
            });
            Assert.False(result.Succeeded); Assert.Null(result.OutputPath);
            Assert.Equal("keep", File.ReadAllText(destination)); Assert.Equal(bytes, File.ReadAllBytes(source));
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task BudgetAndCancellationLeaveExistingOutputIntact(bool cancelled) {
        await InDirectory(async root => {
            string source = Path.Combine(root, "source.pdf"), output = Path.Combine(root, "output.pdf");
            File.WriteAllBytes(source, CreatePdf()); File.WriteAllText(output, "keep");
            using var cts = new CancellationTokenSource(); if (cancelled) cts.Cancel();
            var result = await OfficeWorkflow.ProtectPdf(source, new("reader")).To(output)
                .WithLimits(1000000, 10).OnConflict(OfficeWorkflowConflictPolicy.Replace).RunAsync(cancellationToken: cts.Token);
            Assert.Equal(cancelled ? OfficeWorkflowStatus.Cancelled : OfficeWorkflowStatus.Failed, result.Status);
            Assert.Equal("keep", File.ReadAllText(output)); Assert.Equal(2, Directory.GetFiles(root).Length);
        });
    }

    [Fact]
    public void CanonicalProtectionEnforcesTheGenerationBudget() {
        Assert.Throws<InvalidDataException>(() => PdfSecurityEditor.Encrypt(CreatePdf(), new("reader"), maximumOutputBytes: 10));
    }

    [Fact]
    public async Task ProtectionCannotReplaceItsSource() {
        await InDirectory(async root => {
            string source = Path.Combine(root, "source.pdf"); byte[] original = CreatePdf(); File.WriteAllBytes(source, original);
            var result = await OfficeWorkflow.ProtectPdf(source, new("reader")).To(source).OnConflict(OfficeWorkflowConflictPolicy.Replace).RunAsync();
            Assert.False(result.Succeeded); Assert.Equal(original, File.ReadAllBytes(source));
        });
    }

    private static byte[] CreatePdf() => PdfDocument.Create(document => {
        document.Page(page => page.Size(200, 300)); document.Page(page => page.Size(210, 310));
    }).ToBytes();
    private sealed class CommitStream(Action<byte[]> commit) : MemoryStream {
        private bool _closed;
        protected override void Dispose(bool disposing) {
            if (!_closed) { _closed = true; commit(ToArray()); }
            base.Dispose(disposing);
        }
    }
    private sealed class CancellingAesProvider(CancellationTokenSource cancellation) : IOfficeAesCryptographyProvider {
        public string Name => "Cancellation test AES";
        public int DecryptOperations { get; private set; }
        public byte[] EncryptCbc(byte[] key, byte[] iv, byte[] plaintext, OfficeAesPadding padding) =>
            OfficeManagedAesCryptographyProvider.Default.EncryptCbc(key, iv, plaintext, padding);
        public byte[] DecryptCbc(byte[] key, byte[] iv, byte[] ciphertext, OfficeAesPadding padding) {
            DecryptOperations++; cancellation.Cancel();
            return OfficeManagedAesCryptographyProvider.Default.DecryptCbc(key, iv, ciphertext, padding);
        }
    }
    private static async Task InDirectory(Func<string, Task> action) {
        string root = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "officeimo-protect-tests-" + Guid.NewGuid().ToString("N"))).FullName;
        try { await action(root); } finally { Directory.Delete(root, true); }
    }
}
