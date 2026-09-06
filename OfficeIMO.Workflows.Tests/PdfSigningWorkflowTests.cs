using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Pdf;
using OfficeIMO.Security;

namespace OfficeIMO.Workflows.Tests;

public sealed class PdfSigningWorkflowTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ExistingSignaturePolicyRejectsAnotherSignatureAndPreservesTheOriginal(bool certified) {
        await InDirectory(async root => {
            using var certificate = CreateCertificate();
            using var signer = new PdfCmsExternalSigner(OfficeSecurityProvider.Default, certificate);
            var verifier = new PdfCmsSignatureCryptographyProvider(OfficeSecurityProvider.Default);
            byte[] original = PdfDocument.Load(CreatePdf()).Security.SignExternal(signer, new() {
                FieldName = "First", Profile = certified ? PdfSignatureProfile.Certification : PdfSignatureProfile.Approval,
                CertificationPermission = PdfCertificationPermissionLevel.NoChanges
            }).Pdf;
            string source = Path.Combine(root, "source.pdf"), output = Path.Combine(root, "output.pdf");
            File.WriteAllBytes(source, original);
            var result = await OfficeWorkflow.SignPdf(source, signer, new() { FieldName = "Second" }, verifier).To(output).RunAsync();
            Assert.False(PdfDocument.Load(original).PlanMutation(PdfMutationOperation.PrepareExternalSignature).CanExecute);
            Assert.False(result.Succeeded);
            Assert.Equal(original, File.ReadAllBytes(source));
            Assert.False(File.Exists(output));
            Assert.True(PdfDocument.Load(File.ReadAllBytes(source)).Security.ValidateSignatures(verifier).MathematicalSignaturesVerified);
        });
    }

    [Fact]
    public async Task CancellationInsideSignerPreventsPublicationAndRetainsCallerOwnership() {
        await InDirectory(async root => {
            using var certificate = CreateCertificate();
            using var signer = new PdfCmsExternalSigner(OfficeSecurityProvider.Default, certificate);
            using var cancellation = new CancellationTokenSource();
            var cancelling = new CancellingSigner(signer, cancellation);
            string source = Path.Combine(root, "source.pdf"), output = Path.Combine(root, "output.pdf");
            byte[] original = CreatePdf(); File.WriteAllBytes(source, original); File.WriteAllText(output, "keep");
            var result = await OfficeWorkflow.SignPdf(source, cancelling, new(), new PdfCmsSignatureCryptographyProvider(OfficeSecurityProvider.Default))
                .To(output).OnConflict(OfficeWorkflowConflictPolicy.Replace).RunAsync(cancellationToken: cancellation.Token);
            Assert.Equal(OfficeWorkflowStatus.Cancelled, result.Status); Assert.Equal(1, cancelling.Calls);
            Assert.Equal(original, File.ReadAllBytes(source)); Assert.Equal("keep", File.ReadAllText(output));
            Assert.NotEmpty(PdfDocument.Load(original).Security.SignExternal(signer, new()).Pdf);
        });
    }

    private sealed class CancellingSigner(IPdfExternalSigner signer, CancellationTokenSource cancellation) : IPdfExternalSigner {
        public string Name => signer.Name;
        public int Calls { get; private set; }
        public byte[] Sign(PdfExternalSignatureRequest request) { Calls++; var bytes = signer.Sign(request); cancellation.Cancel(); return bytes; }
    }

    [Fact]
    public async Task WaitingInputUsesCapturedSignerAndAppearanceSettings() {
        await InDirectory(async root => {
            using var certificate = CreateCertificate();
            using var signer = new PdfCmsExternalSigner(OfficeSecurityProvider.Default, certificate);
            var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var resume = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var request = new OfficeWorkflowRequest {
                Operation = OfficeWorkflowOperation.SignPdf, InputPath = "content://signing/source",
                InputStream = new("source.pdf", async token => { entered.TrySetResult(); await resume.Task.WaitAsync(token); return new MemoryStream(CreatePdf(), false); }),
                OutputPath = Path.Combine(root, "signed.pdf"), OutputSigner = signer,
                OutputSignatureOptions = new() { FieldName = "Captured", VisibleAppearance = new() { PageNumber = 2, Width = 100 } },
                OutputSignatureValidator = new PdfCmsSignatureCryptographyProvider(OfficeSecurityProvider.Default)
            };
            Task<OfficeWorkflowResult> pending = new OfficeWorkflowRunner().RunAsync(request);
            await entered.Task; request.OutputSigner = new InvalidSigner(); request.OutputSignatureOptions.FieldName = "Later";
            request.OutputSignatureOptions.VisibleAppearance!.PageNumber = 100;
            resume.SetResult(); var result = await pending;
            Assert.True(result.Succeeded, result.Summary);
            Assert.Equal("Captured", Assert.Single(result.SignatureReport!.Signatures).Signature.FieldName);
        });
    }

    [Fact]
    public async Task SignedCopyCapturesSettingsAndReportsCryptographySeparatelyFromTrust() {
        await InDirectory(async root => {
            byte[] original = CreatePdf(); string source = Path.Combine(root, "source.pdf"); File.WriteAllBytes(source, original);
            using var certificate = CreateCertificate();
            using var signer = new PdfCmsExternalSigner(OfficeSecurityProvider.Default, certificate);
            var verifier = new PdfCmsSignatureCryptographyProvider(OfficeSecurityProvider.Default);
            var options = new PdfExternalSignatureOptions { FieldName = "Approval", Reason = "Reviewed",
                VisibleAppearance = new() { PageNumber = 2, X = 20, Y = 20, Width = 100, Height = 40, Text = "Reviewed copy" } };
            var builder = OfficeWorkflow.SignPdf(source, signer, options, verifier).To(Path.Combine(root, "signed.pdf"));
            options.FieldName = "Later"; options.VisibleAppearance.PageNumber = 1;
            builder.Build().OutputSignatureOptions!.FieldName = "Another";
            var result = await builder.RunAsync();
            Assert.True(result.Succeeded, result.Summary); Assert.True(result.HealthReport!.Verified);
            var report = Assert.IsType<PdfSignatureValidationReport>(result.SignatureReport);
            Assert.True(report.MathematicalSignaturesVerified); Assert.True(report.DigestVerified);
            Assert.False(report.CertificateChainVerified); Assert.False(report.CryptographicTrustVerified);
            Assert.Equal("Approval", Assert.Single(report.Signatures).Signature.FieldName);
            var saved = PdfDocument.Load(File.ReadAllBytes(result.OutputPath!));
            Assert.Equal(2, saved.Inspect().PageCount);
            Assert.True(saved.Security.ValidateSignatures(verifier).MathematicalSignaturesVerified);
            Assert.Equal(original, File.ReadAllBytes(source));
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ProviderPublicationRetainsSignedEvidenceAndRecovery(bool fail) {
        await InDirectory(async root => {
            using var certificate = CreateCertificate();
            using var signer = new PdfCmsExternalSigner(OfficeSecurityProvider.Default, certificate);
            byte[] output = [];
            var recovery = new OfficeWorkflowOutputRecoveryStore(Path.Combine(root, "recovery"));
            var result = await new OfficeWorkflowRunner().RunAsync(new() {
                Operation = OfficeWorkflowOperation.SignPdf, InputPath = "content://signing/source",
                InputStream = new("source.pdf", _ => Task.FromResult<Stream>(new MemoryStream(CreatePdf(), false))),
                OutputPath = "content://signing/output", ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                OutputSigner = signer, OutputSignatureOptions = new(), OutputSignatureValidator = new PdfCmsSignatureCryptographyProvider(OfficeSecurityProvider.Default),
                OutputStream = new("signed.pdf", _ => Task.FromResult<Stream>(new MemoryStream(output, false)), _ => {
                    if (fail) throw new IOException("Write refused");
                    return Task.FromResult<Stream>(new CommitStream(bytes => output = bytes));
                }, recovery)
            });
            Assert.Equal(fail ? OfficeWorkflowStatus.Unconfirmed : OfficeWorkflowStatus.Completed, result.Status);
            Assert.True(result.SignatureReport!.MathematicalSignaturesVerified);
            if (fail) { Assert.Null(result.OutputPath); Assert.NotNull(result.Recovery); await recovery.VerifyAsync(result.Recovery); output = File.ReadAllBytes(result.Recovery.FilePath); }
            Assert.True(PdfDocument.Load(output).Security.ValidateSignatures(new PdfCmsSignatureCryptographyProvider(OfficeSecurityProvider.Default)).MathematicalSignaturesVerified);
        });
    }

    [Theory]
    [InlineData("invalid-signature")]
    [InlineData("cancel")]
    [InlineData("same-source")]
    [InlineData("restricted")]
    [InlineData("budget")]
    public async Task RejectedSigningNeverReplacesExistingBytes(string failure) {
        await InDirectory(async root => {
            using var certificate = CreateCertificate();
            using var signer = new PdfCmsExternalSigner(OfficeSecurityProvider.Default, certificate);
            string source = Path.Combine(root, "source.pdf"), output = Path.Combine(root, "output.pdf");
            byte[] original = CreatePdf();
            if (failure == "restricted") original = PdfDocument.Load(original).Security.Encrypt(new("reader") { OwnerPassword = "owner", AllowedPermissions = PdfStandardPermissions.None }).Pdf;
            File.WriteAllBytes(source, original); File.WriteAllText(output, "keep");
            using var cancellation = new CancellationTokenSource(); if (failure == "cancel") cancellation.Cancel();
            var result = await new OfficeWorkflowRunner().RunAsync(new() {
                Operation = OfficeWorkflowOperation.SignPdf, InputPath = source, OutputPath = failure == "same-source" ? source : output,
                ConflictPolicy = OfficeWorkflowConflictPolicy.Replace, PdfPassword = failure == "restricted" ? "reader" : null,
                OutputSigner = failure == "invalid-signature" ? new InvalidSigner() : signer,
                OutputSignatureOptions = new(), OutputSignatureValidator = new PdfCmsSignatureCryptographyProvider(OfficeSecurityProvider.Default),
                Limits = new() { MaximumOutputBytes = failure == "budget" ? 10 : 1000000 }
            }, cancellationToken: cancellation.Token);
            Assert.False(result.Succeeded); Assert.Null(result.OutputPath);
            Assert.Equal(original, File.ReadAllBytes(source)); Assert.Equal("keep", File.ReadAllText(output));
        });
    }

    private sealed class InvalidSigner : IPdfExternalSigner {
        public string Name => "Invalid test signer";
        public byte[] Sign(PdfExternalSignatureRequest request) => new byte[256];
    }
    private sealed class CommitStream(Action<byte[]> commit) : MemoryStream {
        private bool _closed;
        protected override void Dispose(bool disposing) { if (!_closed) { _closed = true; commit(ToArray()); } base.Dispose(disposing); }
    }
    private static byte[] CreatePdf() => PdfDocument.Create(document => {
        document.Page(page => page.Size(200, 300)); document.Page(page => page.Size(210, 310));
    }).ToBytes();
    private static X509Certificate2 CreateCertificate() {
        using var rsa = RSA.Create(2048);
        var request = new CertificateRequest("CN=OfficeIMO Signing Workflow Test", rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, true));
        return request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddDays(1));
    }
    private static async Task InDirectory(Func<string, Task> run) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-signing-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root); try { await run(root); } finally { Directory.Delete(root, true); }
    }
}
