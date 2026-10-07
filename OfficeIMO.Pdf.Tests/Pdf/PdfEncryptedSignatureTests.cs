using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Pdf;
using OfficeIMO.Security;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfEncryptedSignatureTests {
    [Fact]
    public void EncryptedSigningHonorsOutputBudgetAndCancellationBeforeCallingSigner() {
        byte[] source = PdfDocument.Create(pdf => pdf.Content(c => c.Paragraph(p => p.Text("Protected signing content"))),
            new PdfOptions().SetEncryption(new PdfStandardEncryptionOptions("open") { OwnerPassword = "owner" })).ToBytes();
        PdfDocument document = PdfDocument.Load(source, new PdfLoadOptions { Password = "owner" });
        var signer = new UnexpectedSigner();
        Assert.Throws<InvalidDataException>(() => document.Security.SignExternal(signer,
            new PdfExternalSignatureOptions { MaxPreparedOutputBytes = source.LongLength }));
        using var cancellation = new System.Threading.CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => document.Security.SignExternal(signer,
            new PdfExternalSignatureOptions { CancellationToken = cancellation.Token }));
        Assert.Equal(0, signer.Calls);
        Assert.Equal(source, document.ToBytes());
    }

    [Theory]
    [InlineData(PdfStandardEncryptionAlgorithm.LegacyRc4)]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes128)]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes256)]
    public void EncryptedApprovalAndCountersignaturePreserveEncryptionAndVerifyDigests(PdfStandardEncryptionAlgorithm algorithm) {
        byte[] source = PdfDocument.Create(pdf => pdf.Content(c => c.Paragraph(p => p.Text("Protected signing content"))),
            new PdfOptions().SetEncryption(new PdfStandardEncryptionOptions("open") {
                OwnerPassword = "owner", Algorithm = algorithm,
                AllowedPermissions = PdfStandardPermissions.Print
            })).ToBytes();
        var ownerOptions = new PdfLoadOptions { Password = "owner" };
        using RSA rsa = RSA.Create(2048);
        var request = new CertificateRequest("CN=OfficeIMO encrypted signature test", rsa,
            HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        using X509Certificate2 certificate = request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-1), DateTimeOffset.UtcNow.AddDays(1));
        using var signer = new PdfCmsExternalSigner(OfficeSecurityProvider.Default, certificate);
        var verificationOptions = new CmsVerificationOptions();
        verificationOptions.CertificateValidation.ChainEvaluator = (_, _) => true;
        var verifier = new PdfCmsSignatureCryptographyProvider(OfficeSecurityProvider.Default, verificationOptions);
        byte[] current = source;
        for (int i = 0; i < 2; i++) {
            PdfDocument document = PdfDocument.Load(current, ownerOptions);
            PdfExternalSignatureCompletion completion = document.Security.SignExternal(signer, new PdfExternalSignatureOptions {
                FieldName = "Approval" + i, Name = "Protected signer", Reason = "Protected reason", Location = "Protected location",
                VisibleAppearance = i == 0 ? new PdfVisibleSignatureAppearanceOptions { Text = "Protected visible signature" } : null
            });
            byte[] signed = completion.Pdf;
            Assert.True(PdfReadDocument.Open(completion.ToDocument().ToBytes(), ownerOptions).Security.HasEncryption);
            Assert.Equal(current, signed.Take(current.Length).ToArray());
            PdfSignatureValidationReport report = PdfSignatureValidator.Validate(signed, verifier, ownerOptions);
            Assert.True(report.ObjectGraphParsed, report.ObjectGraphError);
            Assert.True(report.IsStructurallyValid);
            Assert.True(report.DigestVerified);
            Assert.True(report.MathematicalSignaturesVerified);
            Assert.Equal(i + 1, report.SignatureCount);
            PdfReadDocument user = PdfReadDocument.Open(signed, new PdfLoadOptions { Password = "open" });
            Assert.True(user.Security.HasEncryption);
            Assert.Equal(PdfStandardPermissions.Print, user.Security.AllowedStandardPermissions);
            Assert.Throws<PdfPasswordRequiredException>(() => PdfReadDocument.Open(signed));
            Assert.Throws<PdfInvalidPasswordException>(() => PdfReadDocument.Open(signed, new PdfLoadOptions { Password = "wrong" }));
            string raw = System.Text.Encoding.ASCII.GetString(signed);
            Assert.DoesNotContain("Protected signing content", raw);
            Assert.DoesNotContain("Protected reason", raw);
            Assert.Equal("Protected reason", report.Signatures.Last().Signature.Reason);
            current = signed;
        }
    }

    private sealed class UnexpectedSigner : IPdfExternalSigner {
        public string Name => "Unexpected signer";
        internal int Calls { get; private set; }
        public byte[] Sign(PdfExternalSignatureRequest request) {
            Calls++;
            throw new InvalidOperationException("The signer must not run after a preparation failure.");
        }
    }
}
