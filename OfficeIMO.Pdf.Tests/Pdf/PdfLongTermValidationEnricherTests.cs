#if NET8_0_OR_GREATER
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Pdf;
using OfficeIMO.Security;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfLongTermValidationEnricherTests {
    [Fact]
    public void EnrichAppendsDssVriAndPreservesVerifiedSignature() {
        using X509Certificate2 certificate = CreateSigningCertificate();
        byte[] signedPdf = CreateSignedPdf(certificate);
        var provider = CreateProvider();
        PdfSignatureValidationResult signature = Assert.Single(PdfSignatureValidator.Validate(signedPdf, provider).Signatures);
        var evidence = new PdfLongTermValidationEvidence(
            signature.Signature.ObjectNumber,
            certificates: new[] { certificate.RawData },
            ocspResponses: new[] { new byte[] { 0x30, 0x03, 0x0A, 0x01, 0x00 } });

        PdfLongTermValidationEnrichmentResult result = PdfLongTermValidationEnricher.Enrich(signedPdf, evidence, provider);

        byte[] signatureValue = TrimDerContainer(signature.Signature.ContentsBytes!);
        string expectedVriKey = Convert.ToHexString(SHA1.HashData(signatureValue));

        Assert.True(result.IsVerifiedAppendOnlyEnrichment);
        Assert.True(result.MutationReport.OriginalBytesArePrefix);
        Assert.True(result.MutationReport.RevisionChainExtended);
        Assert.True(result.MutationReport.AllExistingSignaturesArePreserved);
        Assert.Equal(PdfMutationExecutionMode.AppendOnly, result.MutationReport.MutationPlan.ExecutionMode);
        Assert.Contains(PdfMutationProof.LongTermValidationReadback, result.MutationReport.MutationPlan.RequiredProofs);
        Assert.Equal(40, result.VriKey.Length);
        Assert.Equal(result.VriKey.ToUpperInvariant(), result.VriKey);
        Assert.Equal(expectedVriKey, result.VriKey);
        Assert.True(result.ValidationAfter.HasOfflineLongTermValidationReadiness);
        Assert.True(result.ValidationAfter.MathematicalSignaturesVerified);
        Assert.True(result.ValidationAfter.DigestVerified);
        Assert.Contains(result.VriKey, result.ValidationAfter.Security.DocumentSecurityStore.VriKeys);
        Assert.Single(result.CertificateObjectNumbers);
        Assert.Single(result.OcspObjectNumbers);
        Assert.Empty(result.CrlObjectNumbers);
    }

    [Fact]
    public void RepeatedEnrichmentRetainsPreviousDssEvidence() {
        using X509Certificate2 certificate = CreateSigningCertificate();
        byte[] signedPdf = CreateSignedPdf(certificate);
        var provider = CreateProvider();
        int signatureObjectNumber = Assert.Single(PdfSignatureValidator.Validate(signedPdf, provider).Signatures).Signature.ObjectNumber;
        var firstEvidence = new PdfLongTermValidationEvidence(signatureObjectNumber, certificates: new[] { certificate.RawData });
        PdfLongTermValidationEnrichmentResult first = PdfLongTermValidationEnricher.Enrich(signedPdf, firstEvidence, provider);
        var secondEvidence = new PdfLongTermValidationEvidence(
            signatureObjectNumber,
            ocspResponses: new[] { new byte[] { 0x30, 0x03, 0x0A, 0x01, 0x00 } });

        PdfLongTermValidationEnrichmentResult second = PdfLongTermValidationEnricher.Enrich(first.Pdf, secondEvidence, provider);

        Assert.True(second.IsVerifiedAppendOnlyEnrichment);
        Assert.Contains(first.CertificateObjectNumbers[0], second.ValidationAfter.Security.DocumentSecurityStore.CertificateObjectNumbers);
        Assert.Contains(first.CertificateObjectNumbers[0], second.ValidationAfter.Security.DocumentSecurityStore.VriCertificateObjectNumbers);
        Assert.Contains(second.OcspObjectNumbers[0], second.ValidationAfter.Security.DocumentSecurityStore.OcspObjectNumbers);
        Assert.Contains(second.OcspObjectNumbers[0], second.ValidationAfter.Security.DocumentSecurityStore.VriOcspObjectNumbers);
        Assert.Equal(first.VriKey, second.VriKey);
        Assert.Single(second.ValidationAfter.Security.DocumentSecurityStore.VriKeys);
    }

    [Fact]
    public void EnrichmentRejectsSignatureWhoseSignedBytesWereTampered() {
        using X509Certificate2 certificate = CreateSigningCertificate();
        byte[] signedPdf = CreateSignedPdf(certificate);
        var provider = CreateProvider();
        int signatureObjectNumber = Assert.Single(PdfSignatureValidator.Validate(signedPdf, provider).Signatures).Signature.ObjectNumber;
        signedPdf[7] = signedPdf[7] == (byte)'4' ? (byte)'5' : (byte)'4';
        var evidence = new PdfLongTermValidationEvidence(signatureObjectNumber, certificates: new[] { certificate.RawData });

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
            () => PdfLongTermValidationEnricher.Enrich(signedPdf, evidence, provider));

        Assert.Contains("valid signature math", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void PlannerBlocksDssVriEnrichmentForUnsignedPdf() {
        byte[] unsignedPdf = PdfDocument.Create().Paragraph(p => p.Text("Unsigned")).ToBytes();

        PdfMutationPlan plan = PdfMutationPlanner.Plan(unsignedPdf, PdfMutationOperation.EnrichLongTermValidation);

        Assert.False(plan.CanExecute);
        Assert.Contains("AppendOnly.Unsigned", plan.BlockerCodes);
        Assert.Contains("AppendOnly.ActionBlocked.LongTermValidation", plan.BlockerCodes);
    }

    private static byte[] CreateSignedPdf(X509Certificate2 certificate) {
        byte[] source = PdfDocument.Create().Paragraph(p => p.Text("LTV source")).ToBytes();
        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            source,
            new PdfExternalSignatureOptions {
                Profile = PdfSignatureProfile.Approval,
                FieldName = "LtvSignature",
                ReservedSignatureContentsBytes = 8192
            });
        using var signer = new PdfCmsExternalSigner(certificate);
        return PdfIncrementalUpdater.ApplyExternalSignature(
            preparation,
            signer.Sign(new PdfExternalSignatureRequest(preparation)));
    }

    private static PdfCmsSignatureCryptographyProvider CreateProvider() {
        var options = new CmsVerificationOptions();
        options.CertificateValidation.ChainEvaluator = static (_, _) => true;
        return new PdfCmsSignatureCryptographyProvider(options);
    }

    private static byte[] TrimDerContainer(byte[] value) {
        int offset = 1;
        int firstLength = value[offset++];
        int contentLength = firstLength;
        if ((firstLength & 0x80) != 0) {
            int lengthBytes = firstLength & 0x7F;
            contentLength = 0;
            for (int index = 0; index < lengthBytes; index++) contentLength = (contentLength << 8) | value[offset++];
        }

        var result = new byte[offset + contentLength];
        Buffer.BlockCopy(value, 0, result, 0, result.Length);
        return result;
    }

    private static X509Certificate2 CreateSigningCertificate() {
        using RSA rsa = RSA.Create(2048);
        var request = new CertificateRequest(
            "CN=OfficeIMO PDF LTV Test",
            rsa,
            HashAlgorithmName.SHA256,
            RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, critical: true));
        return request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddDays(1));
    }
}
#endif
