using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Pdf;
using OfficeIMO.Security;
using Org.BouncyCastle.Ocsp;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

/// <summary>Revocation, proof-of-existence and PAdES level checks against a local test PKI with live responders.</summary>
public sealed class PdfSignatureRevocationTests {
    private static readonly DateTime Now = DateTime.UtcNow;

    [Fact]
    public async Task FetchedOcspShowsRevokedSignerAsInvalid() {
        using var pki = new PdfTestPki(Now);
        using X509Certificate2 signer = pki.IssueSigner("Revoked by OCSP", Now.AddDays(-30), Now.AddDays(30));
        pki.Revoke(signer, Now.AddHours(-2));
        byte[] signed = pki.Sign(signer, timestampTime: Now.AddHours(-1));
        PdfCmsSignatureCryptographyProvider provider = pki.Provider();
        using var fetcher = InternalFetcher();

        PdfLongTermValidationEnrichmentResult enriched = await PdfLongTermValidationEnricher.FetchAndEnrichAsync(
            signed, SignatureObject(signed, provider), provider, fetcher);

        PdfSignatureValidationResult signature = Assert.Single(enriched.ValidationAfter.Signatures);
        Assert.True(pki.OcspRequests > 0);
        Assert.Empty(enriched.FetchFindings);
        Assert.Equal(PdfCryptographicValidationStatus.Invalid, signature.CryptographicResult!.RevocationStatus);
        Assert.Contains(signature.Findings, finding => finding.Code == "RevocationCertificateRevoked");
        Assert.Equal(PdfPadesLevel.BaselineT, signature.PadesLevel);
        Assert.False(enriched.ValidationAfter.CryptographicTrustVerified);
    }

    [Fact]
    public void OnlineCrlShowsRevokedSignerAsInvalidAndIsFetchedOncePerRun() {
        using var pki = new PdfTestPki(Now, ocsp: false);
        using X509Certificate2 signer = pki.IssueSigner("Revoked by CRL", Now.AddDays(-30), Now.AddDays(30));
        pki.Revoke(signer, Now.AddHours(-2));
        byte[] signed = pki.Sign(signer, timestampTime: Now.AddHours(-1));
        using var fetcher = InternalFetcher();
        var provider = new PdfCmsSignatureCryptographyProvider(OfficeSecurityProvider.Default, pki.VerificationOptions()) {
            RevocationPolicy = CertificateRevocationPolicy.EmbeddedThenOnline,
            RevocationFetcher = fetcher
        };

        PdfSignatureValidationReport first = PdfDocument.Load(signed).Security.ValidateSignatures(provider);
        PdfSignatureValidationReport second = PdfDocument.Load(signed).Security.ValidateSignatures(provider);

        foreach (PdfSignatureValidationReport report in new[] { first, second }) {
            PdfSignatureValidationResult signature = Assert.Single(report.Signatures);
            Assert.Equal(PdfCryptographicValidationStatus.Invalid, signature.CryptographicResult!.RevocationStatus);
            Assert.Contains(signature.Findings, finding => finding.Code == "RevocationCertificateRevoked" &&
                finding.Message.Contains("fetched online", StringComparison.Ordinal));
        }
        Assert.Equal(1, pki.CrlRequests);
        Assert.Equal(0, pki.OcspRequests);
    }

    [Fact]
    public void CrlScopedToAuthoritiesDoesNotProveSignerIsGood() {
        using var pki = new PdfTestPki(Now, ocsp: false, crlCoversOnlyAuthorities: true);
        using X509Certificate2 signer = pki.IssueSigner("Out of CRL scope", Now.AddDays(-30), Now.AddDays(30));
        byte[] signed = pki.Sign(signer, timestampTime: Now.AddHours(-1));
        using var fetcher = InternalFetcher();
        var provider = new PdfCmsSignatureCryptographyProvider(OfficeSecurityProvider.Default, pki.VerificationOptions()) {
            RevocationPolicy = CertificateRevocationPolicy.Online,
            RevocationFetcher = fetcher
        };

        PdfSignatureValidationResult signature = Assert.Single(PdfDocument.Load(signed).Security.ValidateSignatures(provider).Signatures);

        Assert.Equal(PdfCryptographicValidationStatus.Indeterminate, signature.CryptographicResult!.RevocationStatus);
        Assert.Contains(signature.Findings, finding => finding.Code == "RevocationDataUnusable" &&
            finding.Message.Contains("covers only CA certificates", StringComparison.Ordinal));
    }

    [Fact]
    public async Task ExpiredSignerRemainsValidWhenDocumentTimestampProvesTheValidationData() {
        // The signer and the OCSP responder certificates have both expired by now. The signature time-stamp
        // proves when the signature existed; the document time-stamp proves the DSS data existed while the
        // responder certificate was still valid.
        using var pki = new PdfTestPki(Now, responderNotAfter: Now.AddDays(-2));
        using X509Certificate2 signer = pki.IssueSigner("Expired signer", Now.AddDays(-10), Now.AddDays(-1));
        byte[] signed = pki.Sign(signer, timestampTime: Now.AddDays(-5));
        PdfCmsSignatureCryptographyProvider provider = pki.Provider();
        pki.ResponderTime = Now.AddDays(-4);
        using var fetcher = InternalFetcher();
        byte[] longTerm = (await PdfLongTermValidationEnricher.FetchAndEnrichAsync(
            signed, SignatureObject(signed, provider), provider, fetcher)).Pdf;

        byte[] archived = pki.AddDocumentTimestamp(longTerm, Now.AddDays(-3));
        PdfSignatureValidationReport report = PdfDocument.Load(archived).Security.ValidateSignatures(provider);

        PdfSignatureValidationResult signature = Assert.Single(report.Signatures, item => !item.Signature.IsDocumentTimestamp);
        PdfSignatureCryptographicResult result = signature.CryptographicResult!;
        Assert.True(result.IsMathematicallyValid);
        Assert.Equal(PdfCryptographicValidationStatus.Valid, result.CertificateChainStatus);
        Assert.Equal(PdfCryptographicValidationStatus.Valid, result.TimestampStatus);
        Assert.Equal(PdfCryptographicValidationStatus.Valid, result.RevocationStatus);
        Assert.Equal(PdfPadesLevel.BaselineLTA, signature.PadesLevel);
        Assert.DoesNotContain(signature.Findings, finding => finding.Code.StartsWith("pades-baseline-", StringComparison.Ordinal));

        // Without the document time-stamp nothing proves the OCSP response predates the responder's expiry.
        PdfSignatureValidationResult withoutArchive = Assert.Single(
            PdfDocument.Load(longTerm).Security.ValidateSignatures(provider).Signatures);
        Assert.Equal(PdfCryptographicValidationStatus.Indeterminate, withoutArchive.CryptographicResult!.RevocationStatus);
        Assert.Contains(withoutArchive.Findings, finding => finding.Code == "RevocationDataUnusable" &&
            finding.Message.Contains("expired", StringComparison.Ordinal));
        Assert.Equal(PdfPadesLevel.BaselineT, withoutArchive.PadesLevel);
    }

    [Fact]
    public async Task TamperedDssOcspResponseDoesNotSettleRevocation() {
        using var pki = new PdfTestPki(Now);
        using X509Certificate2 signer = pki.IssueSigner("Tampered DSS", Now.AddDays(-30), Now.AddDays(30));
        byte[] signed = pki.Sign(signer, timestampTime: Now.AddHours(-1));
        PdfCmsSignatureCryptographyProvider provider = pki.Provider();
        using var fetcher = InternalFetcher();
        PdfLongTermValidationEnrichmentResult enriched = await PdfLongTermValidationEnricher.FetchAndEnrichAsync(
            signed, SignatureObject(signed, provider), provider, fetcher);
        Assert.Equal(PdfPadesLevel.BaselineLT, Assert.Single(enriched.ValidationAfter.Signatures).PadesLevel);

        byte[] tampered = enriched.Pdf;
        foreach (byte[] response in enriched.Evidence.OcspResponses) {
            byte[] signatureValue = ((BasicOcspResp)new OcspResp(response).GetResponseObject()).GetSignature();
            int offset = IndexOf(tampered, signatureValue);
            Assert.True(offset > signed.Length, "The OCSP response must be stored after the signed revision.");
            tampered[offset + signatureValue.Length / 2] ^= 0x5A;
        }

        PdfSignatureValidationResult signature = Assert.Single(PdfDocument.Load(tampered).Security.ValidateSignatures(provider).Signatures);
        Assert.True(signature.CryptographicResult!.IsMathematicallyValid);
        Assert.Equal(PdfCryptographicValidationStatus.Indeterminate, signature.CryptographicResult.RevocationStatus);
        Assert.Contains(signature.Findings, finding => finding.Code == "RevocationDataUnusable" &&
            finding.Message.Contains("signature", StringComparison.Ordinal));
        Assert.Equal(PdfPadesLevel.BaselineT, signature.PadesLevel);
        Assert.Contains(signature.Findings, finding => finding.Code == "pades-baseline-lt-not-reached");
    }

    [Fact]
    public void MissingRevocationDataIsIndeterminateWithReason() {
        using var pki = new PdfTestPki(Now);
        using X509Certificate2 signer = pki.IssueSigner("No revocation data", Now.AddDays(-30), Now.AddDays(30));
        byte[] signed = pki.Sign(signer, timestampTime: Now.AddHours(-1));

        PdfSignatureValidationReport report = PdfDocument.Load(signed).Security.ValidateSignatures(pki.Provider());

        PdfSignatureValidationResult signature = Assert.Single(report.Signatures);
        Assert.Equal(PdfCryptographicValidationStatus.Indeterminate, signature.CryptographicResult!.RevocationStatus);
        Assert.Contains(signature.Findings, finding => finding.Code == "RevocationDataMissing" &&
            finding.Message.Contains("No revocation data", StringComparison.Ordinal));
        Assert.Equal(PdfPadesLevel.BaselineT, signature.PadesLevel);
        Assert.Contains(signature.Findings, finding => finding.Code == "pades-baseline-lt-not-reached");
        Assert.True(report.CryptographicTrustVerified);
        Assert.False(report.RevocationChecked);
        Assert.Equal(0, pki.OcspRequests + pki.CrlRequests);
    }

    [Fact]
    public async Task DefaultFetcherRefusesLoopbackResponders() {
        using var pki = new PdfTestPki(Now);
        using X509Certificate2 signer = pki.IssueSigner("Loopback responder", Now.AddDays(-30), Now.AddDays(30));
        using var fetcher = new CertificateRevocationFetcher();

        CertificateRevocationEvidence evidence = await fetcher.FetchAsync(signer, pki.Root);

        Assert.Empty(evidence.OcspResponses);
        Assert.Empty(evidence.CertificateRevocationLists);
        Assert.Contains(evidence.Findings, finding => finding.Code == "RevocationFetchFailed" &&
            finding.Message.Contains("non-public", StringComparison.Ordinal));
        Assert.Equal(0, pki.OcspRequests + pki.CrlRequests);
    }

    private static CertificateRevocationFetcher InternalFetcher() =>
        new(new CertificateRevocationFetcherOptions { AllowNonPublicAddresses = true });

    private static int SignatureObject(byte[] pdf, IPdfSignatureCryptographyProvider provider) =>
        Assert.Single(PdfDocument.Load(pdf).Security.ValidateSignatures(provider).Signatures).Signature.ObjectNumber;

    private static int IndexOf(byte[] data, byte[] value) {
        for (int index = 0; index <= data.Length - value.Length; index++) {
            int match = 0;
            while (match < value.Length && data[index + match] == value[match]) match++;
            if (match == value.Length) return index;
        }
        return -1;
    }
}
