using System.Security.Cryptography;
using System.Security.Cryptography.Pkcs;
using System.Security.Cryptography.X509Certificates;

namespace OfficeIMO.Security.Benchmarks;

internal static class SecurityCmsBenchmarkValidation {
    private const string Sha256Oid = "2.16.840.1.101.3.4.2.1";
    internal static readonly CmsSigningOptions OfficeSigningOptions = new() {
        DigestAlgorithm = HashAlgorithmName.SHA256,
        IncludeSigningTime = false,
        IncludeCertificateChain = false
    };
    internal static readonly CmsVerificationOptions OfficeVerificationOptions = CreateVerificationOptions();

    private static CmsVerificationOptions CreateVerificationOptions() {
        var options = new CmsVerificationOptions { ValidateTimestamps = false };
        options.CertificateValidation.ValidateChain = false;
        options.CertificateValidation.DisableCertificateDownloads = true;
        options.CertificateValidation.RevocationMode = X509RevocationMode.NoCheck;
        return options;
    }

    internal static byte[] SignOffice(SecurityCmsBenchmarkFixture fixture) =>
        CmsSignedDataSigner.SignDetached(
            fixture.Content,
            fixture.Certificate,
            OfficeSigningOptions);

    internal static byte[] SignPlatform(SecurityCmsBenchmarkFixture fixture) {
        var signed = new SignedCms(new ContentInfo(fixture.Content), detached: true);
        var signer = new CmsSigner(SubjectIdentifierType.IssuerAndSerialNumber, fixture.Certificate) {
            DigestAlgorithm = new Oid(Sha256Oid),
            IncludeOption = X509IncludeOption.EndCertOnly
        };
        signed.ComputeSignature(signer, silent: true);
        return signed.Encode();
    }

    internal static CmsVerificationResult VerifyOffice(byte[] signature, byte[] content) =>
        CmsSignedDataVerifier.VerifyDetached(signature, content, OfficeVerificationOptions);

    internal static PlatformCmsVerificationSnapshot VerifyPlatform(byte[] signature, byte[] content) {
        var signed = new SignedCms(new ContentInfo(content), detached: true);
        signed.Decode(signature);
        signed.CheckSignature(verifySignatureOnly: true);
        if (signed.SignerInfos.Count != 1) {
            throw new InvalidOperationException("The CMS object must contain exactly one signer.");
        }

        SignerInfo signer = signed.SignerInfos[0];
        X509Certificate2 certificate = signer.Certificate
            ?? throw new InvalidOperationException("The CMS signer certificate is missing.");
        X509KeyUsageExtension? keyUsage = certificate.Extensions
            .OfType<X509KeyUsageExtension>()
            .FirstOrDefault();
        bool usageAccepted = keyUsage == null ||
            (keyUsage.KeyUsages & (X509KeyUsageFlags.DigitalSignature | X509KeyUsageFlags.NonRepudiation)) != 0;
        (bool signedAttributesAccepted, int signedAttributeBytes) = InspectPlatformSignedAttributes(signer);
        return new PlatformCmsVerificationSnapshot(
            signed.Detached,
            signed.SignerInfos.Count,
            certificate.RawData,
            certificate.Subject,
            certificate.Issuer,
            certificate.SerialNumber,
            certificate.Thumbprint,
            signer.DigestAlgorithm.Value ?? string.Empty,
            signer.SignatureAlgorithm.Value ?? string.Empty,
            usageAccepted,
            signedAttributesAccepted,
            signedAttributeBytes);
    }

    private static (bool Accepted, int EncodedBytes) InspectPlatformSignedAttributes(SignerInfo signer) {
        const string ContentTypeAttributeOid = "1.2.840.113549.1.9.3";
        const string MessageDigestAttributeOid = "1.2.840.113549.1.9.4";
        const string CounterSignatureAttributeOid = "1.2.840.113549.1.9.6";
        const string AlgorithmProtectionAttributeOid = "1.2.840.113549.1.9.52";
        if (signer.SignedAttributes.Count == 0) return (true, 0);
        int contentTypes = 0;
        int messageDigests = 0;
        int algorithmProtections = 0;
        int encodedBytes = 0;
        bool accepted = true;
        foreach (CryptographicAttributeObject attribute in signer.SignedAttributes) {
            string oid = attribute.Oid.Value ?? string.Empty;
            if (oid == ContentTypeAttributeOid) contentTypes++;
            else if (oid == MessageDigestAttributeOid) messageDigests++;
            else if (oid == AlgorithmProtectionAttributeOid) algorithmProtections++;
            else if (oid == CounterSignatureAttributeOid) accepted = false;
            if (attribute.Values.Count != 1 &&
                (oid == ContentTypeAttributeOid || oid == MessageDigestAttributeOid ||
                 oid == AlgorithmProtectionAttributeOid)) {
                accepted = false;
            }
            foreach (AsnEncodedData value in attribute.Values) encodedBytes += value.RawData.Length;
        }
        accepted &= contentTypes == 1 && messageDigests == 1 && algorithmProtections <= 1;
        return (accepted, encodedBytes);
    }

    internal static SecurityCmsValidationSnapshot Validate(SecurityCmsBenchmarkFixture fixture) {
        byte[] officeSignature = SignOffice(fixture);
        byte[] platformSignature = SignPlatform(fixture);

        CmsVerificationResult officeOnOffice = VerifyOffice(officeSignature, fixture.Content);
        CmsVerificationResult officeOnPlatform = VerifyOffice(platformSignature, fixture.Content);
        PlatformCmsVerificationSnapshot platformOnOffice = VerifyPlatform(officeSignature, fixture.Content);
        PlatformCmsVerificationSnapshot platformOnPlatform = VerifyPlatform(platformSignature, fixture.Content);
        ValidateOfficeResult(officeOnOffice, "OfficeIMO signature");
        ValidateOfficeResult(officeOnPlatform, "platform signature");
        if (platformOnOffice.SignerCount != 1 || platformOnPlatform.SignerCount != 1 ||
            !platformOnOffice.UsageAccepted || !platformOnPlatform.UsageAccepted ||
            !platformOnOffice.SignedAttributesAccepted || !platformOnPlatform.SignedAttributesAccepted) {
            throw new InvalidOperationException(
                "Each CMS output must contain exactly one verifiable signer. " +
                $"Office attributes={platformOnOffice.SignedAttributesAccepted}/{platformOnOffice.SignedAttributeBytes}; " +
                $"platform attributes={platformOnPlatform.SignedAttributesAccepted}/{platformOnPlatform.SignedAttributeBytes}.");
        }

        ValidatePlatformShape(officeSignature, fixture.Content, "OfficeIMO signature");
        ValidatePlatformShape(platformSignature, fixture.Content, "platform signature");
        ValidateTamperRejection(officeSignature, platformSignature, fixture.Content);

        return new SecurityCmsValidationSnapshot(
            fixture.Scale,
            fixture.Content.Length,
            officeSignature.Length,
            platformSignature.Length,
            officeSignature,
            platformSignature);
    }

    private static void ValidateOfficeResult(CmsVerificationResult result, string label) {
        if (!result.Parsed || !result.IsDetached || !result.IsCryptographicallyValid || result.Signers.Count != 1) {
            string findings = string.Join(
                " | ",
                result.Findings.Concat(result.Signers.SelectMany(static signer => signer.Findings))
                    .Select(static finding => finding.Code + ": " + finding.Message));
            throw new InvalidOperationException(label + " failed OfficeIMO verification: " + findings);
        }
        CmsSignerVerificationResult signer = result.Signers[0];
        if (!string.Equals(signer.DigestAlgorithmOid, Sha256Oid, StringComparison.Ordinal)) {
            throw new InvalidOperationException(label + " does not use SHA-256.");
        }
    }

    private static void ValidatePlatformShape(byte[] signature, byte[] content, string label) {
        var signed = new SignedCms(new ContentInfo(content), detached: true);
        signed.Decode(signature);
        if (!signed.Detached || signed.SignerInfos.Count != 1 || signed.Certificates.Count != 1) {
            throw new InvalidOperationException(label + " is not an equivalent one-signer detached CMS object.");
        }
        if (!string.Equals(signed.SignerInfos[0].DigestAlgorithm.Value, Sha256Oid, StringComparison.Ordinal)) {
            throw new InvalidOperationException(label + " does not use SHA-256.");
        }
    }

    private static void ValidateTamperRejection(byte[] officeSignature, byte[] platformSignature, byte[] content) {
        byte[] tampered = (byte[])content.Clone();
        tampered[tampered.Length / 2] ^= 0x5A;
        if (VerifyOffice(officeSignature, tampered).IsCryptographicallyValid
            || VerifyOffice(platformSignature, tampered).IsCryptographicallyValid) {
            throw new InvalidOperationException("OfficeIMO accepted tampered detached content.");
        }
        ExpectPlatformRejection(officeSignature, tampered);
        ExpectPlatformRejection(platformSignature, tampered);
    }

    private static void ExpectPlatformRejection(byte[] signature, byte[] tampered) {
        try {
            VerifyPlatform(signature, tampered);
        } catch (CryptographicException) {
            return;
        }
        throw new InvalidOperationException("The platform verifier accepted tampered detached content.");
    }
}

internal sealed record SecurityCmsValidationSnapshot(
    string Scale,
    int ContentBytes,
    int OfficeSignatureBytes,
    int PlatformSignatureBytes,
    byte[] OfficeSignature,
    byte[] PlatformSignature);

public sealed record PlatformCmsVerificationSnapshot(
    bool Detached,
    int SignerCount,
    byte[] CertificateBytes,
    string Subject,
    string Issuer,
    string SerialNumber,
    string Thumbprint,
    string DigestAlgorithmOid,
    string SignatureAlgorithmOid,
    bool UsageAccepted,
    bool SignedAttributesAccepted,
    int SignedAttributeBytes);
