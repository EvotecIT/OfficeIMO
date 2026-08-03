using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Security;

namespace OfficeIMO.Epub;

public sealed partial class EpubDocument {
    /// <summary>Validates META-INF/signatures.xml and signed EPUB entry digests through an explicit provider.</summary>
    public static OfficeXmlPackageSignatureValidationReport ValidatePackageSignatures(
        string filePath,
        IOfficeSecurityProvider securityProvider,
        OfficeXmlPackageSignatureOptions? options = null) =>
        OfficeXmlPackageSignatureService.Validate(
            filePath, OfficeXmlPackageSignatureFormat.Epub, securityProvider, options);

    /// <summary>Creates, validates, and atomically commits an EPUB XML package signature.</summary>
    public static OfficeXmlPackageSigningResult SignPackage(
        string filePath,
        IOfficeSecurityProvider securityProvider,
        X509Certificate2 signingCertificate,
        OfficeXmlPackageSignatureOptions? options = null) =>
        OfficeXmlPackageSignatureService.Sign(
            filePath, OfficeXmlPackageSignatureFormat.Epub,
            securityProvider, signingCertificate, options);

    /// <summary>Attempts atomic EPUB XML signature creation and returns structured failure evidence.</summary>
    public static OfficeXmlPackageSigningResult TrySignPackage(
        string filePath,
        IOfficeSecurityProvider securityProvider,
        X509Certificate2 signingCertificate,
        OfficeXmlPackageSignatureOptions? options = null) =>
        OfficeXmlPackageSignatureService.TrySign(
            filePath, OfficeXmlPackageSignatureFormat.Epub,
            securityProvider, signingCertificate, options);
}
