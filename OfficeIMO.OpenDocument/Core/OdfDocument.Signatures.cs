using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Security;

namespace OfficeIMO.OpenDocument;

public abstract partial class OdfDocument {
    /// <summary>Validates META-INF/documentsignatures.xml and signed package-entry digests through an explicit provider.</summary>
    public static OfficeXmlPackageSignatureValidationReport ValidatePackageSignatures(
        string filePath,
        IOfficeSecurityProvider securityProvider,
        OfficeXmlPackageSignatureOptions? options = null) =>
        OfficeXmlPackageSignatureService.Validate(
            filePath, OfficeXmlPackageSignatureFormat.OpenDocument, securityProvider, options);

    /// <summary>Creates, validates, and atomically commits an ODF XML package signature.</summary>
    public static OfficeXmlPackageSigningResult SignPackage(
        string filePath,
        IOfficeSecurityProvider securityProvider,
        X509Certificate2 signingCertificate,
        OfficeXmlPackageSignatureOptions? options = null) =>
        OfficeXmlPackageSignatureService.Sign(
            filePath, OfficeXmlPackageSignatureFormat.OpenDocument,
            securityProvider, signingCertificate, options);

    /// <summary>Attempts atomic ODF XML signature creation and returns structured failure evidence.</summary>
    public static OfficeXmlPackageSigningResult TrySignPackage(
        string filePath,
        IOfficeSecurityProvider securityProvider,
        X509Certificate2 signingCertificate,
        OfficeXmlPackageSignatureOptions? options = null) =>
        OfficeXmlPackageSignatureService.TrySign(
            filePath, OfficeXmlPackageSignatureFormat.OpenDocument,
            securityProvider, signingCertificate, options);
}
