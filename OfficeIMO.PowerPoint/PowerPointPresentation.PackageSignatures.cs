using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Security;

namespace OfficeIMO.PowerPoint;

public sealed partial class PowerPointPresentation {
    /// <summary>Inspects a saved PPTX/PPTM package through the shared dependency-light OPC signature engine.</summary>
    public static OfficePackageSignatureInfo InspectPackageSignatures(
        string filePath,
        OfficePackageSignatureInspectionOptions? options = null) =>
        OfficePackageSignatureService.Inspect(filePath, options);

    /// <summary>Validates a saved PPTX/PPTM package through an explicitly supplied security provider.</summary>
    public static OfficePackageSignatureValidationReport ValidatePackageSignatures(
        string filePath,
        IOfficeSecurityProvider securityProvider,
        OfficePackageSignatureValidationOptions? options = null) =>
        OfficePackageSignatureService.Validate(filePath, securityProvider, options);

    /// <summary>Creates and validates an OPC signature in a saved PPTX/PPTM package.</summary>
    public static OfficePackageSigningResult SignPackageSignature(
        string filePath,
        IOfficeSecurityProvider securityProvider,
        X509Certificate2 signingCertificate,
        OfficePackageSigningOptions? options = null) =>
        OfficePackageSignatureService.Sign(filePath, securityProvider, signingCertificate, options);

    /// <summary>Attempts OPC signature creation without throwing for an ordinary signing failure.</summary>
    public static OfficePackageSigningResult TrySignPackageSignature(
        string filePath,
        IOfficeSecurityProvider securityProvider,
        X509Certificate2 signingCertificate,
        OfficePackageSigningOptions? options = null) =>
        OfficePackageSignatureService.TrySign(filePath, securityProvider, signingCertificate, options);
}
