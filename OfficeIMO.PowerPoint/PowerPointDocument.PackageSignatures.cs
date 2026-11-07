using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Security;

namespace OfficeIMO.PowerPoint;

public sealed partial class PowerPointDocument {
    /// <summary>Inspects a saved PPTX/PPTM package through the shared dependency-light OPC signature engine.</summary>
    internal static OfficePackageSignatureInfo InspectSignatures(
        string filePath,
        OfficePackageSignatureInspectionOptions? options = null) =>
        OfficePackageSignatureService.Inspect(filePath, options);

#pragma warning disable RS0026 // The required first parameter selects a file path or stream; both retain the default validation policy.
    /// <summary>Validates a saved PPTX/PPTM package through an explicitly supplied security provider.</summary>
    public static OfficePackageSignatureValidationReport ValidateSignatures(
        string filePath,
        IOfficeSecurityProvider securityProvider,
        OfficePackageSignatureValidationOptions? options = null) =>
        OfficePackageSignatureService.Validate(filePath, securityProvider, options);

    /// <summary>Validates an encoded PPTX/PPTM package from a caller-owned stream.</summary>
    /// <param name="source">A readable stream containing the original package bytes.</param>
    /// <param name="securityProvider">The provider used for cryptographic and certificate validation.</param>
    /// <param name="options">Signature validation policy and resource limits, or <see langword="null"/> for defaults.</param>
    /// <returns>The package signature validation results.</returns>
    /// <remarks>
    /// The stream remains open. Seekable streams are read from the beginning and their original position is restored,
    /// including when validation fails. Non-seekable streams are read forward from their current position.
    /// The configured package byte limit is enforced while reading.
    /// </remarks>
    public static OfficePackageSignatureValidationReport ValidateSignatures(
        Stream source,
        IOfficeSecurityProvider securityProvider,
        OfficePackageSignatureValidationOptions? options = null) =>
        OfficePackageSignatureService.Validate(source, securityProvider, options);
#pragma warning restore RS0026

    /// <summary>Creates and validates an OPC signature in a saved PPTX/PPTM package.</summary>
    public static OfficePackageSigningResult SignPackage(
        string filePath,
        IOfficeSecurityProvider securityProvider,
        X509Certificate2 signingCertificate,
        OfficePackageSigningOptions? options = null) =>
        OfficePackageSignatureService.Sign(filePath, securityProvider, signingCertificate, options);

    /// <summary>Attempts OPC signature creation without throwing for an ordinary signing failure.</summary>
    public static OfficePackageSigningResult TrySignPackage(
        string filePath,
        IOfficeSecurityProvider securityProvider,
        X509Certificate2 signingCertificate,
        OfficePackageSigningOptions? options = null) =>
        OfficePackageSignatureService.TrySign(filePath, securityProvider, signingCertificate, options);
}
