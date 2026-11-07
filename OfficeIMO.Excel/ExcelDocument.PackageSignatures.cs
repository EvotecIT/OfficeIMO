using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Security;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    /// <summary>Inspects a saved XLSX/XLSM package through the shared dependency-light OPC signature engine.</summary>
    internal static OfficePackageSignatureInfo InspectPackageSignatures(
        string filePath,
        OfficePackageSignatureInspectionOptions? options = null) =>
        OfficePackageSignatureService.Inspect(filePath, options);

#pragma warning disable RS0026 // The required first parameter selects a file path or stream; both retain the default validation policy.
    /// <summary>Validates a saved XLSX/XLSM package through an explicitly supplied security provider.</summary>
    public static OfficePackageSignatureValidationReport ValidatePackageSignatures(
        string filePath,
        IOfficeSecurityProvider securityProvider,
        OfficePackageSignatureValidationOptions? options = null) =>
        OfficePackageSignatureService.Validate(filePath, securityProvider, options);

    /// <summary>Validates an encoded XLSX/XLSM package from a caller-owned stream.</summary>
    /// <param name="source">A readable stream containing the original package bytes.</param>
    /// <param name="securityProvider">The provider used for cryptographic and certificate validation.</param>
    /// <param name="options">Signature validation policy and resource limits, or <see langword="null"/> for defaults.</param>
    /// <returns>The package signature validation results.</returns>
    /// <remarks>
    /// The stream remains open. Seekable streams are read from the beginning and their original position is restored,
    /// including when validation fails. Non-seekable streams are read forward from their current position.
    /// The configured package byte limit is enforced while reading.
    /// </remarks>
    public static OfficePackageSignatureValidationReport ValidatePackageSignatures(
        Stream source,
        IOfficeSecurityProvider securityProvider,
        OfficePackageSignatureValidationOptions? options = null) =>
        OfficePackageSignatureService.Validate(source, securityProvider, options);
#pragma warning restore RS0026

    /// <summary>Creates and validates an OPC signature in a saved XLSX/XLSM package.</summary>
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
