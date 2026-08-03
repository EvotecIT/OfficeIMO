using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Security;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    /// <summary>Inspects a saved XLSX/XLSM package through the shared dependency-light OPC signature engine.</summary>
    public static OfficePackageSignatureInfo InspectPackageSignatures(
        string filePath,
        OfficePackageSignatureInspectionOptions? options = null) =>
        OfficePackageSignatureService.Inspect(filePath, options);

    /// <summary>Validates a saved XLSX/XLSM package through an explicitly supplied security provider.</summary>
    public static OfficePackageSignatureValidationReport ValidatePackageSignatures(
        string filePath,
        IOfficeSecurityProvider securityProvider,
        OfficePackageSignatureValidationOptions? options = null) =>
        OfficePackageSignatureService.Validate(filePath, securityProvider, options);

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
