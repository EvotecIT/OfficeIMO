using OfficeIMO.Security;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    /// <summary>Inspects legacy, agile, and V3 VBA signature profiles in a saved XLSM/XLTM/XLAM package.</summary>
    public static OfficeVbaSignatureInfo InspectVbaSignatures(
        string filePath,
        OfficeVbaSignatureInspectionOptions? options = null) =>
        OfficeVbaSignatureService.Inspect(filePath, options);

    /// <summary>Validates VBA CMS, trust, timestamp, and Office SIP content binding.</summary>
    public static OfficeVbaSignatureValidationResult ValidateVbaSignatures(
        string filePath,
        IOfficeSecurityProvider securityProvider,
        OfficeVbaSignatureInspectionOptions? options = null) =>
        OfficeVbaSignatureService.Validate(filePath, securityProvider, options);

    /// <summary>Creates and atomically validates legacy, agile, and V3 VBA signatures.</summary>
    public static OfficeVbaSigningResult SignVbaProject(
        string filePath,
        IOfficeSecurityProvider securityProvider,
        string certificateThumbprint,
        OfficeVbaSigningOptions? options = null) =>
        OfficeVbaSignatureService.Sign(filePath, securityProvider, certificateThumbprint, options);

    /// <summary>Attempts atomic VBA signing and returns structured failure evidence.</summary>
    public static OfficeVbaSigningResult TrySignVbaProject(
        string filePath,
        IOfficeSecurityProvider securityProvider,
        string certificateThumbprint,
        OfficeVbaSigningOptions? options = null) =>
        OfficeVbaSignatureService.TrySign(filePath, securityProvider, certificateThumbprint, options);
}
