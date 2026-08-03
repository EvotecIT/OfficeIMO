using OfficeIMO.Security;

namespace OfficeIMO.Word;

public partial class WordDocument {
    /// <summary>Inspects VBA signatures through the package-neutral core shared with Excel and PowerPoint.</summary>
    public static OfficeVbaSignatureInfo InspectVbaSignatures(
        string filePath,
        OfficeVbaSignatureInspectionOptions? options = null) =>
        OfficeVbaSignatureService.Inspect(filePath, options);

    /// <summary>Validates VBA signatures through the package-neutral core shared with Excel and PowerPoint.</summary>
    public static OfficeVbaSignatureValidationResult ValidateVbaSignatures(
        string filePath,
        IOfficeSecurityProvider securityProvider,
        OfficeVbaSignatureInspectionOptions? options = null) =>
        OfficeVbaSignatureService.Validate(filePath, securityProvider, options);

    /// <summary>Creates and atomically validates legacy, agile, and V3 VBA signatures through the shared owner.</summary>
    public static OfficeVbaSigningResult SignVbaProject(
        string filePath,
        IOfficeSecurityProvider securityProvider,
        string certificateThumbprint,
        OfficeVbaSigningOptions? options = null) =>
        OfficeVbaSignatureService.Sign(filePath, securityProvider, certificateThumbprint, options);

    /// <summary>Attempts shared atomic VBA signing and returns structured failure evidence.</summary>
    public static OfficeVbaSigningResult TrySignVbaProject(
        string filePath,
        IOfficeSecurityProvider securityProvider,
        string certificateThumbprint,
        OfficeVbaSigningOptions? options = null) =>
        OfficeVbaSignatureService.TrySign(filePath, securityProvider, certificateThumbprint, options);
}
