using OfficeIMO.Security;
using System.Collections.Generic;
using System.Security.Cryptography.X509Certificates;

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
        X509Certificate2 signingCertificate,
        OfficeVbaSigningOptions? options = null,
        IEnumerable<X509Certificate2>? certificateChain = null) =>
        OfficeVbaSignatureService.Sign(filePath, securityProvider, signingCertificate, options, certificateChain);

    /// <summary>Attempts shared atomic VBA signing and returns structured failure evidence.</summary>
    public static OfficeVbaSigningResult TrySignVbaProject(
        string filePath,
        IOfficeSecurityProvider securityProvider,
        X509Certificate2 signingCertificate,
        OfficeVbaSigningOptions? options = null,
        IEnumerable<X509Certificate2>? certificateChain = null) =>
        OfficeVbaSignatureService.TrySign(filePath, securityProvider, signingCertificate, options, certificateChain);
}
