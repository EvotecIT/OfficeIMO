using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Security;

// The shared saved-package overloads require explicit options. Existing Word-specific
// overloads keep their defaults and richer evidence; static and instance validation
// also have distinct inputs. These intentional 3.6 overload families remain unambiguous.
[assembly: System.Diagnostics.CodeAnalysis.SuppressMessage("ApiDesign", "RS0027",
    Justification = "Distinct Word evidence and shared package options are intentional 3.6 overload contracts.",
    Scope = "member", Target = "~M:OfficeIMO.Word.WordDocument.ValidateSignatures(OfficeIMO.Security.IOfficeSecurityProvider,OfficeIMO.Word.WordSignatureValidationOptions)")]
[assembly: System.Diagnostics.CodeAnalysis.SuppressMessage("ApiDesign", "RS0027",
    Justification = "Distinct Word evidence and shared package options are intentional 3.6 overload contracts.",
    Scope = "member", Target = "~M:OfficeIMO.Word.WordDocument.SignPackage(System.String,OfficeIMO.Security.IOfficeSecurityProvider,System.Security.Cryptography.X509Certificates.X509Certificate2,OfficeIMO.Word.WordPackageSigningOptions)")]
[assembly: System.Diagnostics.CodeAnalysis.SuppressMessage("ApiDesign", "RS0027",
    Justification = "Distinct Word evidence and shared package options are intentional 3.6 overload contracts.",
    Scope = "member", Target = "~M:OfficeIMO.Word.WordDocument.TrySignPackage(System.String,OfficeIMO.Security.IOfficeSecurityProvider,System.Security.Cryptography.X509Certificates.X509Certificate2,OfficeIMO.Word.WordPackageSigningOptions)")]

[assembly: System.Diagnostics.CodeAnalysis.SuppressMessage("ApiDesign", "RS0026",
    Justification = "Static saved-package validation and instance validation have distinct inputs.",
    Scope = "member", Target = "~M:OfficeIMO.Word.WordDocument.ValidateSignatures(System.String,OfficeIMO.Security.IOfficeSecurityProvider,OfficeIMO.Security.OfficePackageSignatureValidationOptions)")]

namespace OfficeIMO.Word;

public partial class WordDocument {
    /// <summary>Inspects a saved DOCX/DOCM package through the shared dependency-light OPC signature engine.</summary>
    internal static OfficeIMO.Security.OfficePackageSignatureInfo InspectSignatures(
        string filePath,
        OfficePackageSignatureInspectionOptions? options = null) =>
        OfficePackageSignatureService.Inspect(filePath, options);

    /// <summary>Validates a saved DOCX/DOCM package through the shared optional-provider engine.</summary>
    public static OfficePackageSignatureValidationReport ValidateSignatures(
        string filePath,
        IOfficeSecurityProvider securityProvider,
        OfficePackageSignatureValidationOptions? options = null) =>
        OfficePackageSignatureService.Validate(filePath, securityProvider, options);

    /// <summary>Creates and validates an OPC signature through the shared cross-host engine.</summary>
    public static OfficePackageSigningResult SignPackage(
        string filePath,
        IOfficeSecurityProvider securityProvider,
        X509Certificate2 signingCertificate,
        OfficePackageSigningOptions? options) =>
        OfficePackageSignatureService.Sign(filePath, securityProvider, signingCertificate, options);

    /// <summary>Attempts shared OPC signature creation without throwing for an ordinary signing failure.</summary>
    public static OfficePackageSigningResult TrySignPackage(
        string filePath,
        IOfficeSecurityProvider securityProvider,
        X509Certificate2 signingCertificate,
        OfficePackageSigningOptions? options) =>
        OfficePackageSignatureService.TrySign(filePath, securityProvider, signingCertificate, options);
}
