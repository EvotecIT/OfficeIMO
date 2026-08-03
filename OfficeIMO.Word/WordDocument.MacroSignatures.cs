using OfficeIMO.Security;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>
        /// Inspects VBA project and signature-part metadata cross-platform without claiming that the
        /// signatures bind to the current macro project. Use <see cref="ValidateMacroProjectSignature(string, IOfficeSecurityProvider, WordMacroProjectSignatureValidationOptions?)"/>
        /// for Microsoft Office SIP content-binding validation.
        /// </summary>
        /// <param name="filePath">Path to a saved DOCM or DOTM package.</param>
        /// <param name="options">Optional package, signature, CMS, trust, revocation, and timestamp policy.</param>
        public static WordMacroProjectSignatureInfo InspectMacroProjectSignatures(
            string filePath,
            WordMacroProjectSignatureInspectionOptions? options = null) =>
            WordMacroProjectSignatureInspector.Inspect(filePath, options);

        /// <summary>
        /// Inspects VBA project and signature-part metadata and evaluates CMS, certificate, revocation,
        /// and timestamp evidence through an explicitly supplied security provider.
        /// </summary>
        /// <param name="filePath">Path to a saved DOCM or DOTM package.</param>
        /// <param name="securityProvider">Cryptographic provider supplied by the optional OfficeIMO.Security package.</param>
        /// <param name="options">Optional package, signature, CMS, trust, revocation, and timestamp policy.</param>
        public static WordMacroProjectSignatureInfo InspectMacroProjectSignatures(
            string filePath,
            IOfficeSecurityProvider securityProvider,
            WordMacroProjectSignatureInspectionOptions? options = null) =>
            WordMacroProjectSignatureInspector.Inspect(filePath, securityProvider, options);

        /// <summary>
        /// Validates the highest-precedence VBA signature against the macro project through Microsoft's
        /// registered Office SIP, then applies caller CMS, certificate-chain, revocation, and timestamp policy.
        /// </summary>
        /// <param name="filePath">Path to a saved DOCM or DOTM package.</param>
        /// <param name="securityProvider">Cryptographic provider supplied by the optional OfficeIMO.Security package.</param>
        /// <param name="options">Optional native-tool and validation policy.</param>
        public static WordMacroProjectSignatureValidationResult ValidateMacroProjectSignature(
            string filePath,
            IOfficeSecurityProvider securityProvider,
            WordMacroProjectSignatureValidationOptions? options = null) =>
            WordMacroProjectSignatureService.Validate(filePath, securityProvider, options);

        /// <summary>
        /// Attempts to clear existing VBA signatures, create Microsoft's legacy, agile, and V3 profiles,
        /// verify every profile as it is created, prove VBA-project preservation, and atomically replace the file.
        /// </summary>
        /// <param name="filePath">Path to a saved DOCM or DOTM package.</param>
        /// <param name="securityProvider">Cryptographic provider supplied by the optional OfficeIMO.Security package.</param>
        /// <param name="certificateThumbprint">SHA-1 thumbprint of a certificate with an accessible private key in the configured store.</param>
        /// <param name="options">Optional certificate-store, OfficeSips, timestamp, validation, and resource policy.</param>
        public static WordMacroProjectSigningResult TrySignMacroProject(
            string filePath,
            IOfficeSecurityProvider securityProvider,
            string certificateThumbprint,
            WordMacroProjectSigningOptions? options = null) =>
            WordMacroProjectSignatureService.TrySign(filePath, securityProvider, certificateThumbprint, options);

        /// <summary>
        /// Clears existing VBA signatures, creates and verifies Microsoft's legacy, agile, and V3 profiles,
        /// and atomically commits the signed package or throws with a structured result.
        /// </summary>
        /// <param name="filePath">Path to a saved DOCM or DOTM package.</param>
        /// <param name="securityProvider">Cryptographic provider supplied by the optional OfficeIMO.Security package.</param>
        /// <param name="certificateThumbprint">SHA-1 thumbprint of a certificate with an accessible private key in the configured store.</param>
        /// <param name="options">Optional certificate-store, OfficeSips, timestamp, validation, and resource policy.</param>
        public static WordMacroProjectSigningResult SignMacroProject(
            string filePath,
            IOfficeSecurityProvider securityProvider,
            string certificateThumbprint,
            WordMacroProjectSigningOptions? options = null) {
            WordMacroProjectSigningResult result = TrySignMacroProject(
                filePath,
                securityProvider,
                certificateThumbprint,
                options);
            if (!result.Succeeded) throw new WordMacroProjectSigningException(result);
            return result;
        }
    }
}
