using OfficeIMO.Security;

namespace OfficeIMO.Word {
    /// <summary>Stable diagnostic for one OPC XML-signature validation check.</summary>
    public sealed class WordSignatureValidationFinding {
        internal WordSignatureValidationFinding(
            string code,
            WordSignatureValidationState state,
            string message,
            string? signaturePartUri = null,
            string? referenceUri = null) {
            Code = code;
            State = state;
            Message = message;
            SignaturePartUri = signaturePartUri;
            ReferenceUri = referenceUri;
        }

        /// <summary>Gets the stable machine-readable finding code.</summary>
        public string Code { get; }

        /// <summary>Gets the validation state represented by the finding.</summary>
        public WordSignatureValidationState State { get; }

        /// <summary>Gets the human-readable finding message.</summary>
        public string Message { get; }

        /// <summary>Gets the owning signature part URI when applicable.</summary>
        public string? SignaturePartUri { get; }

        /// <summary>Gets the signed reference URI when applicable.</summary>
        public string? ReferenceUri { get; }
    }

    /// <summary>Cryptographic, trust, revocation, and timestamp validation for one XML signature part.</summary>
    public sealed class WordSignaturePartValidationResult {
        internal WordSignaturePartValidationResult(
            WordSignaturePartInfo signaturePart,
            WordSignatureValidationState cryptographicStatus,
            WordSignatureValidationState certificateChainStatus,
            WordSignatureValidationState revocationStatus,
            bool revocationCheckRequired,
            WordSignatureValidationState timestampStatus,
            CertificateValidationResult? certificateValidation,
            IReadOnlyList<Rfc3161TimestampVerificationResult> timestampTokens,
            IReadOnlyList<WordSignatureValidationFinding> findings) {
            SignaturePart = signaturePart;
            CryptographicStatus = cryptographicStatus;
            CertificateChainStatus = certificateChainStatus;
            RevocationStatus = revocationStatus;
            RevocationCheckRequired = revocationCheckRequired;
            TimestampStatus = timestampStatus;
            CertificateValidation = certificateValidation;
            TimestampTokens = timestampTokens;
            Findings = findings;
        }

        /// <summary>Gets parsed metadata for the validated signature part.</summary>
        public WordSignaturePartInfo SignaturePart { get; }

        /// <summary>Gets XML DSig signature-value and signed-object validation status.</summary>
        public WordSignatureValidationState CryptographicStatus { get; }

        /// <summary>Gets signer certificate-chain trust status.</summary>
        public WordSignatureValidationState CertificateChainStatus { get; }

        /// <summary>Gets signer certificate revocation status under caller policy.</summary>
        public WordSignatureValidationState RevocationStatus { get; }

        /// <summary>Gets whether caller policy required a conclusive signer revocation check.</summary>
        public bool RevocationCheckRequired { get; }

        /// <summary>Gets combined RFC 3161 timestamp-token status.</summary>
        public WordSignatureValidationState TimestampStatus { get; }

        /// <summary>Gets the shared certificate validation result when a signer certificate was found.</summary>
        public CertificateValidationResult? CertificateValidation { get; }

        /// <summary>Gets every RFC 3161 timestamp token result.</summary>
        public IReadOnlyList<Rfc3161TimestampVerificationResult> TimestampTokens { get; }

        /// <summary>Gets stable validation findings for this signature part.</summary>
        public IReadOnlyList<WordSignatureValidationFinding> Findings { get; }

        /// <summary>Gets whether signature math, package-reference digests, and signer trust passed without a failed revocation or timestamp check.</summary>
        public bool IsValidUnderPolicy =>
            CryptographicStatus == WordSignatureValidationState.Passed &&
            SignaturePart.SignedReferences.Count > 0 &&
            SignaturePart.SignedReferences.All(reference =>
                reference.IsPackagePartReference &&
                reference.TargetPartExists == true &&
                reference.DigestVerificationStatus == WordSignatureValidationState.Passed) &&
            CertificateChainStatus == WordSignatureValidationState.Passed &&
            RevocationStatus != WordSignatureValidationState.Failed &&
            (!RevocationCheckRequired || RevocationStatus == WordSignatureValidationState.Passed) &&
            TimestampStatus != WordSignatureValidationState.Failed &&
            TimestampStatus != WordSignatureValidationState.Unsupported;
    }
}
