namespace OfficeIMO.Word {
    /// <summary>
    /// Signature validation state for one validation dimension.
    /// </summary>
    public enum WordSignatureValidationState {
        /// <summary>No signature metadata exists for this validation dimension.</summary>
        NotPresent,

        /// <summary>The validation dimension has passed the checks OfficeIMO can perform.</summary>
        Passed,

        /// <summary>The validation dimension failed the checks OfficeIMO can perform.</summary>
        Failed,

        /// <summary>The validation dimension was intentionally not checked.</summary>
        NotChecked,

        /// <summary>The validation dimension is not supported by the current OfficeIMO implementation.</summary>
        Unsupported
    }

    /// <summary>
    /// Separates structural package parsing from cryptographic trust claims for Word package signatures.
    /// </summary>
    public sealed class WordSignatureValidationReport {
        internal WordSignatureValidationReport(
            WordSignatureInfo signatureInfo,
            WordSignatureValidationState packageStructureStatus,
            WordSignatureValidationState xmlSignatureStatus,
            WordSignatureValidationState cryptographicStatus,
            WordSignatureValidationState certificateChainStatus,
            WordSignatureValidationState revocationStatus,
            WordSignatureValidationState timestampStatus,
            WordSignatureValidationState signedPartCoverageStatus,
            WordSignatureValidationState signedPartDigestStatus,
            IReadOnlyList<string> findings) {
            SignatureInfo = signatureInfo ?? throw new ArgumentNullException(nameof(signatureInfo));
            PackageStructureStatus = packageStructureStatus;
            XmlSignatureStatus = xmlSignatureStatus;
            CryptographicStatus = cryptographicStatus;
            CertificateChainStatus = certificateChainStatus;
            RevocationStatus = revocationStatus;
            TimestampStatus = timestampStatus;
            SignedPartCoverageStatus = signedPartCoverageStatus;
            SignedPartDigestStatus = signedPartDigestStatus;
            Findings = findings ?? Array.Empty<string>();
        }

        /// <summary>
        /// Gets the signature inspection metadata used to build this validation report.
        /// </summary>
        public WordSignatureInfo SignatureInfo { get; }

        /// <summary>
        /// Gets whether the document has any signature package metadata.
        /// </summary>
        public bool HasSignatures => SignatureInfo.HasSignatures;

        /// <summary>
        /// Gets package-level signature structure status, such as origin part and signature part presence.
        /// </summary>
        public WordSignatureValidationState PackageStructureStatus { get; }

        /// <summary>
        /// Gets XML signature parse status for discovered signature parts.
        /// </summary>
        public WordSignatureValidationState XmlSignatureStatus { get; }

        /// <summary>
        /// Gets cryptographic signature status. OfficeIMO.Word currently does not perform cryptographic validation.
        /// </summary>
        public WordSignatureValidationState CryptographicStatus { get; }

        /// <summary>
        /// Gets certificate-chain trust status. OfficeIMO.Word currently does not perform certificate-chain validation.
        /// </summary>
        public WordSignatureValidationState CertificateChainStatus { get; }

        /// <summary>
        /// Gets revocation status. OfficeIMO.Word currently does not perform revocation checks.
        /// </summary>
        public WordSignatureValidationState RevocationStatus { get; }

        /// <summary>
        /// Gets timestamp status. OfficeIMO.Word reports presence separately from timestamp validation, which is not performed.
        /// </summary>
        public WordSignatureValidationState TimestampStatus { get; }

        /// <summary>
        /// Gets signed package-part reference coverage status. OfficeIMO.Word checks target existence separately from digest verification.
        /// </summary>
        public WordSignatureValidationState SignedPartCoverageStatus { get; }

        /// <summary>
        /// Gets bounded signed package-part digest verification status for simple references without XML DSig transforms.
        /// </summary>
        public WordSignatureValidationState SignedPartDigestStatus { get; }

        /// <summary>
        /// Gets deterministic validation findings and unsupported validation details.
        /// </summary>
        public IReadOnlyList<string> Findings { get; }

        /// <summary>
        /// Gets whether all checks OfficeIMO.Word currently performs passed.
        /// </summary>
        public bool IsStructurallyValid =>
            PackageStructureStatus == WordSignatureValidationState.Passed &&
            XmlSignatureStatus == WordSignatureValidationState.Passed &&
            SignedPartCoverageStatus == WordSignatureValidationState.Passed &&
            SignedPartDigestStatus != WordSignatureValidationState.Failed;

        /// <summary>
        /// Creates a validation report from signature inspection metadata without performing cryptographic trust checks.
        /// </summary>
        public static WordSignatureValidationReport From(WordSignatureInfo signatureInfo) {
            if (signatureInfo == null) throw new ArgumentNullException(nameof(signatureInfo));

            var findings = new List<string>();
            WordSignatureValidationState packageStructureStatus;
            WordSignatureValidationState xmlSignatureStatus;
            WordSignatureValidationState signedPartCoverageStatus;
            WordSignatureValidationState signedPartDigestStatus;

            if (!signatureInfo.HasSignatures) {
                packageStructureStatus = WordSignatureValidationState.NotPresent;
                xmlSignatureStatus = WordSignatureValidationState.NotPresent;
                signedPartCoverageStatus = WordSignatureValidationState.NotPresent;
                signedPartDigestStatus = WordSignatureValidationState.NotPresent;
                findings.Add("No digital-signature package metadata was found.");
            } else if (!signatureInfo.HasDigitalSignatureOriginPart && signatureInfo.HasApplicationSignatureMetadata) {
                packageStructureStatus = WordSignatureValidationState.Unsupported;
                xmlSignatureStatus = WordSignatureValidationState.NotPresent;
                signedPartCoverageStatus = WordSignatureValidationState.NotPresent;
                signedPartDigestStatus = WordSignatureValidationState.NotPresent;
                findings.Add("Extended application properties contain digital-signature metadata, but no digital-signature origin part was found.");
            } else if (signatureInfo.HasDigitalSignatureOriginPart && signatureInfo.SignatureParts.Count == 0) {
                packageStructureStatus = WordSignatureValidationState.Failed;
                xmlSignatureStatus = WordSignatureValidationState.NotPresent;
                signedPartCoverageStatus = WordSignatureValidationState.NotPresent;
                signedPartDigestStatus = WordSignatureValidationState.NotPresent;
                findings.Add("A digital-signature origin part exists, but no XML signature parts were found.");
            } else {
                packageStructureStatus = WordSignatureValidationState.Passed;
                findings.Add("Digital-signature origin and XML signature parts were found.");

                bool hasParseError = signatureInfo.SignatureParts.Any(part => part.HasParseError);
                bool allRequiredXmlMetadataPresent = signatureInfo.SignatureParts.All(part =>
                    !string.IsNullOrWhiteSpace(part.SignatureMethodAlgorithm) &&
                    part.SignedReferences.Count > 0 &&
                    part.SignedReferences.All(reference =>
                        !string.IsNullOrWhiteSpace(reference.DigestMethodAlgorithm) &&
                        reference.HasDigestValue));

                if (hasParseError) {
                    xmlSignatureStatus = WordSignatureValidationState.Failed;
                    findings.Add("At least one XML signature part could not be parsed.");
                } else if (!allRequiredXmlMetadataPresent) {
                    xmlSignatureStatus = WordSignatureValidationState.Unsupported;
                    findings.Add("XML signature parts parsed, but required SignatureMethod, Reference DigestMethod, or Reference DigestValue metadata was missing.");
                } else {
                    xmlSignatureStatus = WordSignatureValidationState.Passed;
                    findings.Add("XML signature parts parsed with signature method, reference digest methods, and digest values.");
                }

                signedPartCoverageStatus = DetermineSignedPartCoverageStatus(signatureInfo, findings);
                signedPartDigestStatus = DetermineSignedPartDigestStatus(signatureInfo, findings);
            }

            if (signatureInfo.HasSignatures) {
                findings.Add("Cryptographic signature validation is not performed by OfficeIMO.Word.");
                findings.Add("Certificate-chain trust validation is not performed by OfficeIMO.Word.");
                findings.Add("Revocation validation is not performed by OfficeIMO.Word.");
                if (HasTimestampMetadata(signatureInfo)) {
                    findings.Add("Signature timestamp metadata was found, but timestamp validation is not performed by OfficeIMO.Word.");
                } else {
                    findings.Add("No signature timestamp metadata was found.");
                }
            }

            foreach (string detail in signatureInfo.UnsupportedDetails) {
                findings.Add(detail);
            }

            return new WordSignatureValidationReport(
                signatureInfo,
                packageStructureStatus,
                xmlSignatureStatus,
                signatureInfo.HasSignatures ? WordSignatureValidationState.NotChecked : WordSignatureValidationState.NotPresent,
                signatureInfo.HasSignatures ? WordSignatureValidationState.NotChecked : WordSignatureValidationState.NotPresent,
                signatureInfo.HasSignatures ? WordSignatureValidationState.NotChecked : WordSignatureValidationState.NotPresent,
                DetermineTimestampStatus(signatureInfo),
                signedPartCoverageStatus,
                signedPartDigestStatus,
                findings.Distinct(StringComparer.OrdinalIgnoreCase).ToArray());
        }

        private static WordSignatureValidationState DetermineTimestampStatus(WordSignatureInfo signatureInfo) {
            if (!signatureInfo.HasSignatures) {
                return WordSignatureValidationState.NotPresent;
            }

            return HasTimestampMetadata(signatureInfo)
                ? WordSignatureValidationState.NotChecked
                : WordSignatureValidationState.NotPresent;
        }

        private static bool HasTimestampMetadata(WordSignatureInfo signatureInfo) {
            return signatureInfo.SignatureParts.Any(part => part.Timestamps.Count > 0);
        }

        private static WordSignatureValidationState DetermineSignedPartCoverageStatus(WordSignatureInfo signatureInfo, List<string> findings) {
            List<WordSignatureReferenceInfo> references = signatureInfo.SignatureParts
                .SelectMany(part => part.SignedReferences)
                .ToList();

            if (references.Count == 0) {
                findings.Add("No XML signature Reference entries were found.");
                return WordSignatureValidationState.Unsupported;
            }

            List<WordSignatureReferenceInfo> packageReferences = references
                .Where(reference => reference.IsPackagePartReference)
                .ToList();

            if (packageReferences.Count == 0) {
                findings.Add("XML signature Reference entries were found, but none point at OPC package parts.");
                return WordSignatureValidationState.Unsupported;
            }

            List<WordSignatureReferenceInfo> missingPackageReferences = packageReferences
                .Where(reference => reference.TargetPartExists == false)
                .ToList();

            if (missingPackageReferences.Count > 0) {
                findings.Add("At least one XML signature Reference points at a missing package part: " +
                             string.Join(", ", missingPackageReferences.Select(reference => reference.TargetPartUri).Where(uri => !string.IsNullOrWhiteSpace(uri))) + ".");
                return WordSignatureValidationState.Failed;
            }

            if (references.Count != packageReferences.Count) {
                findings.Add("Some XML signature Reference entries are not OPC package part references and were left for cryptographic validators.");
            }

            findings.Add("XML signature package-part references resolve to existing package parts.");
            return WordSignatureValidationState.Passed;
        }

        private static WordSignatureValidationState DetermineSignedPartDigestStatus(WordSignatureInfo signatureInfo, List<string> findings) {
            List<WordSignatureReferenceInfo> packageReferences = signatureInfo.SignatureParts
                .SelectMany(part => part.SignedReferences)
                .Where(reference => reference.IsPackagePartReference)
                .ToList();

            if (packageReferences.Count == 0) {
                findings.Add("No OPC package-part references were available for digest verification.");
                return WordSignatureValidationState.NotPresent;
            }

            foreach (string detail in packageReferences
                .Select(reference => reference.DigestVerificationDetail)
                .Where(detail => !string.IsNullOrWhiteSpace(detail))
                .Select(detail => detail!)) {
                findings.Add(detail);
            }

            if (packageReferences.Any(reference => reference.DigestVerificationStatus == WordSignatureValidationState.Failed)) {
                findings.Add("At least one signed package-part digest did not match the package content.");
                return WordSignatureValidationState.Failed;
            }

            if (packageReferences.Any(reference => reference.DigestVerificationStatus == WordSignatureValidationState.Unsupported)) {
                findings.Add("At least one signed package-part digest was not checked because the reference requires unsupported transform-aware validation.");
                return WordSignatureValidationState.Unsupported;
            }

            if (packageReferences.Any(reference => reference.DigestVerificationStatus == WordSignatureValidationState.NotChecked)) {
                findings.Add("At least one signed package-part digest was not checked because digest metadata or package content was unavailable.");
                return WordSignatureValidationState.Unsupported;
            }

            findings.Add("Signed package-part digests match for simple transform-free references.");
            return WordSignatureValidationState.Passed;
        }
    }
}
