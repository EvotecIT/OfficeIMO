using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Shared;

namespace OfficeIMO.Word {
    /// <summary>
    /// Describes digital-signature package metadata found in a Word document.
    /// </summary>
    public sealed class WordSignatureInfo {
        internal WordSignatureInfo(
            bool hasDigitalSignatureOriginPart,
            string? originPartUri,
            string? originRelationshipId,
            bool hasApplicationSignatureMetadata,
            IReadOnlyList<WordSignaturePartInfo> signatureParts,
            IReadOnlyList<string> unsupportedDetails,
            IReadOnlyList<string> details) {
            HasDigitalSignatureOriginPart = hasDigitalSignatureOriginPart;
            OriginPartUri = originPartUri;
            OriginRelationshipId = originRelationshipId;
            HasApplicationSignatureMetadata = hasApplicationSignatureMetadata;
            SignatureParts = signatureParts;
            UnsupportedDetails = unsupportedDetails;
            Details = details;
        }

        /// <summary>
        /// Gets whether any package signature metadata was found.
        /// </summary>
        public bool HasSignatures => HasDigitalSignatureOriginPart || HasApplicationSignatureMetadata || SignatureParts.Count > 0;

        /// <summary>
        /// Gets whether the package contains a digital-signature origin part.
        /// </summary>
        public bool HasDigitalSignatureOriginPart { get; }

        /// <summary>
        /// Gets the signature origin part URI when present.
        /// </summary>
        public string? OriginPartUri { get; }

        /// <summary>
        /// Gets the package relationship id for the signature origin part when available.
        /// </summary>
        public string? OriginRelationshipId { get; }

        /// <summary>
        /// Gets whether extended application properties contain digital-signature metadata.
        /// </summary>
        public bool HasApplicationSignatureMetadata { get; }

        /// <summary>
        /// Gets signature XML parts discovered under the signature origin.
        /// </summary>
        public IReadOnlyList<WordSignaturePartInfo> SignatureParts { get; }

        /// <summary>
        /// Gets unsupported or unknown details callers should not treat as validation proof.
        /// </summary>
        public IReadOnlyList<string> UnsupportedDetails { get; }

        /// <summary>
        /// Gets human-readable package details suitable for feature reports.
        /// </summary>
        public IReadOnlyList<string> Details { get; }

        /// <summary>
        /// Gets a count suitable for feature-report findings.
        /// </summary>
        public int FindingCount =>
            (HasDigitalSignatureOriginPart ? 1 : 0) +
            SignatureParts.Count +
            (HasApplicationSignatureMetadata ? 1 : 0);

        internal static WordSignatureInfo FromPackageInfo(OfficePackageSignatureInfo packageInfo) {
            return new WordSignatureInfo(
                packageInfo.HasDigitalSignatureOriginPart,
                packageInfo.OriginPartUri,
                packageInfo.OriginRelationshipId,
                packageInfo.HasApplicationSignatureMetadata,
                packageInfo.SignatureParts.Select(WordSignaturePartInfo.FromPackagePart).ToArray(),
                packageInfo.UnsupportedDetails,
                packageInfo.Details);
        }
    }

    /// <summary>
    /// Describes one XML signature part in a Word package.
    /// </summary>
    public sealed class WordSignaturePartInfo {
        internal WordSignaturePartInfo(
            string uri,
            string contentType,
            string? relationshipId,
            long? length,
            string? signatureMethodAlgorithm,
            IReadOnlyList<string> digestMethodAlgorithms,
            IReadOnlyList<WordSignatureReferenceInfo> signedReferences,
            IReadOnlyList<WordSignatureTimestampInfo> timestamps,
            IReadOnlyList<string> x509SubjectNames,
            string? parseError,
            IReadOnlyList<string> unsupportedDetails) {
            Uri = uri;
            ContentType = contentType;
            RelationshipId = relationshipId;
            Length = length;
            SignatureMethodAlgorithm = signatureMethodAlgorithm;
            DigestMethodAlgorithms = digestMethodAlgorithms;
            SignedReferences = signedReferences;
            Timestamps = timestamps;
            X509SubjectNames = x509SubjectNames;
            ParseError = parseError;
            UnsupportedDetails = unsupportedDetails;
        }

        /// <summary>Gets the signature part URI.</summary>
        public string Uri { get; }

        /// <summary>Gets the signature part content type.</summary>
        public string ContentType { get; }

        /// <summary>Gets the relationship id from the signature origin part when available.</summary>
        public string? RelationshipId { get; }

        /// <summary>Gets the signature part byte length when the stream supports length.</summary>
        public long? Length { get; }

        /// <summary>Gets the XML DSig signature method algorithm when parseable.</summary>
        public string? SignatureMethodAlgorithm { get; }

        /// <summary>Gets XML DSig digest method algorithms when parseable.</summary>
        public IReadOnlyList<string> DigestMethodAlgorithms { get; }

        /// <summary>Gets XML DSig signed references discovered in the signature part.</summary>
        public IReadOnlyList<WordSignatureReferenceInfo> SignedReferences { get; }

        /// <summary>Gets timestamp declarations discovered in the signature XML.</summary>
        public IReadOnlyList<WordSignatureTimestampInfo> Timestamps { get; }

        /// <summary>Gets XML DSig X509 subject names when parseable.</summary>
        public IReadOnlyList<string> X509SubjectNames { get; }

        /// <summary>Gets the XML parse error, if the signature part could not be parsed.</summary>
        public string? ParseError { get; }

        /// <summary>Gets whether the XML signature part could not be parsed.</summary>
        public bool HasParseError => !string.IsNullOrWhiteSpace(ParseError);

        /// <summary>Gets unsupported or parse details for this signature part.</summary>
        public IReadOnlyList<string> UnsupportedDetails { get; }

        internal static WordSignaturePartInfo FromPackagePart(OfficePackageSignaturePartInfo packagePart) {
            return new WordSignaturePartInfo(
                packagePart.Uri,
                packagePart.ContentType,
                packagePart.RelationshipId,
                packagePart.Length,
                packagePart.SignatureMethodAlgorithm,
                packagePart.DigestMethodAlgorithms,
                packagePart.SignedReferences.Select(WordSignatureReferenceInfo.FromPackageReference).ToArray(),
                packagePart.Timestamps.Select(WordSignatureTimestampInfo.FromPackageTimestamp).ToArray(),
                packagePart.X509SubjectNames,
                packagePart.ParseError,
                packagePart.UnsupportedDetails);
        }
    }

    /// <summary>
    /// Describes one XML DSig reference entry in a signature part.
    /// </summary>
    public sealed class WordSignatureReferenceInfo {
        internal WordSignatureReferenceInfo(
            string? uri,
            string? digestMethodAlgorithm,
            string? digestValue,
            bool isPackagePartReference,
            string? targetPartUri,
            bool? targetPartExists,
            IReadOnlyList<string> transformAlgorithms,
            WordSignatureValidationState digestVerificationStatus,
            string? digestVerificationDetail) {
            Uri = uri;
            DigestMethodAlgorithm = digestMethodAlgorithm;
            DigestValue = digestValue;
            IsPackagePartReference = isPackagePartReference;
            TargetPartUri = targetPartUri;
            TargetPartExists = targetPartExists;
            TransformAlgorithms = transformAlgorithms;
            DigestVerificationStatus = digestVerificationStatus;
            DigestVerificationDetail = digestVerificationDetail;
        }

        /// <summary>Gets the XML DSig Reference URI value.</summary>
        public string? Uri { get; }

        /// <summary>Gets the reference digest method algorithm when parseable.</summary>
        public string? DigestMethodAlgorithm { get; }

        /// <summary>Gets the reference digest value when parseable.</summary>
        public string? DigestValue { get; }

        /// <summary>Gets whether the reference includes a digest value.</summary>
        public bool HasDigestValue => !string.IsNullOrWhiteSpace(DigestValue);

        /// <summary>Gets whether the reference points at an OPC package part URI.</summary>
        public bool IsPackagePartReference { get; }

        /// <summary>Gets the normalized target package part URI when the reference points at a package part.</summary>
        public string? TargetPartUri { get; }

        /// <summary>Gets whether the target package part exists, or null when the reference is not a package part reference.</summary>
        public bool? TargetPartExists { get; }

        /// <summary>Gets XML DSig transform algorithms declared on the reference.</summary>
        public IReadOnlyList<string> TransformAlgorithms { get; }

        /// <summary>Gets bounded digest-verification status for simple package-part references.</summary>
        public WordSignatureValidationState DigestVerificationStatus { get; }

        /// <summary>Gets a deterministic digest-verification detail or unsupported reason.</summary>
        public string? DigestVerificationDetail { get; }

        internal static WordSignatureReferenceInfo FromPackageReference(OfficePackageSignatureReferenceInfo packageReference) {
            return new WordSignatureReferenceInfo(
                packageReference.Uri,
                packageReference.DigestMethodAlgorithm,
                packageReference.DigestValue,
                packageReference.IsPackagePartReference,
                packageReference.TargetPartUri,
                packageReference.TargetPartExists,
                packageReference.TransformAlgorithms,
                MapDigestStatus(packageReference.DigestVerificationStatus),
                packageReference.DigestVerificationDetail);
        }

        private static WordSignatureValidationState MapDigestStatus(OfficePackageSignatureDigestVerificationStatus status) {
            switch (status) {
                case OfficePackageSignatureDigestVerificationStatus.Passed:
                    return WordSignatureValidationState.Passed;
                case OfficePackageSignatureDigestVerificationStatus.Failed:
                    return WordSignatureValidationState.Failed;
                case OfficePackageSignatureDigestVerificationStatus.Unsupported:
                    return WordSignatureValidationState.Unsupported;
                default:
                    return WordSignatureValidationState.NotChecked;
            }
        }
    }

    /// <summary>
    /// Describes timestamp metadata declared inside one XML signature part.
    /// </summary>
    public sealed class WordSignatureTimestampInfo {
        internal WordSignatureTimestampInfo(string kind, string? value, string? format) {
            Kind = kind;
            Value = value;
            Format = format;
        }

        /// <summary>Gets the recognized timestamp declaration kind.</summary>
        public string Kind { get; }

        /// <summary>Gets the timestamp value when the declaration exposes one as text.</summary>
        public string? Value { get; }

        /// <summary>Gets the timestamp format when the declaration exposes one.</summary>
        public string? Format { get; }

        internal static WordSignatureTimestampInfo FromPackageTimestamp(OfficePackageSignatureTimestampInfo packageTimestamp) {
            return new WordSignatureTimestampInfo(
                packageTimestamp.Kind,
                packageTimestamp.Value,
                packageTimestamp.Format);
        }
    }

    internal static class WordSignatureInspector {
        internal static WordSignatureInfo Inspect(
            WordprocessingDocument package,
            DigitalSignatureOriginPart? originPart,
            bool hasApplicationSignatureMetadata) {
            if (package == null) throw new ArgumentNullException(nameof(package));

            return WordSignatureInfo.FromPackageInfo(
                OfficePackageSignatureInspector.Inspect(package, originPart, hasApplicationSignatureMetadata));
        }
    }
}
