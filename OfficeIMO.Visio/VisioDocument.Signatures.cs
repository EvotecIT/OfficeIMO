using System;
using OfficeIMO.Security;

namespace OfficeIMO.Visio {
    /// <summary>Controls saving when a VSDX package contains digital-signature metadata.</summary>
    public enum VisioSignatureMutationPolicy {
        /// <summary>Block save because rebuilding the package would invalidate its signatures.</summary>
        BlockSave,

        /// <summary>Allow the rebuilt package to omit the invalidated signature carrier.</summary>
        RemoveInvalidatedSignatures
    }

    /// <summary>Describes package-level digital-signature metadata found in a VSDX package.</summary>
    public sealed class VisioSignatureInfo {
        internal VisioSignatureInfo(int originRelationshipCount, int originPartCount, int xmlSignaturePartCount) {
            OriginRelationshipCount = originRelationshipCount;
            OriginPartCount = originPartCount;
            XmlSignaturePartCount = xmlSignaturePartCount;
        }

        /// <summary>Gets the number of package relationships that declare a signature origin.</summary>
        public int OriginRelationshipCount { get; }

        /// <summary>Gets the number of declared signature-origin parts that exist in the package.</summary>
        public int OriginPartCount { get; }

        /// <summary>Gets the number of XML digital-signature parts present in the package.</summary>
        public int XmlSignaturePartCount { get; }

        /// <summary>Gets whether the package contains any signature carrier metadata.</summary>
        public bool HasSignatures => OriginRelationshipCount > 0 || OriginPartCount > 0 || XmlSignaturePartCount > 0;
    }

    /// <summary>Raised when a save would invalidate an existing VSDX package signature.</summary>
    public sealed class VisioSignedDocumentMutationException : InvalidOperationException {
        internal VisioSignedDocumentMutationException(VisioSignatureInfo signatureInfo)
            : base("Saving would rebuild a Visio package that contains digital-signature metadata. "
                + "Set SignatureMutationPolicy to RemoveInvalidatedSignatures to omit that invalid carrier explicitly.") {
            SignatureInfo = signatureInfo;
        }

        /// <summary>Gets the signature metadata that caused the save to be blocked.</summary>
        public VisioSignatureInfo SignatureInfo { get; }
    }

    public partial class VisioDocument {
        private VisioSignatureInfo _loadedSignatureInfo = new(0, 0, 0);

        /// <summary>
        /// Gets or sets the policy applied before a loaded signed VSDX package is rebuilt.
        /// The safe default blocks the save.
        /// </summary>
        public VisioSignatureMutationPolicy SignatureMutationPolicy { get; set; } =
            VisioSignatureMutationPolicy.BlockSave;

        /// <summary>
        /// Inspects the signature carrier discovered when this document was loaded.
        /// This is structural inspection; it does not validate cryptographic trust.
        /// </summary>
        public VisioSignatureInfo InspectSignatures() => _loadedSignatureInfo;

        private static VisioSignatureInfo InspectPackageSignatures(byte[] packageBytes) {
            OfficePackageSignatureInfo info = OfficePackageSignatureService.Inspect(
                packageBytes,
                new OfficePackageSignatureInspectionOptions { VerifyDigests = false });
            return new VisioSignatureInfo(
                info.OriginRelationshipCount,
                info.OriginPartCount,
                info.SignatureParts.Count);
        }

        private void ApplySignatureMutationPolicy() {
            if (!_loadedSignatureInfo.HasSignatures) return;
            if (SignatureMutationPolicy == VisioSignatureMutationPolicy.BlockSave) {
                throw new VisioSignedDocumentMutationException(_loadedSignatureInfo);
            }
        }
    }
}
