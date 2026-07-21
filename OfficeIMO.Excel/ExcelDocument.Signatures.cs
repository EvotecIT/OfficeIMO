using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Excel {
    /// <summary>Describes digital-signature package metadata found in an Excel workbook.</summary>
    public sealed class ExcelSignatureInfo {
        internal ExcelSignatureInfo(
            bool hasDigitalSignatureOriginPart,
            int xmlSignaturePartCount,
            bool hasApplicationSignatureMetadata) {
            HasDigitalSignatureOriginPart = hasDigitalSignatureOriginPart;
            XmlSignaturePartCount = xmlSignaturePartCount;
            HasApplicationSignatureMetadata = hasApplicationSignatureMetadata;
        }

        /// <summary>Gets whether any package signature metadata was found.</summary>
        public bool HasSignatures => HasDigitalSignatureOriginPart
            || XmlSignaturePartCount > 0
            || HasApplicationSignatureMetadata;

        /// <summary>Gets whether the package contains a digital-signature origin part.</summary>
        public bool HasDigitalSignatureOriginPart { get; }

        /// <summary>Gets the number of XML signature parts under the signature origin.</summary>
        public int XmlSignaturePartCount { get; }

        /// <summary>Gets whether extended application properties contain digital-signature metadata.</summary>
        public bool HasApplicationSignatureMetadata { get; }
    }

    /// <summary>Raised when save policy blocks rewriting a signed Excel workbook.</summary>
    public sealed class ExcelSignedWorkbookMutationException : InvalidOperationException {
        internal ExcelSignedWorkbookMutationException(ExcelSignatureInfo signatureInfo)
            : base("Saving would rewrite an Excel workbook that contains digital-signature metadata. "
                + "Choose RemoveInvalidatedSignatures or PreserveSignatureMarkup explicitly to continue.") {
            SignatureInfo = signatureInfo;
        }

        /// <summary>Gets the signature metadata that caused the save to be blocked.</summary>
        public ExcelSignatureInfo SignatureInfo { get; }
    }

    public partial class ExcelDocument {
        /// <summary>Inspects package-level digital-signature metadata without validating cryptographic trust.</summary>
        public ExcelSignatureInfo InspectSignatures() {
            DigitalSignatureOriginPart? originPart = _spreadSheetDocument.DigitalSignatureOriginPart;
            return new ExcelSignatureInfo(
                originPart != null,
                originPart?.XmlSignatureParts.Count() ?? 0,
                _spreadSheetDocument.ExtendedFilePropertiesPart?.Properties?.DigitalSignature != null);
        }

        private void ApplySignatureMutationPolicy(ExcelSaveOptions? options) {
            ExcelSignatureInfo signatureInfo = InspectSignatures();
            if (!signatureInfo.HasSignatures) {
                return;
            }

            ExcelSignatureMutationPolicy policy = options?.SignatureMutationPolicy
                ?? ExcelSignatureMutationPolicy.BlockSave;
            if (policy == ExcelSignatureMutationPolicy.BlockSave) {
                throw new ExcelSignedWorkbookMutationException(signatureInfo);
            }

            if (policy != ExcelSignatureMutationPolicy.RemoveInvalidatedSignatures) {
                return;
            }

            DigitalSignatureOriginPart? originPart = _spreadSheetDocument.DigitalSignatureOriginPart;
            if (originPart != null) {
                _spreadSheetDocument.DeletePart(originPart);
            }

            if (_spreadSheetDocument.ExtendedFilePropertiesPart?.Properties?.DigitalSignature != null) {
                _spreadSheetDocument.ExtendedFilePropertiesPart.Properties.DigitalSignature = null;
                _spreadSheetDocument.ExtendedFilePropertiesPart.Properties.Save();
            }
        }
    }
}
