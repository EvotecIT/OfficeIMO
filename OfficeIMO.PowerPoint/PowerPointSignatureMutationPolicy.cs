using System;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.PowerPoint {
    /// <summary>Policy applied before saving a presentation that carries digital-signature metadata.</summary>
    public enum PowerPointSignatureMutationPolicy {
        /// <summary>Block save to prevent silently invalidating an existing signature.</summary>
        BlockSave,
        /// <summary>Remove signature parts and application metadata before saving the mutated package.</summary>
        RemoveInvalidatedSignatures,
        /// <summary>Preserve signature markup even though package mutation can invalidate it.</summary>
        PreserveSignatureMarkup
    }

    /// <summary>Action taken by the latest signature mutation check.</summary>
    public enum PowerPointSignatureMutationAction {
        /// <summary>No signature metadata was present.</summary>
        None,
        /// <summary>Save was blocked.</summary>
        Blocked,
        /// <summary>Signature metadata was removed.</summary>
        Removed,
        /// <summary>Signature metadata was preserved by explicit policy.</summary>
        Preserved
    }

    /// <summary>Structured signature inspection and mutation-policy evidence.</summary>
    public sealed class PowerPointSignatureReport {
        internal PowerPointSignatureReport(bool hasOriginPart, int xmlSignaturePartCount,
            bool hasApplicationSignatureFlag, bool hasLegacyBinarySignatureStream,
            bool hasLegacyXmlSignatureStorage, PowerPointSignatureMutationPolicy policy,
            PowerPointSignatureMutationAction action) {
            HasOriginPart = hasOriginPart;
            XmlSignaturePartCount = xmlSignaturePartCount;
            HasApplicationSignatureFlag = hasApplicationSignatureFlag;
            HasLegacyBinarySignatureStream = hasLegacyBinarySignatureStream;
            HasLegacyXmlSignatureStorage = hasLegacyXmlSignatureStorage;
            Policy = policy;
            Action = action;
        }

        /// <summary>Whether the package contains a digital-signature origin part.</summary>
        public bool HasOriginPart { get; }
        /// <summary>Number of XML signature parts.</summary>
        public int XmlSignaturePartCount { get; }
        /// <summary>Whether extended application properties advertise a digital signature.</summary>
        public bool HasApplicationSignatureFlag { get; }
        /// <summary>Whether a binary PowerPoint package contains the legacy <c>_signatures</c> stream.</summary>
        public bool HasLegacyBinarySignatureStream { get; }
        /// <summary>Whether a binary PowerPoint package contains the legacy <c>_xmlsignatures</c> storage.</summary>
        public bool HasLegacyXmlSignatureStorage { get; }
        /// <summary>Configured save policy.</summary>
        public PowerPointSignatureMutationPolicy Policy { get; }
        /// <summary>Policy action taken.</summary>
        public PowerPointSignatureMutationAction Action { get; }
        /// <summary>Whether any signature metadata was detected.</summary>
        public bool HasSignatureMetadata => HasOriginPart || XmlSignaturePartCount > 0
            || HasApplicationSignatureFlag || HasLegacyBinarySignatureStream
            || HasLegacyXmlSignatureStorage;

        /// <summary>Serializes the report as deterministic JSON.</summary>
        public string ToJson() => new StringBuilder()
            .Append("{\"hasOriginPart\":").Append(HasOriginPart ? "true" : "false")
            .Append(",\"xmlSignaturePartCount\":").Append(XmlSignaturePartCount)
            .Append(",\"hasApplicationSignatureFlag\":").Append(HasApplicationSignatureFlag ? "true" : "false")
            .Append(",\"hasLegacyBinarySignatureStream\":").Append(HasLegacyBinarySignatureStream ? "true" : "false")
            .Append(",\"hasLegacyXmlSignatureStorage\":").Append(HasLegacyXmlSignatureStorage ? "true" : "false")
            .Append(",\"policy\":\"").Append(Policy)
            .Append("\",\"action\":\"").Append(Action).Append("\"}").ToString();
    }

    /// <summary>Raised when the signature mutation policy blocks a save.</summary>
    public sealed class PowerPointSignedPresentationMutationException : InvalidOperationException {
        internal PowerPointSignedPresentationMutationException(PowerPointSignatureReport report)
            : base("Saving would mutate a presentation that contains digital-signature metadata. " +
                   "Choose RemoveInvalidatedSignatures or PreserveSignatureMarkup explicitly to continue.") {
            Report = report;
        }

        /// <summary>Signature evidence that caused the block.</summary>
        public PowerPointSignatureReport Report { get; }
    }

    public sealed partial class PowerPointPresentation {
        /// <summary>
        /// Signature policy applied before save. The safe default blocks mutation of signed packages.
        /// </summary>
        public PowerPointSignatureMutationPolicy SignatureMutationPolicy { get; set; } =
            PowerPointSignatureMutationPolicy.BlockSave;

        internal PowerPointSignatureReport? LastSignatureReport { get; private set; }

        /// <summary>Inspects package signature metadata without mutating it.</summary>
        public PowerPointSignatureReport InspectSignatures() {
            ThrowIfDisposed();
            LastSignatureReport = CreateSignatureReport(PowerPointSignatureMutationAction.None);
            return LastSignatureReport;
        }

        private void ApplySignatureMutationPolicy() {
            PowerPointSignatureReport inspection = CreateSignatureReport(PowerPointSignatureMutationAction.None);
            if (!inspection.HasSignatureMetadata) {
                LastSignatureReport = inspection;
                return;
            }

            if (SignatureMutationPolicy == PowerPointSignatureMutationPolicy.BlockSave) {
                LastSignatureReport = CreateSignatureReport(PowerPointSignatureMutationAction.Blocked);
                throw new PowerPointSignedPresentationMutationException(LastSignatureReport);
            }
            if (SignatureMutationPolicy == PowerPointSignatureMutationPolicy.RemoveInvalidatedSignatures) {
                DigitalSignatureOriginPart? origin = _document!.DigitalSignatureOriginPart;
                if (origin != null) _document.DeletePart(origin);
                if (_document.ExtendedFilePropertiesPart?.Properties?.DigitalSignature != null) {
                    _document.ExtendedFilePropertiesPart.Properties.DigitalSignature = null;
                    _document.ExtendedFilePropertiesPart.Properties.Save();
                }
                LastSignatureReport = new PowerPointSignatureReport(inspection.HasOriginPart,
                    inspection.XmlSignaturePartCount, inspection.HasApplicationSignatureFlag,
                    inspection.HasLegacyBinarySignatureStream, inspection.HasLegacyXmlSignatureStorage,
                    SignatureMutationPolicy, PowerPointSignatureMutationAction.Removed);
                return;
            }

            LastSignatureReport = CreateSignatureReport(PowerPointSignatureMutationAction.Preserved);
        }

        private PowerPointSignatureReport CreateSignatureReport(PowerPointSignatureMutationAction action) {
            DigitalSignatureOriginPart? origin = _document?.DigitalSignatureOriginPart;
            int count = origin?.XmlSignatureParts.Count() ?? 0;
            bool applicationFlag = _document?.ExtendedFilePropertiesPart?.Properties?.DigitalSignature != null;
            bool legacyBinarySignature = _legacyPptPackage?.HasBinarySignatureStream == true;
            bool legacyXmlSignature = _legacyPptPackage?.HasXmlSignatureStorage == true;
            return new PowerPointSignatureReport(origin != null, count, applicationFlag,
                legacyBinarySignature, legacyXmlSignature,
                SignatureMutationPolicy, action);
        }
    }
}
