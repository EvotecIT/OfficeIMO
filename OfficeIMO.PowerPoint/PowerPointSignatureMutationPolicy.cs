using System;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Security;

namespace OfficeIMO.PowerPoint {
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
            bool hasLegacyXmlSignatureStorage, OfficeSignatureMutationPolicy policy,
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
        public OfficeSignatureMutationPolicy Policy { get; }
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
        public OfficeSignatureMutationPolicy SignatureMutationPolicy { get; set; } =
            OfficeSignatureMutationPolicy.BlockSave;

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

            if (SignatureMutationPolicy == OfficeSignatureMutationPolicy.BlockSave) {
                LastSignatureReport = CreateSignatureReport(PowerPointSignatureMutationAction.Blocked);
                throw new PowerPointSignedPresentationMutationException(LastSignatureReport);
            }
            if (SignatureMutationPolicy == OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures) {
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
            OfficePackageSignatureInfo? shared = null;
            // A binary presentation is projected into a temporary Open XML package. Cloning that
            // projection can materialize deferred DOM state and invalidate the exact-no-op
            // preservation fingerprint, so its original binary signature carriers are inspected
            // directly below without serializing the projection.
            if (_document != null && _legacyPptPackage == null) {
                bool hasLiveSignatureCarrier =
                    _document.DigitalSignatureOriginPart != null
                    || _document.ExtendedFilePropertiesPart?.Properties?
                        .DigitalSignature != null;

                // Ordinary unsigned saves should not need to clone and serialize
                // the complete package merely to prove the absence of signature
                // metadata. Inspect the current package bytes first so malformed
                // or relationship-only carriers remain fail-closed. A live carrier
                // still takes the full snapshot path below because unsaved Open XML
                // mutations may not yet be reflected in the package stream.
                if (!hasLiveSignatureCarrier
                    && _packageStream is MemoryStream currentPackage) {
                    try {
                        shared = OfficePackageSignatureService.Inspect(
                            currentPackage.ToArray(),
                            new OfficePackageSignatureInspectionOptions {
                                VerifyDigests = false
                            });
                        if (!shared.HasSignatures) {
                            return CreateSignatureReport(shared, action);
                        }
                    } catch (InvalidDataException) {
                        // A package that cannot be inspected from its current
                        // stream state is handled by the existing clone path.
                    } catch (IOException) {
                        // Preserve the fail-closed snapshot fallback.
                    } catch (System.Xml.XmlException) {
                        // Some package implementations expose a structurally
                        // open stream before all XML parts are finalized.
                    }
                }

                using var snapshot = new MemoryStream();
                using (_document.Clone(snapshot)) { }
                shared = OfficePackageSignatureService.Inspect(
                    snapshot.ToArray(),
                    new OfficePackageSignatureInspectionOptions { VerifyDigests = false });
            }
            bool legacyBinarySignature = _legacyPptPackage?.HasBinarySignatureStream == true;
            bool legacyXmlSignature = _legacyPptPackage?.HasXmlSignatureStorage == true;
            return CreateSignatureReport(shared, action,
                legacyBinarySignature, legacyXmlSignature);
        }

        private PowerPointSignatureReport CreateSignatureReport(
            OfficePackageSignatureInfo? shared,
            PowerPointSignatureMutationAction action,
            bool? hasLegacyBinarySignature = null,
            bool? hasLegacyXmlSignature = null) =>
            new PowerPointSignatureReport(shared?.HasDigitalSignatureOriginPart == true,
                shared?.SignatureParts.Count ?? 0,
                shared?.HasApplicationSignatureMetadata == true,
                hasLegacyBinarySignature
                    ?? _legacyPptPackage?.HasBinarySignatureStream == true,
                hasLegacyXmlSignature
                    ?? _legacyPptPackage?.HasXmlSignatureStorage == true,
                SignatureMutationPolicy, action);
    }
}
