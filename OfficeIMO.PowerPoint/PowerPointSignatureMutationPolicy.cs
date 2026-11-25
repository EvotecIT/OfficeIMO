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
        /// <summary>Save policy the report was produced under. Inspection outside a save reports the default <see cref="OfficeSignatureMutationPolicy.BlockSave"/>.</summary>
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
    public sealed class PowerPointSignedDocumentMutationException : InvalidOperationException {
        internal PowerPointSignedDocumentMutationException(PowerPointSignatureReport report)
            : base("Saving would mutate a presentation that contains digital-signature metadata. " +
                   "Choose RemoveInvalidatedSignatures or PreserveSignatureMarkup explicitly to continue.") {
            Report = report;
        }

        /// <summary>Signature evidence that caused the block.</summary>
        public PowerPointSignatureReport Report { get; }
    }

    public sealed partial class PowerPointDocument {
        internal PowerPointSignatureReport? LastSignatureReport { get; private set; }

        /// <summary>Inspects package signature metadata without mutating it.</summary>
        public PowerPointSignatureReport InspectSignatures() {
            ThrowIfDisposed();
            LastSignatureReport = CreateSignatureReport(OfficeSignatureMutationPolicy.BlockSave,
                PowerPointSignatureMutationAction.None);
            return LastSignatureReport;
        }

        private void ApplySignatureMutationPolicy(PowerPointSaveOptions? options) {
            OfficeSignatureMutationPolicy policy = options?.SignatureMutationPolicy
                ?? OfficeSignatureMutationPolicy.BlockSave;
            PowerPointSignatureReport inspection = CreateSignatureReport(policy, PowerPointSignatureMutationAction.None);
            if (!inspection.HasSignatureMetadata) {
                LastSignatureReport = inspection;
                return;
            }

            if (policy == OfficeSignatureMutationPolicy.BlockSave) {
                LastSignatureReport = CreateSignatureReport(policy, PowerPointSignatureMutationAction.Blocked);
                throw new PowerPointSignedDocumentMutationException(LastSignatureReport);
            }
            if (policy == OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures) {
                DigitalSignatureOriginPart? origin = SdkDocument!.DigitalSignatureOriginPart;
                if (origin != null) SdkDocument.DeletePart(origin);
                if (SdkDocument.ExtendedFilePropertiesPart?.Properties?.DigitalSignature != null) {
                    SdkDocument.ExtendedFilePropertiesPart.Properties.DigitalSignature = null;
                    SdkDocument.ExtendedFilePropertiesPart.Properties.Save();
                }
                LastSignatureReport = new PowerPointSignatureReport(inspection.HasOriginPart,
                    inspection.XmlSignaturePartCount, inspection.HasApplicationSignatureFlag,
                    inspection.HasLegacyBinarySignatureStream, inspection.HasLegacyXmlSignatureStorage,
                    policy, PowerPointSignatureMutationAction.Removed);
                return;
            }

            LastSignatureReport = CreateSignatureReport(policy, PowerPointSignatureMutationAction.Preserved);
        }

        private PowerPointSignatureReport CreateSignatureReport(OfficeSignatureMutationPolicy policy,
            PowerPointSignatureMutationAction action) {
            OfficePackageSignatureInfo? shared = null;
            // A binary presentation is projected into a temporary Open XML package. Cloning that
            // projection can materialize deferred DOM state and invalidate the exact-no-op
            // preservation fingerprint, so its original binary signature carriers are inspected
            // directly below without serializing the projection.
            if (_legacyPptPackage == null) {
                // Before the SDK opens, the package is as it was loaded, and a signed one opens the SDK at load.
                bool hasLiveSignatureCarrier = _sdkDocument != null
                    && (_sdkDocument.DigitalSignatureOriginPart != null
                        || _sdkDocument.ExtendedFilePropertiesPart?.Properties?.DigitalSignature != null);

                // Inspect the encoded package directly for read-only documents, since
                // a clone can omit unreferenced signature parts. For editable packages,
                // an unsigned stream can skip cloning. A live carrier still needs a
                // snapshot because the stream may not yet reflect unsaved edits.
                bool readOnly = AccessMode == DocumentAccessMode.ReadOnly;
                if ((readOnly || !hasLiveSignatureCarrier)
                    && _packageStream is { CanRead: true, CanSeek: true } currentPackage) {
                    try {
                        shared = OfficePackageSignatureService.Inspect(
                            currentPackage,
                            new OfficePackageSignatureInspectionOptions {
                                VerifyDigests = false
                            });
                        if (readOnly || !shared.HasSignatures) {
                            return CreateSignatureReport(shared, policy, action);
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
                using (SdkDocument.Clone(snapshot)) { }
                shared = OfficePackageSignatureService.Inspect(
                    snapshot.ToArray(),
                    new OfficePackageSignatureInspectionOptions { VerifyDigests = false });
            }
            bool legacyBinarySignature = _legacyPptPackage?.HasBinarySignatureStream == true;
            bool legacyXmlSignature = _legacyPptPackage?.HasXmlSignatureStorage == true;
            return CreateSignatureReport(shared, policy, action,
                legacyBinarySignature, legacyXmlSignature);
        }

        private PowerPointSignatureReport CreateSignatureReport(
            OfficePackageSignatureInfo? shared,
            OfficeSignatureMutationPolicy policy,
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
                policy, action);
    }
}
