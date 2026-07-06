using OfficeIMO.Pdf;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Reader.Pdf;

public static partial class DocumentReaderPdfExtensions {
    private static void AddSecurityMetadata(List<OfficeDocumentMetadataEntry> entries, PdfDocumentSecurityInfo security) {
        bool hasSecurityState =
            security.HasEncryption ||
            security.HasSignatures ||
            security.HasReadableEncryptionSettings ||
            security.HasIncrementalUpdates ||
            security.HasDocMDPPermissions ||
            security.HasUsageRights ||
            security.HasDocumentSecurityStore;
        if (!hasSecurityState) {
            return;
        }

        AddCountMetadata(entries, "pdf-security-state-count", "pdf.security", "SecurityStateCount", 1);
        AddCountMetadata(entries, "pdf-security-signature-field-count", "pdf.security.signature", "SignatureFieldCount", security.SignatureFieldCount);
        AddCountMetadata(entries, "pdf-security-signature-count", "pdf.security.signature", "SignatureCount", security.SignatureCount);
        AddCountMetadata(entries, "pdf-security-signature-value-count", "pdf.security.signature", "SignatureValueCount", security.SignatureValueCount);
        AddCountMetadata(entries, "pdf-security-byte-range-segment-count", "pdf.security.signature", "ByteRangeSegmentCount", security.ByteRangeSegmentCount);
        AddCountMetadata(entries, "pdf-security-dss-vri-count", "pdf.security.dss", "VriEntryCount", security.DocumentSecurityStore.VriEntryCount);
        AddCountMetadata(entries, "pdf-security-dss-evidence-count", "pdf.security.dss", "EvidenceObjectCount", security.DocumentSecurityStore.TopLevelEvidenceObjectCount + security.DocumentSecurityStore.VriEvidenceObjectCount);
        AddCountMetadata(entries, "pdf-security-revision-count", "pdf.security.revision", "RevisionCount", security.RevisionCount);
        AddCountMetadata(entries, "pdf-security-startxref-count", "pdf.security.revision", "StartXrefCount", security.StartXrefCount);

        entries.Add(BuildSecurityStateMetadataEntry(security));

        for (int i = 0; i < security.Signatures.Count; i++) {
            entries.Add(BuildSignatureMetadataEntry(security.Signatures[i], i));
        }

        if (security.DocumentSecurityStore.HasDss) {
            entries.Add(BuildDocumentSecurityStoreMetadataEntry(security.DocumentSecurityStore));
        }
    }

    private static OfficeDocumentMetadataEntry BuildSecurityStateMetadataEntry(PdfDocumentSecurityInfo security) {
        var attributes = new Dictionary<string, string>(StringComparer.Ordinal) {
            ["hasEncryption"] = ToMetadataText(security.HasEncryption),
            ["hasSignatures"] = ToMetadataText(security.HasSignatures),
            ["hasByteRange"] = ToMetadataText(security.HasByteRange),
            ["acroFormSignaturesExist"] = ToMetadataText(security.AcroFormSignaturesExist),
            ["acroFormAppendOnly"] = ToMetadataText(security.AcroFormAppendOnly),
            ["hasDocMDPPermissions"] = ToMetadataText(security.HasDocMDPPermissions),
            ["hasUsageRights"] = ToMetadataText(security.HasUsageRights),
            ["hasDocumentSecurityStore"] = ToMetadataText(security.HasDocumentSecurityStore),
            ["hasLongTermValidationEvidence"] = ToMetadataText(security.HasLongTermValidationEvidence),
            ["requiresAppendOnlyMutation"] = ToMetadataText(security.RequiresAppendOnlyMutation),
            ["blocksOfficeIMOAppendOnlyMutation"] = ToMetadataText(security.BlocksOfficeIMOAppendOnlyMutation),
            ["blocksOfficeIMOFullRewriteMutation"] = ToMetadataText(security.BlocksOfficeIMOFullRewriteMutation),
            ["hasIncrementalUpdates"] = ToMetadataText(security.HasIncrementalUpdates),
            ["hasPreviousRevision"] = ToMetadataText(security.HasPreviousRevision),
            ["hasXrefStreams"] = ToMetadataText(security.HasXrefStreams),
            ["hasObjectStreams"] = ToMetadataText(security.HasObjectStreams),
            ["hasTrailerId"] = ToMetadataText(security.HasTrailerId),
            ["signatureFieldCount"] = security.SignatureFieldCount.ToString(CultureInfo.InvariantCulture),
            ["signatureCount"] = security.SignatureCount.ToString(CultureInfo.InvariantCulture),
            ["signatureValueCount"] = security.SignatureValueCount.ToString(CultureInfo.InvariantCulture),
            ["byteRangeValueCount"] = security.ByteRangeValueCount.ToString(CultureInfo.InvariantCulture),
            ["byteRangeSegmentCount"] = security.ByteRangeSegmentCount.ToString(CultureInfo.InvariantCulture),
            ["startXrefCount"] = security.StartXrefCount.ToString(CultureInfo.InvariantCulture),
            ["revisionCount"] = security.RevisionCount.ToString(CultureInfo.InvariantCulture)
        };

        AddAttribute(attributes, "encryptObjectNumber", security.EncryptObjectNumber);
        AddAttribute(attributes, "encryptionFilter", security.EncryptionFilter);
        AddAttribute(attributes, "encryptionSubFilter", security.EncryptionSubFilter);
        AddAttribute(attributes, "encryptionVersion", security.EncryptionVersion);
        AddAttribute(attributes, "encryptionRevision", security.EncryptionRevision);
        AddAttribute(attributes, "encryptionLengthBits", security.EncryptionLengthBits);
        AddAttribute(attributes, "encryptionPermissions", security.EncryptionPermissions);
        AddAttribute(attributes, "encryptMetadata", ToMetadataText(security.EncryptMetadata));
        AddAttribute(attributes, "allowsPrinting", ToMetadataText(security.AllowsPrinting));
        AddAttribute(attributes, "allowsModification", ToMetadataText(security.AllowsModification));
        AddAttribute(attributes, "allowsCopying", ToMetadataText(security.AllowsCopying));
        AddAttribute(attributes, "allowsAnnotationChanges", ToMetadataText(security.AllowsAnnotationChanges));
        AddAttribute(attributes, "allowsFormFilling", ToMetadataText(security.AllowsFormFilling));
        AddAttribute(attributes, "allowsAccessibilityExtraction", ToMetadataText(security.AllowsAccessibilityExtraction));
        AddAttribute(attributes, "allowsDocumentAssembly", ToMetadataText(security.AllowsDocumentAssembly));
        AddAttribute(attributes, "allowsHighQualityPrinting", ToMetadataText(security.AllowsHighQualityPrinting));
        AddAttribute(attributes, "signatureFieldObjectNumbers", FormatPdfIntegerComponents(security.SignatureFieldObjectNumbers));
        AddAttribute(attributes, "signatureFieldNames", FormatPdfStringComponents(security.SignatureFieldNames));
        AddAttribute(attributes, "acroFormSignatureFlags", security.AcroFormSignatureFlags);
        AddAttribute(attributes, "docMDPSignatureObjectNumber", security.DocMDPSignatureObjectNumber);
        AddAttribute(attributes, "docMDPTransformMethod", security.DocMDPTransformMethod);
        AddAttribute(attributes, "docMDPTransformVersion", security.DocMDPTransformVersion);
        AddAttribute(attributes, "docMDPPermissionLevel", security.DocMDPPermissionLevel);
        AddAttribute(attributes, "usageRightsObjectNumbers", FormatPdfIntegerComponents(security.UsageRightsObjectNumbers));
        AddAttribute(attributes, "rootObjectNumber", security.RootObjectNumber);
        AddAttribute(attributes, "rootObjectGeneration", security.RootObjectGeneration);
        AddAttribute(attributes, "infoObjectNumber", security.InfoObjectNumber);
        AddAttribute(attributes, "infoObjectGeneration", security.InfoObjectGeneration);
        AddAttribute(attributes, "lastStartXrefOffset", security.LastStartXrefOffset);
        AddAttribute(attributes, "startXrefOffsets", FormatPdfIntegerComponents(security.StartXrefOffsets));
        AddAttribute(attributes, "previousXrefOffsets", FormatPdfIntegerComponents(security.PreviousXrefOffsets));

        return new OfficeDocumentMetadataEntry {
            Id = "pdf-security-state",
            Category = "pdf.security",
            Name = "SecurityState",
            Value = security.RequiresAppendOnlyMutation ? "AppendOnlyRequired" : "SecurityState",
            ValueType = "object",
            SourceObjectId = security.RootObjectNumber.HasValue
                ? security.RootObjectNumber.Value.ToString(CultureInfo.InvariantCulture)
                : null,
            Attributes = attributes
        };
    }

    private static OfficeDocumentMetadataEntry BuildSignatureMetadataEntry(PdfSignatureInfo signature, int signatureIndex) {
        string id = "pdf-security-signature-" + signatureIndex.ToString("D4", CultureInfo.InvariantCulture);
        var attributes = new Dictionary<string, string>(StringComparer.Ordinal) {
            ["objectNumber"] = signature.ObjectNumber.ToString(CultureInfo.InvariantCulture),
            ["hasFieldLock"] = ToMetadataText(signature.HasFieldLock),
            ["hasSeedValue"] = ToMetadataText(signature.HasSeedValue),
            ["hasByteRange"] = ToMetadataText(signature.HasByteRange),
            ["byteRangeValueCount"] = signature.ByteRangeValueCount.ToString(CultureInfo.InvariantCulture),
            ["byteRangeSegmentCount"] = signature.ByteRangeSegmentCount.ToString(CultureInfo.InvariantCulture),
            ["hasContents"] = ToMetadataText(signature.HasContents),
            ["hasNonEmptyContents"] = ToMetadataText(signature.HasNonEmptyContents),
            ["referenceCount"] = signature.ReferenceCount.ToString(CultureInfo.InvariantCulture),
            ["hasRecognizedSubFilter"] = ToMetadataText(signature.HasRecognizedSubFilter),
            ["usesDetachedCmsSubFilter"] = ToMetadataText(signature.UsesDetachedCmsSubFilter),
            ["usesCadesSubFilter"] = ToMetadataText(signature.UsesCadesSubFilter),
            ["isDocumentTimestamp"] = ToMetadataText(signature.IsDocumentTimestamp)
        };

        AddAttribute(attributes, "fieldObjectNumber", signature.FieldObjectNumber);
        AddAttribute(attributes, "fieldName", signature.FieldName);
        AddAttribute(attributes, "filter", signature.Filter);
        AddAttribute(attributes, "subFilter", signature.SubFilter);
        AddAttribute(attributes, "signerName", signature.SignerName);
        AddAttribute(attributes, "location", signature.Location);
        AddAttribute(attributes, "reason", signature.Reason);
        AddAttribute(attributes, "contactInfo", signature.ContactInfo);
        AddAttribute(attributes, "signingTimeRaw", signature.SigningTimeRaw);
        AddAttribute(attributes, "byteRangeValues", FormatPdfLongComponents(signature.ByteRangeValues));
        AddAttribute(attributes, "contentsSizeBytes", signature.ContentsSizeBytes);
        AddAttribute(attributes, "contentsEncodedSizeBytes", signature.ContentsEncodedSizeBytes);

        if (signature.FieldLock != null) {
            AddAttribute(attributes, "fieldLockAction", signature.FieldLock.Action);
            AddAttribute(attributes, "fieldLockFields", FormatPdfStringComponents(signature.FieldLock.Fields));
            AddAttribute(attributes, "fieldLockLocksAllFields", ToMetadataText(signature.FieldLock.LocksAllFields));
            AddAttribute(attributes, "fieldLockLocksIncludedFields", ToMetadataText(signature.FieldLock.LocksIncludedFields));
            AddAttribute(attributes, "fieldLockLocksAllExceptListedFields", ToMetadataText(signature.FieldLock.LocksAllExceptListedFields));
        }

        if (signature.SeedValue != null) {
            AddAttribute(attributes, "seedValueFilter", signature.SeedValue.Filter);
            AddAttribute(attributes, "seedValueSubFilters", FormatPdfStringComponents(signature.SeedValue.SubFilters));
            AddAttribute(attributes, "seedValueDigestMethods", FormatPdfStringComponents(signature.SeedValue.DigestMethods));
            AddAttribute(attributes, "seedValueReasons", FormatPdfStringComponents(signature.SeedValue.Reasons));
            AddAttribute(attributes, "seedValueFlags", signature.SeedValue.Flags);
            AddAttribute(attributes, "seedValueAddRevInfo", ToMetadataText(signature.SeedValue.AddRevInfo));
            AddAttribute(attributes, "seedValueMdpPermissionLevel", signature.SeedValue.MDPPermissionLevel);
        }

        return new OfficeDocumentMetadataEntry {
            Id = id,
            Category = "pdf.security.signature",
            Name = signature.FieldName ?? signature.SignerName ?? "Signature",
            Value = signature.SignerName ?? signature.Reason ?? signature.SubFilter,
            ValueType = "object",
            SourceObjectId = signature.ObjectNumber.ToString(CultureInfo.InvariantCulture),
            Attributes = attributes
        };
    }

    private static OfficeDocumentMetadataEntry BuildDocumentSecurityStoreMetadataEntry(PdfDocumentDssInfo documentSecurityStore) {
        var attributes = new Dictionary<string, string>(StringComparer.Ordinal) {
            ["hasDss"] = ToMetadataText(documentSecurityStore.HasDss),
            ["hasValidationEvidence"] = ToMetadataText(documentSecurityStore.HasValidationEvidence),
            ["vriEntryCount"] = documentSecurityStore.VriEntryCount.ToString(CultureInfo.InvariantCulture),
            ["topLevelEvidenceObjectCount"] = documentSecurityStore.TopLevelEvidenceObjectCount.ToString(CultureInfo.InvariantCulture),
            ["vriEvidenceObjectCount"] = documentSecurityStore.VriEvidenceObjectCount.ToString(CultureInfo.InvariantCulture)
        };

        AddAttribute(attributes, "objectNumber", documentSecurityStore.ObjectNumber);
        AddAttribute(attributes, "vriKeys", FormatPdfStringComponents(documentSecurityStore.VriKeys));
        AddAttribute(attributes, "certificateObjectNumbers", FormatPdfIntegerComponents(documentSecurityStore.CertificateObjectNumbers));
        AddAttribute(attributes, "ocspObjectNumbers", FormatPdfIntegerComponents(documentSecurityStore.OcspObjectNumbers));
        AddAttribute(attributes, "crlObjectNumbers", FormatPdfIntegerComponents(documentSecurityStore.CrlObjectNumbers));
        AddAttribute(attributes, "vriCertificateObjectNumbers", FormatPdfIntegerComponents(documentSecurityStore.VriCertificateObjectNumbers));
        AddAttribute(attributes, "vriOcspObjectNumbers", FormatPdfIntegerComponents(documentSecurityStore.VriOcspObjectNumbers));
        AddAttribute(attributes, "vriCrlObjectNumbers", FormatPdfIntegerComponents(documentSecurityStore.VriCrlObjectNumbers));
        AddAttribute(attributes, "timestampObjectNumbers", FormatPdfIntegerComponents(documentSecurityStore.TimestampObjectNumbers));

        return new OfficeDocumentMetadataEntry {
            Id = "pdf-security-dss",
            Category = "pdf.security.dss",
            Name = "DocumentSecurityStore",
            Value = FormatPdfStringComponents(documentSecurityStore.VriKeys),
            ValueType = "object",
            SourceObjectId = documentSecurityStore.ObjectNumber.HasValue
                ? documentSecurityStore.ObjectNumber.Value.ToString(CultureInfo.InvariantCulture)
                : null,
            Attributes = attributes
        };
    }

    private static string? FormatPdfLongComponents(IReadOnlyList<long> components) {
        if (components.Count == 0) {
            return null;
        }

        var builder = new StringBuilder();
        for (int i = 0; i < components.Count; i++) {
            if (i > 0) {
                builder.Append(',');
            }

            builder.Append(components[i].ToString(CultureInfo.InvariantCulture));
        }

        return builder.ToString();
    }
}
