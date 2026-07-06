using OfficeIMO.Pdf;

namespace OfficeIMO.Reader.Pdf;

public static partial class DocumentReaderPdfExtensions {
    private static IReadOnlyList<OfficeDocumentMetadataEntry> BuildPdfPreflightMetadata(PdfDocumentPreflight? preflight) {
        if (preflight == null) {
            return Array.Empty<OfficeDocumentMetadataEntry>();
        }

        var entries = new List<OfficeDocumentMetadataEntry>();
        AddPdfPreflightCapabilityMetadata(entries, "pdf-preflight-can-read", "CanRead", preflight.CanRead);
        AddPdfPreflightCapabilityMetadata(entries, "pdf-preflight-can-rewrite", "CanRewrite", preflight.CanRewrite);
        AddPdfPreflightCapabilityMetadata(entries, "pdf-preflight-can-extract-text", "CanExtractText", preflight.CanExtractText);
        AddPdfPreflightCapabilityMetadata(entries, "pdf-preflight-can-extract-images", "CanExtractImages", preflight.CanExtractImages);
        AddPdfPreflightCapabilityMetadata(entries, "pdf-preflight-can-read-logical-objects", "CanReadLogicalObjects", preflight.CanReadLogicalObjects);
        AddPdfPreflightCapabilityMetadata(entries, "pdf-preflight-can-manipulate-pages", "CanManipulatePages", preflight.CanManipulatePages);
        AddPdfPreflightCapabilityMetadata(entries, "pdf-preflight-can-append-metadata-revision", "CanAppendMetadataRevision", preflight.CanAppendMetadataRevision);
        AddPdfPreflightCapabilityMetadata(entries, "pdf-preflight-can-append-form-field-revision", "CanAppendFormFieldRevision", preflight.CanAppendFormFieldRevision);
        AddPdfPreflightCapabilityMetadata(entries, "pdf-preflight-can-prepare-external-signature-revision", "CanPrepareExternalSignatureRevision", preflight.CanPrepareExternalSignatureRevision);
        AddPdfPreflightCapabilityMetadata(entries, "pdf-preflight-can-fill-simple-form-fields", "CanFillSimpleFormFields", preflight.CanFillSimpleFormFields);
        AddPdfPreflightCapabilityMetadata(entries, "pdf-preflight-can-flatten-simple-form-fields", "CanFlattenSimpleFormFields", preflight.CanFlattenSimpleFormFields);
        AddPdfPreflightCapabilityMetadata(entries, "pdf-preflight-can-fill-and-flatten-simple-form-fields", "CanFillAndFlattenSimpleFormFields", preflight.CanFillAndFlattenSimpleFormFields);

        return entries.AsReadOnly();
    }

    private static void AddPdfPreflightCapabilityMetadata(List<OfficeDocumentMetadataEntry> entries, string id, string name, bool value) {
        entries.Add(new OfficeDocumentMetadataEntry {
            Id = id,
            Category = "pdf.preflight.capability",
            Name = name,
            Value = ToMetadataText(value),
            ValueType = "boolean"
        });
    }
}
