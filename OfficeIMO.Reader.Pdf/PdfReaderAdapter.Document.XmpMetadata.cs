using OfficeIMO.Pdf;
using System.Globalization;

namespace OfficeIMO.Reader.Pdf;

internal static partial class PdfReaderAdapter {
    private static void AddXmpMetadata(List<OfficeDocumentMetadataEntry> entries, PdfXmpMetadataInfo? xmpMetadata) {
        if (xmpMetadata == null) {
            return;
        }

        AddCountMetadata(entries, "pdf-xmp-metadata-count", "pdf.xmp", "Count", 1);
        AddCountMetadata(entries, "pdf-xmp-unsupported-filter-count", "pdf.xmp", "UnsupportedFilterCount", xmpMetadata.UnsupportedFilters.Count);
        AddCountMetadata(entries, "pdf-xmp-pdfa-identification-count", "pdf.xmp", "PdfAIdentificationCount", xmpMetadata.HasPdfAIdentification ? 1 : 0);
        AddCountMetadata(entries, "pdf-xmp-pdfua-identification-count", "pdf.xmp", "PdfUaIdentificationCount", xmpMetadata.HasPdfUaIdentification ? 1 : 0);
        AddCountMetadata(entries, "pdf-xmp-electronic-invoice-metadata-count", "pdf.xmp", "ElectronicInvoiceMetadataCount", xmpMetadata.HasElectronicInvoiceMetadata ? 1 : 0);
        entries.Add(BuildXmpMetadataEntry(xmpMetadata));
    }

    private static OfficeDocumentMetadataEntry BuildXmpMetadataEntry(PdfXmpMetadataInfo xmpMetadata) {
        var attributes = new Dictionary<string, string>(StringComparer.Ordinal) {
            ["streamSizeBytes"] = xmpMetadata.StreamSizeBytes.ToString(CultureInfo.InvariantCulture),
            ["decodedSizeBytes"] = xmpMetadata.DecodedSizeBytes.ToString(CultureInfo.InvariantCulture),
            ["isWellFormedXml"] = ToMetadataText(xmpMetadata.IsWellFormedXml)
        };

        AddAttribute(attributes, "objectNumber", xmpMetadata.ObjectNumber);
        AddAttribute(attributes, "subtype", xmpMetadata.Subtype);
        AddAttribute(attributes, "filter", xmpMetadata.Filter);
        AddAttribute(attributes, "unsupportedFilters", FormatPdfStringComponents(xmpMetadata.UnsupportedFilters));
        AddAttribute(attributes, "title", xmpMetadata.Title);
        AddAttribute(attributes, "creator", xmpMetadata.Creator);
        AddAttribute(attributes, "description", xmpMetadata.Description);
        AddAttribute(attributes, "subjects", FormatPdfStringComponents(xmpMetadata.Subjects));
        AddAttribute(attributes, "producer", xmpMetadata.Producer);
        AddAttribute(attributes, "keywords", xmpMetadata.Keywords);
        AddAttribute(attributes, "pdfAPart", xmpMetadata.PdfAPart);
        AddAttribute(attributes, "pdfAConformance", xmpMetadata.PdfAConformance);
        AddAttribute(attributes, "pdfUaPart", xmpMetadata.PdfUaPart);
        AddAttribute(attributes, "electronicInvoiceDocumentType", xmpMetadata.ElectronicInvoiceDocumentType);
        AddAttribute(attributes, "electronicInvoiceDocumentFileName", xmpMetadata.ElectronicInvoiceDocumentFileName);
        AddAttribute(attributes, "electronicInvoiceVersion", xmpMetadata.ElectronicInvoiceVersion);
        AddAttribute(attributes, "electronicInvoiceConformanceLevel", xmpMetadata.ElectronicInvoiceConformanceLevel);

        return new OfficeDocumentMetadataEntry {
            Id = "pdf-xmp-metadata",
            Category = "pdf.xmp",
            Name = "XmpMetadata",
            Value = xmpMetadata.Title ?? xmpMetadata.Description ?? xmpMetadata.Producer,
            ValueType = "object",
            SourceObjectId = xmpMetadata.ObjectNumber.HasValue
                ? xmpMetadata.ObjectNumber.Value.ToString(CultureInfo.InvariantCulture)
                : null,
            Attributes = attributes
        };
    }
}
