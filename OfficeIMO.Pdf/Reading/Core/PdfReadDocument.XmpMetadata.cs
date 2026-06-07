using System.Xml.Linq;
using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadDocument {
    private const string DublinCoreNamespaceUri = "http://purl.org/dc/elements/1.1/";
    private const string PdfAIdentificationNamespaceUri = "http://www.aiim.org/pdfa/ns/id/";

    /// <summary>Catalog XMP metadata stream discovered from /Metadata.</summary>
    public PdfXmpMetadataInfo? XmpMetadata { get; }

    private PdfXmpMetadataInfo? ExtractXmpMetadata() {
        PdfDictionary? catalog = FindCatalog();
        if (catalog is null ||
            !catalog.Items.TryGetValue("Metadata", out PdfObject? metadataObject)) {
            return null;
        }

        int? objectNumber = metadataObject is PdfReference reference ? reference.ObjectNumber : null;
        if (ResolveObject(metadataObject) is not PdfStream stream) {
            return null;
        }

        byte[] decoded = StreamDecoder.Decode(stream.Dictionary, stream.Data, _objects);
        string? rawXml = DecodeMetadataText(decoded);
        XDocument? document = TryParseXml(rawXml);
        return new PdfXmpMetadataInfo(
            objectNumber,
            TryReadName(stream.Dictionary, "Subtype"),
            TryReadStreamFilter(stream),
            stream.Data.Length,
            decoded.Length,
            StreamDecoder.GetUnsupportedFilters(stream.Dictionary, _objects).AsReadOnly(),
            rawXml,
            document is not null,
            document is null ? null : ReadAltText(document, "title"),
            document is null ? null : ReadFirstCollectionText(document, "creator"),
            document is null ? null : ReadAltText(document, "description"),
            document is null ? Array.Empty<string>() : ReadCollectionText(document, "subject"),
            document is null ? null : ReadElementText(document, "Producer"),
            document is null ? null : ReadElementText(document, "Keywords"),
            document is null ? null : ReadIntegerElementByNamespace(document, "part", PdfAIdentificationNamespaceUri),
            document is null ? null : ReadElementTextByNamespace(document, "conformance", PdfAIdentificationNamespaceUri),
            document is null ? null : ReadIntegerElementByNamespace(document, "part", PdfUaIdentification.NamespaceUri),
            document is null ? null : ReadElementTextByNamespace(document, "DocumentType", PdfElectronicInvoiceMetadata.FacturXNamespaceUri),
            document is null ? null : ReadElementTextByNamespace(document, "DocumentFileName", PdfElectronicInvoiceMetadata.FacturXNamespaceUri),
            document is null ? null : ReadElementTextByNamespace(document, "Version", PdfElectronicInvoiceMetadata.FacturXNamespaceUri),
            document is null ? null : ReadElementTextByNamespace(document, "ConformanceLevel", PdfElectronicInvoiceMetadata.FacturXNamespaceUri));
    }

    private static string? DecodeMetadataText(byte[] data) {
        if (data.Length == 0) {
            return string.Empty;
        }

        if (data.Length >= 3 &&
            data[0] == 0xEF &&
            data[1] == 0xBB &&
            data[2] == 0xBF) {
            return Encoding.UTF8.GetString(data, 3, data.Length - 3);
        }

        if (data.Length >= 2 &&
            data[0] == 0xFE &&
            data[1] == 0xFF) {
            return Encoding.BigEndianUnicode.GetString(data, 2, data.Length - 2);
        }

        if (data.Length >= 2 &&
            data[0] == 0xFF &&
            data[1] == 0xFE) {
            return Encoding.Unicode.GetString(data, 2, data.Length - 2);
        }

        return Encoding.UTF8.GetString(data);
    }

    private static XDocument? TryParseXml(string? rawXml) {
        if (string.IsNullOrWhiteSpace(rawXml)) {
            return null;
        }

        try {
            return XDocument.Parse(rawXml!, LoadOptions.None);
        } catch (Exception ex) when (ex is System.Xml.XmlException || ex is InvalidOperationException) {
            return null;
        }
    }

    private static string? ReadAltText(XDocument document, string localName) {
        XElement? element = FindElementByNamespace(document, localName, DublinCoreNamespaceUri);
        if (element is null) {
            return null;
        }

        XElement? defaultItem = element
            .Descendants()
            .FirstOrDefault(e => e.Name.LocalName == "li" &&
                string.Equals((string?)e.Attribute(XNamespace.Xml + "lang"), "x-default", StringComparison.OrdinalIgnoreCase));

        return NormalizeXmlText(defaultItem?.Value) ?? NormalizeXmlText(element.Descendants().FirstOrDefault(e => e.Name.LocalName == "li")?.Value);
    }

    private static string? ReadFirstCollectionText(XDocument document, string localName) {
        IReadOnlyList<string> values = ReadCollectionText(document, localName);
        return values.Count == 0 ? null : values[0];
    }

    private static IReadOnlyList<string> ReadCollectionText(XDocument document, string localName) {
        XElement? element = FindElementByNamespace(document, localName, DublinCoreNamespaceUri);
        if (element is null) {
            return Array.Empty<string>();
        }

        var values = new List<string>();
        foreach (XElement item in element.Descendants().Where(e => e.Name.LocalName == "li")) {
            string? text = NormalizeXmlText(item.Value);
            if (text is not null) {
                values.Add(text);
            }
        }

        return values.Count == 0 ? Array.Empty<string>() : values.AsReadOnly();
    }

    private static string? ReadElementText(XDocument document, string localName) {
        return NormalizeXmlText(document.Descendants().FirstOrDefault(e => e.Name.LocalName == localName)?.Value);
    }

    private static string? ReadElementTextByNamespace(XDocument document, string localName, string namespaceUri) {
        XElement? element = FindElementByNamespace(document, localName, namespaceUri);
        return NormalizeXmlText(element?.Value);
    }

    private static int? ReadIntegerElementByNamespace(XDocument document, string localName, string namespaceUri) {
        string? value = ReadElementTextByNamespace(document, localName, namespaceUri);
        return int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int result)
            ? result
            : null;
    }

    private static XElement? FindElementByNamespace(XDocument document, string localName, string namespaceUri) {
        return document.Descendants().FirstOrDefault(e =>
            e.Name.LocalName == localName &&
            string.Equals(e.Name.NamespaceName, namespaceUri, StringComparison.Ordinal));
    }

    private static string? NormalizeXmlText(string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return null;
        }

        return value!.Trim();
    }
}
