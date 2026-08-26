using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadDocument {
    private const string DublinCoreNamespaceUri = "http://purl.org/dc/elements/1.1/";
    private const string RdfNamespaceUri = "http://www.w3.org/1999/02/22-rdf-syntax-ns#";
    private const string PdfAIdentificationNamespaceUri = "http://www.aiim.org/pdfa/ns/id/";
    private const string PdfNamespaceUri = "http://ns.adobe.com/pdf/1.3/";
    private const string XmpNamespaceUri = "http://ns.adobe.com/xap/1.0/";
    private const string XmpMediaManagementNamespaceUri = "http://ns.adobe.com/xap/1.0/mm/";
    /// <summary>Maximum decoded XMP metadata size parsed as XML.</summary>
    public const int MaxXmpMetadataBytes = 4_000_000;

    /// <summary>Catalog XMP metadata stream discovered from /Metadata.</summary>
    public PdfXmpMetadataInfo? XmpMetadata => ReadLogicalContent(_xmpMetadata);

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

        byte[] decoded;
        bool decodedWithinLimit;
        try {
            decoded = _decodedStreamBudget.DecodeRequired(stream, _objects, MaxXmpMetadataBytes);
            decodedWithinLimit = true;
        } catch (PdfReadLimitException exception) when (
            exception.Kind == PdfReadLimitKind.DecodedStreamBytes &&
            exception.Limit == MaxXmpMetadataBytes) {
            decoded = Array.Empty<byte>();
            decodedWithinLimit = false;
        } catch (PdfReadLimitException) {
            throw;
        } catch (InvalidDataException) {
            decoded = Array.Empty<byte>();
            decodedWithinLimit = false;
        }
        string? rawXml = decodedWithinLimit ? DecodeMetadataText(decoded) : null;
        int decodedSizeBytes = decodedWithinLimit ? decoded.Length : MaxXmpMetadataBytes + 1;
        XDocument? document = rawXml is null ? null : TryParseXml(rawXml);
        return new PdfXmpMetadataInfo(
            objectNumber,
            TryReadName(stream.Dictionary, "Type"),
            TryReadName(stream.Dictionary, "Subtype"),
            TryReadStreamFilter(stream),
            stream.Data.Length,
            decodedSizeBytes,
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
            document is null ? null : ReadElementTextByNamespace(document, "GTS_PDFXVersion", PdfXIdentification.NamespaceUri),
            document is null ? null : ReadElementTextByNamespace(document, "GTS_PDFXConformance", PdfXIdentification.NamespaceUri),
            document is null ? null : ReadDateElementByNamespace(document, "CreateDate", XmpNamespaceUri),
            document is null ? null : ReadDateElementByNamespace(document, "ModifyDate", XmpNamespaceUri),
            document is null ? null : ReadDateElementByNamespace(document, "MetadataDate", XmpNamespaceUri),
            document is null ? null : ReadElementTextByNamespace(document, "DocumentID", XmpMediaManagementNamespaceUri),
            document is null ? null : ReadElementTextByNamespace(document, "InstanceID", XmpMediaManagementNamespaceUri),
            document is null ? null : ReadElementTextByNamespace(document, "VersionID", XmpMediaManagementNamespaceUri),
            document is null ? null : ReadElementTextByNamespace(document, "RenditionClass", XmpMediaManagementNamespaceUri),
            document is null ? null : ParseTrappingStatus(ReadElementTextByNamespace(document, "Trapped", PdfNamespaceUri)),
            document is null ? null : ReadElementTextByNamespace(document, "DocumentType", PdfElectronicInvoiceMetadata.FacturXNamespaceUri),
            document is null ? null : ReadElementTextByNamespace(document, "DocumentFileName", PdfElectronicInvoiceMetadata.FacturXNamespaceUri),
            document is null ? null : ReadElementTextByNamespace(document, "Version", PdfElectronicInvoiceMetadata.FacturXNamespaceUri),
            document is null ? null : ReadElementTextByNamespace(document, "ConformanceLevel", PdfElectronicInvoiceMetadata.FacturXNamespaceUri));
    }

    private static string? DecodeMetadataText(byte[] data) {
        if (data.Length == 0) {
            return string.Empty;
        }

        try {
            if (data.Length >= 3 &&
                data[0] == 0xEF &&
                data[1] == 0xBB &&
                data[2] == 0xBF) {
                return StrictUtf8.GetString(data, 3, data.Length - 3);
            }

            if (data.Length >= 2 &&
                data[0] == 0xFE &&
                data[1] == 0xFF) {
                return StrictBigEndianUnicode.GetString(data, 2, data.Length - 2);
            }

            if (data.Length >= 2 &&
                data[0] == 0xFF &&
                data[1] == 0xFE) {
                return StrictLittleEndianUnicode.GetString(data, 2, data.Length - 2);
            }

            return StrictUtf8.GetString(data);
        } catch (DecoderFallbackException) {
            return null;
        }
    }

    private static readonly Encoding StrictUtf8 = new UTF8Encoding(
        encoderShouldEmitUTF8Identifier: false,
        throwOnInvalidBytes: true);
    private static readonly Encoding StrictBigEndianUnicode = new UnicodeEncoding(
        bigEndian: true,
        byteOrderMark: false,
        throwOnInvalidBytes: true);
    private static readonly Encoding StrictLittleEndianUnicode = new UnicodeEncoding(
        bigEndian: false,
        byteOrderMark: false,
        throwOnInvalidBytes: true);

    private static XDocument? TryParseXml(string? rawXml) {
        if (string.IsNullOrWhiteSpace(rawXml)) {
            return null;
        }

        try {
            var settings = new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                MaxCharactersInDocument = MaxXmpMetadataBytes,
                XmlResolver = null
            };
            using var stringReader = new StringReader(rawXml!);
            using XmlReader reader = XmlReader.Create(stringReader, settings);
            return XDocument.Load(reader, LoadOptions.None);
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
        string? elementValue = NormalizeXmlText(document.Descendants().FirstOrDefault(e => e.Name.LocalName == localName)?.Value);
        return elementValue ?? NormalizeXmlText(FindRdfDescriptionAttribute(document, localName, namespaceUri: null)?.Value);
    }

    private static string? ReadElementTextByNamespace(XDocument document, string localName, string namespaceUri) {
        var values = new HashSet<string>(StringComparer.Ordinal);
        foreach (XElement description in FindDocumentSubjectDescriptions(document)) {
            foreach (XElement element in description.Elements().Where(e =>
                         e.Name.LocalName == localName &&
                         string.Equals(e.Name.NamespaceName, namespaceUri, StringComparison.Ordinal))) {
                string? value = NormalizeXmlText(element.Value);
                if (value is not null) values.Add(value);
            }
            foreach (XAttribute attribute in description.Attributes().Where(a =>
                         a.Name.LocalName == localName &&
                         string.Equals(a.Name.NamespaceName, namespaceUri, StringComparison.Ordinal))) {
                string? value = NormalizeXmlText(attribute.Value);
                if (value is not null) values.Add(value);
            }
        }
        return values.Count == 1 ? values.Single() : null;
    }

    private static XAttribute? FindRdfDescriptionAttribute(
        XDocument document,
        string localName,
        string? namespaceUri) =>
        FindDocumentSubjectDescriptions(document)
            .Attributes()
            .FirstOrDefault(a =>
                a.Name.LocalName == localName &&
                (namespaceUri is null || string.Equals(a.Name.NamespaceName, namespaceUri, StringComparison.Ordinal)));

    private static IEnumerable<XElement> FindDocumentSubjectDescriptions(XDocument document) =>
        document.Descendants().Where(e =>
            e.Name.LocalName == "Description" &&
            string.Equals(e.Name.NamespaceName, RdfNamespaceUri, StringComparison.Ordinal) &&
            string.IsNullOrEmpty((string?)e.Attribute(XName.Get("about", RdfNamespaceUri))));

    private static int? ReadIntegerElementByNamespace(XDocument document, string localName, string namespaceUri) {
        string? value = ReadElementTextByNamespace(document, localName, namespaceUri);
        return int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int result)
            ? result
            : null;
    }

    private static DateTimeOffset? ReadDateElementByNamespace(XDocument document, string localName, string namespaceUri) {
        string? value = ReadElementTextByNamespace(document, localName, namespaceUri);
        if (value is null) return null;
        string[] formats;
        System.Globalization.DateTimeStyles styles;
        if (value.Length > 0 && value[value.Length - 1] == 'Z') {
            formats = new[] {
                "yyyy-MM-dd'T'HH:mm:ss'Z'",
                "yyyy-MM-dd'T'HH:mm:ss.FFFFFFF'Z'"
            };
            styles = System.Globalization.DateTimeStyles.AssumeUniversal |
                System.Globalization.DateTimeStyles.AdjustToUniversal;
        } else {
            formats = new[] {
                "yyyy-MM-dd'T'HH:mm:sszzz",
                "yyyy-MM-dd'T'HH:mm:ss.FFFFFFFzzz"
            };
            styles = System.Globalization.DateTimeStyles.None;
        }
        return DateTimeOffset.TryParseExact(
            value,
            formats,
            System.Globalization.CultureInfo.InvariantCulture,
            styles,
            out DateTimeOffset result)
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
