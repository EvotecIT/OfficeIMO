using System.Xml;
using System.Xml.Linq;
using System.Threading;
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

    private PdfXmpMetadataInfo? ExtractXmpMetadata(System.Threading.CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
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
            decoded = _decodedStreamBudget.DecodeRequired(stream, _objects, MaxXmpMetadataBytes, cancellationToken);
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
        string? rawXml = decodedWithinLimit ? DecodeMetadataText(decoded, cancellationToken) : null;
        int decodedSizeBytes = decodedWithinLimit ? decoded.Length : MaxXmpMetadataBytes + 1;
        cancellationToken.ThrowIfCancellationRequested();
        XDocument? document = rawXml is null ? null : TryParseXml(rawXml, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return new PdfXmpMetadataInfo(
            objectNumber,
            TryReadName(stream.Dictionary, "Type"),
            TryReadName(stream.Dictionary, "Subtype"),
            TryReadStreamFilter(stream, cancellationToken),
            stream.DataLength,
            decodedSizeBytes,
            StreamDecoder.GetUnsupportedFilters(stream.Dictionary, _objects).AsReadOnly(),
            rawXml,
            document is not null,
            document is null ? null : ReadAltText(document, "title", cancellationToken),
            document is null ? null : ReadFirstCollectionText(document, "creator", cancellationToken),
            document is null ? null : ReadAltText(document, "description", cancellationToken),
            document is null ? Array.Empty<string>() : ReadCollectionText(document, "subject", cancellationToken),
            document is null ? null : ReadElementTextByNamespace(document, "Producer", PdfNamespaceUri, cancellationToken),
            document is null ? null : ReadElementTextByNamespace(document, "Keywords", PdfNamespaceUri, cancellationToken),
            document is null ? null : ReadIntegerElementByNamespace(document, "part", PdfAIdentificationNamespaceUri, cancellationToken),
            document is null ? null : ReadElementTextByNamespace(document, "conformance", PdfAIdentificationNamespaceUri, cancellationToken),
            document is null ? null : ReadIntegerElementByNamespace(document, "part", PdfUaIdentification.NamespaceUri, cancellationToken),
            document is null ? null : ReadElementTextByNamespace(document, "GTS_PDFXVersion", PdfXIdentification.NamespaceUri, cancellationToken),
            document is null ? null : ReadElementTextByNamespace(document, "GTS_PDFXConformance", PdfXIdentification.NamespaceUri, cancellationToken),
            document is null ? null : ReadDateElementByNamespace(document, "CreateDate", XmpNamespaceUri, cancellationToken),
            document is null ? null : ReadDateElementByNamespace(document, "ModifyDate", XmpNamespaceUri, cancellationToken),
            document is null ? null : ReadDateElementByNamespace(document, "MetadataDate", XmpNamespaceUri, cancellationToken),
            document is null ? null : ReadElementTextByNamespace(document, "DocumentID", XmpMediaManagementNamespaceUri, cancellationToken),
            document is null ? null : ReadElementTextByNamespace(document, "InstanceID", XmpMediaManagementNamespaceUri, cancellationToken),
            document is null ? null : ReadElementTextByNamespace(document, "VersionID", XmpMediaManagementNamespaceUri, cancellationToken),
            document is null ? null : ReadElementTextByNamespace(document, "RenditionClass", XmpMediaManagementNamespaceUri, cancellationToken),
            document is null ? null : ParseTrappingStatus(ReadElementTextByNamespace(document, "Trapped", PdfNamespaceUri, cancellationToken)),
            document is null ? null : ReadElementTextByNamespace(document, "DocumentType", PdfElectronicInvoiceMetadata.FacturXNamespaceUri, cancellationToken),
            document is null ? null : ReadElementTextByNamespace(document, "DocumentFileName", PdfElectronicInvoiceMetadata.FacturXNamespaceUri, cancellationToken),
            document is null ? null : ReadElementTextByNamespace(document, "Version", PdfElectronicInvoiceMetadata.FacturXNamespaceUri, cancellationToken),
            document is null ? null : ReadElementTextByNamespace(document, "ConformanceLevel", PdfElectronicInvoiceMetadata.FacturXNamespaceUri, cancellationToken));
    }

    private static string? DecodeMetadataText(byte[] data, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (data.Length == 0) {
            return string.Empty;
        }

        try {
            if (data.Length >= 3 &&
                data[0] == 0xEF &&
                data[1] == 0xBB &&
                data[2] == 0xBF) {
                return PdfEncoding.DecodeCancellable(StrictUtf8, data, 3, data.Length - 3, cancellationToken);
            }

            if (data.Length >= 2 &&
                data[0] == 0xFE &&
                data[1] == 0xFF) {
                return PdfEncoding.DecodeCancellable(StrictBigEndianUnicode, data, 2, data.Length - 2, cancellationToken);
            }

            if (data.Length >= 2 &&
                data[0] == 0xFF &&
                data[1] == 0xFE) {
                return PdfEncoding.DecodeCancellable(StrictLittleEndianUnicode, data, 2, data.Length - 2, cancellationToken);
            }

            return PdfEncoding.DecodeCancellable(StrictUtf8, data, 0, data.Length, cancellationToken);
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

    private static XDocument? TryParseXml(string? rawXml, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (string.IsNullOrWhiteSpace(rawXml)) {
            return null;
        }

        try {
            var settings = new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                MaxCharactersInDocument = MaxXmpMetadataBytes,
                XmlResolver = null
            };
            using var stringReader = new CancellationCheckingTextReader(rawXml!, cancellationToken);
            using XmlReader reader = XmlReader.Create(stringReader, settings);
            XDocument document = XDocument.Load(reader, LoadOptions.None);
            cancellationToken.ThrowIfCancellationRequested();
            return document;
        } catch (Exception ex) when (ex is System.Xml.XmlException || ex is InvalidOperationException) {
            return null;
        }
    }

    private static string? ReadAltText(XDocument document, string localName, CancellationToken cancellationToken) {
        XElement? element = FindElementByNamespace(document, localName, DublinCoreNamespaceUri, cancellationToken);
        if (element is null) {
            return null;
        }

        XElement? defaultItem = DescendantsWithCancellation(element, cancellationToken)
            .FirstOrDefault(e => e.Name.LocalName == "li" &&
                string.Equals((string?)e.Attribute(XNamespace.Xml + "lang"), "x-default", StringComparison.OrdinalIgnoreCase));

        return NormalizeXmlText(ReadElementText(defaultItem, cancellationToken), cancellationToken) ??
            NormalizeXmlText(ReadElementText(DescendantsWithCancellation(element, cancellationToken).FirstOrDefault(e => e.Name.LocalName == "li"), cancellationToken), cancellationToken);
    }

    private static string? ReadFirstCollectionText(XDocument document, string localName, CancellationToken cancellationToken) {
        IReadOnlyList<string> values = ReadCollectionText(document, localName, cancellationToken);
        return values.Count == 0 ? null : values[0];
    }

    private static IReadOnlyList<string> ReadCollectionText(XDocument document, string localName, CancellationToken cancellationToken) {
        XElement? element = FindElementByNamespace(document, localName, DublinCoreNamespaceUri, cancellationToken);
        if (element is null) {
            return Array.Empty<string>();
        }

        var values = new List<string>();
        foreach (XElement item in DescendantsWithCancellation(element, cancellationToken).Where(e => e.Name.LocalName == "li")) {
            string? text = NormalizeXmlText(ReadElementText(item, cancellationToken), cancellationToken);
            if (text is not null) {
                values.Add(text);
            }
        }

        return values.Count == 0 ? Array.Empty<string>() : values.AsReadOnly();
    }

    private static string? ReadElementTextByNamespace(XDocument document, string localName, string namespaceUri, CancellationToken cancellationToken) {
        var values = new HashSet<string>(StringComparer.Ordinal);
        foreach (XElement description in FindDocumentSubjectDescriptions(document, cancellationToken)) {
            foreach (XElement element in WithCancellation(description.Elements(), cancellationToken).Where(e =>
                         e.Name.LocalName == localName &&
                         string.Equals(e.Name.NamespaceName, namespaceUri, StringComparison.Ordinal))) {
                string? value = NormalizeXmlText(ReadElementText(element, cancellationToken), cancellationToken);
                if (value is not null) values.Add(value);
            }
            foreach (XAttribute attribute in WithCancellation(description.Attributes(), cancellationToken).Where(a =>
                         a.Name.LocalName == localName &&
                         string.Equals(a.Name.NamespaceName, namespaceUri, StringComparison.Ordinal))) {
                string? value = NormalizeXmlText(attribute.Value, cancellationToken);
                if (value is not null) values.Add(value);
            }
        }
        return values.Count == 1 ? values.Single() : null;
    }

    private static IEnumerable<XElement> FindDocumentSubjectDescriptions(XDocument document, CancellationToken cancellationToken) =>
        DescendantsWithCancellation(document, cancellationToken).Where(e =>
            e.Name.LocalName == "Description" &&
            string.Equals(e.Name.NamespaceName, RdfNamespaceUri, StringComparison.Ordinal) &&
            string.IsNullOrEmpty((string?)e.Attribute(XName.Get("about", RdfNamespaceUri))));

    private static int? ReadIntegerElementByNamespace(XDocument document, string localName, string namespaceUri, CancellationToken cancellationToken) {
        string? value = ReadElementTextByNamespace(document, localName, namespaceUri, cancellationToken);
        return int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int result)
            ? result
            : null;
    }

    private static DateTimeOffset? ReadDateElementByNamespace(XDocument document, string localName, string namespaceUri, CancellationToken cancellationToken) {
        string? value = ReadElementTextByNamespace(document, localName, namespaceUri, cancellationToken);
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

    private static XElement? FindElementByNamespace(XDocument document, string localName, string namespaceUri, CancellationToken cancellationToken) {
        return DescendantsWithCancellation(document, cancellationToken).FirstOrDefault(e =>
            e.Name.LocalName == localName &&
            string.Equals(e.Name.NamespaceName, namespaceUri, StringComparison.Ordinal));
    }

    private static string? ReadElementText(XElement? element, CancellationToken cancellationToken) {
        if (element is null) return null;
        var text = new System.Text.StringBuilder();
        foreach (XText node in WithCancellation(element.DescendantNodes().OfType<XText>(), cancellationToken)) {
            string value = node.Value;
            for (int index = 0; index < value.Length; index += 1024) {
                cancellationToken.ThrowIfCancellationRequested();
                text.Append(value, index, Math.Min(1024, value.Length - index));
            }
        }
        cancellationToken.ThrowIfCancellationRequested();
        return text.ToString();
    }

    private static string? NormalizeXmlText(string? value, CancellationToken cancellationToken) {
        if (value is null) {
            return null;
        }
        int start = 0;
        int end = value.Length;
        while (start < end && char.IsWhiteSpace(value[start])) {
            if ((start & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
            start++;
        }
        if (start == end) return null;
        while (end > start && char.IsWhiteSpace(value[end - 1])) {
            if ((end & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
            end--;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return value.Substring(start, end - start);
    }

    private static IEnumerable<XElement> DescendantsWithCancellation(XContainer container, CancellationToken cancellationToken) =>
        WithCancellation(container.Descendants(), cancellationToken);

    private static IEnumerable<T> WithCancellation<T>(IEnumerable<T> values, CancellationToken cancellationToken) {
        foreach (T value in values) {
            cancellationToken.ThrowIfCancellationRequested();
            yield return value;
        }
    }

    private sealed class CancellationCheckingTextReader : TextReader {
        private readonly StringReader _reader;
        private readonly CancellationToken _cancellationToken;

        internal CancellationCheckingTextReader(string value, CancellationToken cancellationToken) {
            _reader = new StringReader(value);
            _cancellationToken = cancellationToken;
        }

        public override int Peek() {
            _cancellationToken.ThrowIfCancellationRequested();
            return _reader.Peek();
        }

        public override int Read() {
            _cancellationToken.ThrowIfCancellationRequested();
            return _reader.Read();
        }

        public override int Read(char[] buffer, int index, int count) {
            _cancellationToken.ThrowIfCancellationRequested();
            return _reader.Read(buffer, index, Math.Min(count, 4096));
        }

#if NET8_0_OR_GREATER
        public override int Read(Span<char> buffer) {
            _cancellationToken.ThrowIfCancellationRequested();
            return _reader.Read(buffer.Slice(0, Math.Min(buffer.Length, 4096)));
        }
#endif

        protected override void Dispose(bool disposing) {
            if (disposing) _reader.Dispose();
            base.Dispose(disposing);
        }
    }
}
