using System.Globalization;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Invoices;

/// <summary>
/// An immutable CII namespace-100 XML document with bounded header editing.
/// Loading does not establish schema, business-rule, tax, or profile compliance.
/// </summary>
public sealed class CiiInvoiceDocument {
    /// <summary>Maximum accepted XML byte length, including any byte order mark.</summary>
    public const int MaximumXmlBytes = 16 * 1024 * 1024;
    private const int MaximumDepth = 128;
    private static readonly XNamespace Rsm = "urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:100";
    private static readonly XNamespace Ram = "urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100";
    private static readonly XNamespace Udt = "urn:un:unece:uncefact:data:standard:UnqualifiedDataType:100";
    private static readonly XNamespace Dsig = "http://www.w3.org/2000/09/xmldsig#";
    private readonly byte[] _bytes;
    private readonly XDocument _document;

    private CiiInvoiceDocument(byte[] bytes, XDocument document) {
        _bytes = bytes;
        _document = document;
    }

    /// <summary>Invoice identifier from the direct exchanged-document header, if present.</summary>
    public string? DocumentId => ReadScalar(Rsm + "ExchangedDocument", Ram + "ID");

    /// <summary>Declared document type code, without interpreting its profile validity.</summary>
    public string? TypeCode => ReadScalar(Rsm + "ExchangedDocument", Ram + "TypeCode");

    /// <summary>Declared guideline identifier, without claiming that its requirements are satisfied.</summary>
    public string? GuidelineId => ReadScalar(Rsm + "ExchangedDocumentContext", Ram + "GuidelineSpecifiedDocumentContextParameter", Ram + "ID");

    /// <summary>Declared invoice currency, without validating code lists or monetary amounts.</summary>
    public string? CurrencyCode => ReadScalar(Rsm + "SupplyChainTradeTransaction", Ram + "ApplicableHeaderTradeSettlement", Ram + "InvoiceCurrencyCode");

    /// <summary>Issue date for a valid format-102 calendar date; null for missing or other date representations.</summary>
    public DateTime? IssueDate {
        get {
            XElement? value = FindUnique(_document, Rsm + "ExchangedDocument", Ram + "IssueDateTime", Udt + "DateTimeString");
            if (value == null || (string?)value.Attribute("format") != "102") return null;
            return DateTime.TryParseExact(Scalar(value), "yyyyMMdd", CultureInfo.InvariantCulture,
                DateTimeStyles.None, out DateTime date) ? date : (DateTime?)null;
        }
    }

    /// <summary>Whether an XML Signature element is present. Signatures are not verified by this model.</summary>
    public bool HasXmlSignature => _document.Descendants(Dsig + "Signature").Any();

    /// <summary>
    /// Loads a defensive copy of CII XML. DTDs, external entities, excessive depth, and oversized input are rejected.
    /// Unknown elements and attributes are retained. Namespace prefixes may vary; the root namespace must be version 100.
    /// </summary>
    /// <exception cref="InvalidDataException">The size, depth, or root is unsupported.</exception>
    /// <exception cref="XmlException">The input is not well-formed XML or includes a DTD.</exception>
    public static CiiInvoiceDocument Load(byte[] xml) {
        if (xml == null) throw new ArgumentNullException(nameof(xml), "Invoice XML is required.");
        if (xml.Length == 0 || xml.Length > MaximumXmlBytes) {
            throw new InvalidDataException("CII XML must contain between 1 and 16777216 bytes.");
        }
        byte[] bytes = (byte[])xml.Clone();
        // Check depth before constructing a tree, including content that is unknown to the header model.
        using (var stream = new MemoryStream(bytes, false))
        using (XmlReader reader = XmlReader.Create(stream, ReaderSettings())) {
            while (reader.Read()) {
                if (reader.Depth > MaximumDepth) throw new InvalidDataException("CII XML exceeds the supported depth of 128.");
            }
        }
        XDocument document;
        using (var stream = new MemoryStream(bytes, false))
        using (XmlReader reader = XmlReader.Create(stream, ReaderSettings())) {
            document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        }
        if (document.Root?.Name != Rsm + "CrossIndustryInvoice") {
            throw new InvalidDataException("Expected a namespace-100 UN/CEFACT CrossIndustryInvoice root.");
        }
        return new CiiInvoiceDocument(bytes, document);
    }

    /// <summary>
    /// Returns a copy of the stored XML bytes. An unedited document is byte-identical to its input.
    /// Edited documents use UTF-8 and retain XML content, but not original lexical formatting or byte identity.
    /// </summary>
    public byte[] ToBytes() => (byte[])_bytes.Clone();

    /// <summary>
    /// Creates an edited document by replacing one existing, unambiguous invoice identifier.
    /// Other fields, including payment references and visible PDF content, are not updated.
    /// Signed XML and structured or missing identifier fields are rejected without changing the original.
    /// </summary>
    public CiiInvoiceDocument WithDocumentId(string documentId) {
        if (string.IsNullOrWhiteSpace(documentId)) throw new ArgumentException("An invoice identifier is required.", nameof(documentId));
        if (documentId.Length > MaximumXmlBytes) throw new ArgumentException("The invoice identifier exceeds the XML size limit.", nameof(documentId));
        XmlConvert.VerifyXmlChars(documentId);
        return ReplaceScalar(documentId, false, Rsm + "ExchangedDocument", Ram + "ID");
    }

    /// <summary>
    /// Creates an edited document by replacing one existing format-102 issue date.
    /// Time and time zone are ignored; delivery and payment dates are retained.
    /// Signed XML, ambiguous fields, and other date representations are rejected.
    /// </summary>
    public CiiInvoiceDocument WithIssueDate(DateTime issueDate) =>
        ReplaceScalar(issueDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture), true,
            Rsm + "ExchangedDocument", Ram + "IssueDateTime", Udt + "DateTimeString");

    private CiiInvoiceDocument ReplaceScalar(string value, bool requireDateFormat, params XName[] path) {
        if (HasXmlSignature) throw new InvalidOperationException("Editing XML with a Signature element is not supported. Preserve the original bytes or use a signature-aware workflow.");
        var copy = new XDocument(_document);
        XElement element = FindUnique(copy, path) ?? throw new InvalidOperationException("The field must already exist before it can be edited.");
        if (requireDateFormat && (string?)element.Attribute("format") != "102") {
            throw new InvalidOperationException("Only existing format-102 issue dates can be edited.");
        }
        // Do not remove nested extension content, comments, or processing instructions from a field.
        if (element.Nodes().Any(node => node is not XText)) {
            throw new InvalidOperationException("The field contains structured XML content and cannot be replaced as text.");
        }
        element.Value = value;
        using (var stream = new MemoryStream()) {
            using (XmlWriter writer = XmlWriter.Create(stream, new XmlWriterSettings {
                Encoding = new UTF8Encoding(false), Indent = false, NewLineHandling = NewLineHandling.Entitize
            })) {
                copy.Save(writer);
            }
            return Load(stream.ToArray());
        }
    }

    private string? ReadScalar(params XName[] path) {
        XElement? element = FindUnique(_document, path);
        return element == null ? null : Scalar(element);
    }

    private static string Scalar(XElement element) {
        if (element.HasElements) throw new InvalidDataException("A CII scalar field contains nested elements.");
        return element.Value;
    }

    private static XElement? FindUnique(XDocument document, params XName[] path) {
        XElement? current = document.Root;
        foreach (XName name in path) {
            if (current == null) return null;
            XElement[] matches = current.Elements(name).Take(2).ToArray();
            if (matches.Length > 1) throw new InvalidDataException("The CII field path is ambiguous at " + name.LocalName + ".");
            current = matches.FirstOrDefault();
        }
        return current;
    }

    private static XmlReaderSettings ReaderSettings() => new XmlReaderSettings {
        DtdProcessing = DtdProcessing.Prohibit,
        XmlResolver = null,
        MaxCharactersInDocument = MaximumXmlBytes
    };
}
