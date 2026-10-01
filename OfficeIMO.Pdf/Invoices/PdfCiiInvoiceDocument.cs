using System.Globalization;
using System.Xml;
using OfficeIMO.Internal.Invoicing;

namespace OfficeIMO.Pdf;

/// <summary>An immutable CII namespace-100 XML document with bounded header editing. Loading does not establish schema, business-rule, tax, or profile compliance.</summary>
public sealed partial class PdfCiiInvoiceDocument {
    private readonly InvoicePreservedXml _source;
    private PdfCiiInvoiceDocument(InvoicePreservedXml source) {
        if (!source.IsCii) throw new InvalidDataException("Expected a namespace-100 UN/CEFACT CrossIndustryInvoice root.");
        _source = source;
    }
    /// <summary>Maximum accepted XML byte length, including any byte order mark.</summary>
    public const int MaximumXmlBytes = InvoiceProfileDeclaration.MaximumXmlBytes;
    /// <summary>Maximum XML nodes retained by the invoice tree.</summary>
    public const int MaximumXmlNodes = InvoiceXml.MaximumNodes;
    /// <summary>Maximum XML attributes retained by the invoice tree.</summary>
    public const int MaximumXmlAttributes = InvoiceXml.MaximumAttributes;
    /// <summary>Invoice identifier from the direct exchanged-document header, if present.</summary>
    public string? DocumentId => _source.Read(InvoiceSourceField.DocumentId);
    /// <summary>Declared document type code, without interpreting its profile validity.</summary>
    public string? TypeCode => _source.Read(InvoiceSourceField.TypeCode);
    /// <summary>Declared guideline identifier, without claiming that its requirements are satisfied.</summary>
    public string? GuidelineId => _source.Read(InvoiceSourceField.GuidelineId);
    /// <summary>Declared invoice currency, without validating code lists or monetary amounts.</summary>
    public string? CurrencyCode => _source.Read(InvoiceSourceField.CurrencyCode);
    /// <summary>Issue date for a valid format-102 calendar date; null for missing or other date representations.</summary>
    public DateTime? IssueDate => _source.ReadDate(InvoiceSourceField.IssueDate);
    /// <summary>Whether an XML Signature element is present. Signatures are not verified by this model.</summary>
    public bool HasXmlSignature => _source.HasXmlSignature;
    /// <summary>Loads a defensive copy of CII XML. DTDs, external entities and excessive size, depth, nodes and attributes are rejected. Unknown content is retained; namespace prefixes may vary.</summary>
    public static PdfCiiInvoiceDocument Load(byte[] xml) => new(InvoicePreservedXml.Load(xml));
    /// <summary>Returns stored XML bytes as a defensive copy. Unedited documents retain byte identity; edited documents use UTF-8 and preserve XML content rather than original lexical formatting.</summary>
    public byte[] ToBytes() => _source.ToBytes();
    /// <summary>Replaces one existing, unambiguous plaintext invoice identifier. Other fields and visible PDF content are retained. Signed, structured or missing fields are rejected.</summary>
    public PdfCiiInvoiceDocument WithDocumentId(string documentId) {
        if (string.IsNullOrWhiteSpace(documentId)) throw new ArgumentException("An invoice identifier is required.", nameof(documentId));
        if (documentId.Length > MaximumXmlBytes) throw new ArgumentException("The invoice identifier exceeds the XML size limit.", nameof(documentId));
        XmlConvert.VerifyXmlChars(documentId);
        return Apply(InvoiceSourceField.DocumentId, documentId);
    }
    /// <summary>Replaces one existing format-102 issue date. Time/timezone and other dates are retained. Signed XML, ambiguous fields and other date representations are rejected.</summary>
    public PdfCiiInvoiceDocument WithIssueDate(DateTime issueDate) => Apply(InvoiceSourceField.IssueDate, issueDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture));
    private PdfCiiInvoiceDocument Apply(InvoiceSourceField field, string value) {
        try { return new(_source.Apply(new[] { new InvoiceSourceScalarEdit(field, value) })); }
        catch (InvoiceSourceFieldException exception) {
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(exception.InnerException!).Throw();
            throw;
        }
    }
}
