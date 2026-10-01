namespace OfficeIMO.Invoicing;

/// <summary>Immutable CII/UBL source XML for bounded edits that retain unmapped extensions. Loading does not establish standards or business-rule compliance.</summary>
public sealed class InvoiceSourceDocument {
    private readonly InvoicePreservedXml _source;
    internal InvoiceSourceDocument(InvoicePreservedXml source) => _source = source;
    /// <summary>Maximum accepted and edited XML bytes, including a byte order mark.</summary>
    public const int MaximumXmlBytes = InvoiceProfileDeclaration.MaximumXmlBytes;
    /// <summary>Maximum XML nodes accepted before tree materialization.</summary>
    public const int MaximumXmlNodes = InvoiceXml.MaximumNodes;
    /// <summary>Maximum XML attributes accepted before tree materialization.</summary>
    public const int MaximumXmlAttributes = InvoiceXml.MaximumAttributes;
    /// <summary>Source syntax. UBL includes Invoice and CreditNote roots.</summary>
    public InvoiceSyntax Syntax => _source.IsCii ? InvoiceSyntax.Cii : InvoiceSyntax.Ubl;
    /// <summary>Whether the source root is a UBL credit note.</summary>
    public bool IsUblCreditNote => _source.IsCreditNote;
    /// <summary>Whether XML Signature content is present. Signatures are not verified.</summary>
    public bool HasXmlSignature => _source.HasXmlSignature;
    /// <summary>Existing invoice number, or null when missing. Ambiguous or nested fields are rejected.</summary>
    public string? Number => _source.Read(InvoiceSourceField.DocumentId);
    /// <summary>Declared document type, without establishing compliance.</summary>
    public string? TypeCode => _source.Read(InvoiceSourceField.TypeCode);
    /// <summary>Declared profile guideline/customization identifier.</summary>
    public string? GuidelineId => _source.Read(InvoiceSourceField.GuidelineId);
    /// <summary>Declared invoice currency, without validating its monetary amounts.</summary>
    public string? Currency => _source.Read(InvoiceSourceField.CurrencyCode);
    /// <summary>Issue date for a format-102 CII date or plain ISO UBL date; null for other representations.</summary>
    public DateTime? IssueDate => _source.ReadDate(InvoiceSourceField.IssueDate);
    /// <summary>Existing payment due date in the supported date representation. UBL credit notes have no supported due-date edit contract.</summary>
    public DateTime? DueDate => IsUblCreditNote ? null : _source.ReadDate(InvoiceSourceField.DueDate);
    /// <summary>Existing buyer reference.</summary>
    public string? BuyerReference => _source.Read(InvoiceSourceField.BuyerReference);
    /// <summary>Existing remittance reference. Multiple UBL payment-means occurrences are ambiguous and rejected.</summary>
    public string? PaymentReference => _source.Read(InvoiceSourceField.PaymentReference);

    /// <summary>Captures a defensive copy of supported source XML with DTDs/external entities disabled and size, depth, node and attribute limits.</summary>
    public static InvoiceSourceDocument Load(byte[] xml) => new(InvoicePreservedXml.Load(xml));
    /// <summary>Reads from the current stream position, leaves the stream open, and consumes at most the byte limit plus one byte.</summary>
    public static InvoiceSourceDocument Load(Stream stream) => new(InvoicePreservedXml.Load(stream));
    /// <summary>Loads a file under the same source limits.</summary>
    public static InvoiceSourceDocument Load(string path) { using var stream = File.OpenRead(path); return Load(stream); }
    /// <summary>Returns a defensive copy. Unedited documents are byte-identical; edited documents retain XML content using UTF-8, without preserving original lexical formatting.</summary>
    public byte[] ToBytes() => _source.ToBytes();
    internal InvoiceSourceDocument Apply(IReadOnlyList<InvoiceSourceScalarEdit> edits) => new(_source.Apply(edits));
    internal string FormatDate(DateTime date) => date.ToString(_source.IsCii ? "yyyyMMdd" : "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture);
}
