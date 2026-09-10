using System.Xml.Linq;

namespace OfficeIMO.Invoicing;

/// <summary>The namespace-aware declaration in an invoice. This is identification, not schema or business-rule validation.</summary>
public sealed class InvoiceProfileDeclaration {
    /// <summary>Maximum input size accepted by the invoice XML reader.</summary>
    public const int MaximumXmlBytes = 16 * 1024 * 1024;

    private InvoiceProfileDeclaration(InvoiceSyntax syntax, string guidelineId, InvoiceProfile? profile) {
        Syntax = syntax;
        GuidelineId = guidelineId;
        Profile = profile;
    }

    /// <summary>Identified XML syntax.</summary>
    public InvoiceSyntax Syntax { get; }
    /// <summary>Exact trimmed guideline/customization identifier.</summary>
    public string GuidelineId { get; }
    /// <summary>Recognized profile, or null for an unsupported identifier.</summary>
    public InvoiceProfile? Profile { get; }

    /// <summary>Reads one unambiguous profile declaration, rejecting malformed, oversized, DTD-bearing, and ambiguous XML.</summary>
    public static InvoiceProfileDeclaration Read(byte[] xml) => Read(InvoiceXml.Parse(xml));

    internal static InvoiceProfileDeclaration Read(XDocument document) {
        XElement root = document.Root ?? throw new InvalidDataException("Invoice XML has no document element.");
        InvoiceSyntax syntax;
        XElement? identifier;
        if (root.Name == InvoiceXml.Rsm + "CrossIndustryInvoice") {
            syntax = InvoiceSyntax.Cii;
            XElement? context = InvoiceXml.Unique(root, InvoiceXml.Rsm + "ExchangedDocumentContext");
            XElement? guideline = context == null ? null : InvoiceXml.Unique(context, InvoiceXml.Ram + "GuidelineSpecifiedDocumentContextParameter");
            identifier = guideline == null ? null : InvoiceXml.Unique(guideline, InvoiceXml.Ram + "ID");
        } else if (root.Name == InvoiceXml.UblInvoice + "Invoice" || root.Name == InvoiceXml.UblCreditNote + "CreditNote") {
            syntax = InvoiceSyntax.Ubl;
            identifier = InvoiceXml.Unique(root, InvoiceXml.Cbc + "CustomizationID");
        } else {
            throw new InvalidDataException("Expected CII namespace-100 CrossIndustryInvoice or UBL Invoice/CreditNote XML.");
        }
        string value = identifier == null ? string.Empty : InvoiceXml.Scalar(identifier);
        if (value.Length == 0) throw new InvalidDataException("Invoice XML requires a non-empty guideline/customization identifier.");
        InvoiceProfile? profile = InvoiceProfiles.TryFromGuidelineId(value, out InvoiceProfile known) ? known : null;
        return new InvoiceProfileDeclaration(syntax, value, profile);
    }
}
