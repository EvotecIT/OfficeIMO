using OfficeIMO.Invoicing;

namespace OfficeIMO.Pdf;

/// <summary>Captures a typed invoice once for visible PDF content and its embedded CII attachment.</summary>
public sealed partial class PdfInvoiceDocument {
    private readonly byte[] _xml;
    private readonly Invoice _invoice;
    private readonly InvoiceCalculation _amounts;

    private PdfInvoiceDocument(byte[] xml) {
        _xml = xml;
        InvoiceReadResult parsed = InvoiceParser.Read(xml);
        if (!parsed.HasCompleteMapping) throw new InvalidDataException("Generated invoice cannot be represented completely for PDF presentation.");
        _invoice = parsed.Invoice;
        InvoiceModelValidationResult validation = InvoiceModelValidator.Validate(_invoice);
        validation.ThrowIfInvalid();
        _amounts = validation.Calculation!;
    }

    /// <summary>Creates an independent snapshot. Later edits to the supplied invoice cannot change its PDF or XML.</summary>
    public static PdfInvoiceDocument Create(Invoice invoice, InvoiceProfile profile = InvoiceProfile.En16931) =>
        new PdfInvoiceDocument(InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(InvoiceSyntax.Cii, profile)));

    /// <summary>Captured invoice number.</summary>
    public string Number => _invoice.Number;
    /// <summary>Captured amount due in the document currency.</summary>
    public decimal PayableAmount => _amounts.PayableAmount;
    /// <summary>Returns a defensive copy of the exact CII bytes embedded in generated PDFs.</summary>
    public byte[] ToXmlBytes() => (byte[])_xml.Clone();
    /// <summary>Returns an independent editable model. Create a new snapshot after edits.</summary>
    public Invoice ToInvoice() => InvoiceParser.Read(_xml).Invoice;

    /// <summary>
    /// Renders the captured invoice and attaches its exact CII XML using Factur-X PDF/A-3 groundwork.
    /// Supply embedded fonts through the PDF options. Validate the exact output with PDF/A and invoice validators before claiming compliance.
    /// </summary>
    public byte[] ToPdfBytes(PdfOptions? options = null) {
        PdfOptions configured = options?.Clone() ?? new PdfOptions();
        configured.UseFacturX(_xml, relationship: PdfAssociatedFileRelationship.Alternative);
        PdfDocument document = PdfDocument.Create(configured);
        document.Meta(title: (_invoice.TypeCode == "381" ? "Credit note " : "Invoice ") + _invoice.Number, author: _invoice.Seller.Name);
        Compose(document.Content);
        return document.ToBytes();
    }
}
