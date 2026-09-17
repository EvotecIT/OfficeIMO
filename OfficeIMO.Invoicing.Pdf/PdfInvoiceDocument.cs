using OfficeIMO.Pdf;

namespace OfficeIMO.Invoicing.Pdf;

/// <summary>Captures a typed invoice once for visible PDF content and its embedded CII attachment.</summary>
public sealed partial class PdfInvoiceDocument {
    private readonly byte[] _xml;
    private readonly Invoice _invoice;
    private readonly PdfInvoiceAmounts _amounts;
    private readonly InvoicePdfLayoutOptions _layout;
    private readonly DateTimeOffset _capturedAt = DateTimeOffset.UtcNow;
    private string DocumentTitle => Label(_invoice.TypeCode == "381" ? InvoicePdfText.CreditNote : InvoicePdfText.Invoice) + " " + _invoice.Number;

    private PdfInvoiceDocument(byte[] xml, InvoiceXmlOptions contract, InvoicePdfLayoutOptions layout) {
        _xml = xml;
        _layout = layout.Snapshot();
        InvoiceReadResult parsed = InvoiceParser.Read(xml);
        if (!parsed.HasCompleteMapping) throw new InvalidDataException("Generated invoice cannot be represented completely for PDF presentation.");
        Profile = parsed.Declaration.Profile ?? throw new InvalidDataException("Generated invoice has no recognized profile.");
        Release = contract.Release;
        _invoice = parsed.Invoice;
        if (_invoice.TypeCode != "380" && _invoice.TypeCode != "381")
            throw new NotSupportedException("PDF invoice presentation supports document type 380 (invoice) and 381 (credit note) only.");
        _amounts = PdfInvoiceAmounts.Create(_invoice, Profile);
    }

    /// <summary>Creates an independent snapshot for type 380 (invoice) or 381 (credit note). Later edits to the supplied invoice cannot change its PDF or XML.</summary>
    public static PdfInvoiceDocument Create(Invoice invoice, InvoiceXmlOptions contract, InvoicePdfLayoutOptions? layout = null) {
#if NET8_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(contract);
#else
        if (contract == null) throw new ArgumentNullException(nameof(contract));
#endif
        if (contract.Syntax != InvoiceSyntax.Cii)
            throw new NotSupportedException("Factur-X PDF authoring requires a CII output contract.");
        return new PdfInvoiceDocument(InvoiceSerializer.Write(invoice, contract), contract, layout ?? new InvoicePdfLayoutOptions());
    }

    /// <summary>Captured invoice number.</summary>
    public string Number => _invoice.Number;
    /// <summary>Captured CII authoring profile. Pass this value to <see cref="Create"/> when recapturing an edited model.</summary>
    public InvoiceProfile Profile { get; }
    /// <summary>Captured specification release. Preserve it when recapturing an edited model.</summary>
    public InvoiceSpecificationRelease Release { get; }
    /// <summary>Captured amount due in the document currency.</summary>
    public decimal PayableAmount => _amounts.PayableAmount;
    /// <summary>Returns a defensive copy of the exact CII bytes embedded in generated PDFs.</summary>
    public byte[] ToXmlBytes() => (byte[])_xml.Clone();
    /// <summary>Returns an independent editable model. After edits, call <see cref="Create"/> with an explicit contract preserving this snapshot's release and profile.</summary>
    public Invoice ToInvoice() => InvoiceParser.Read(_xml).Invoice;

    /// <summary>
    /// Renders the captured invoice and attaches its exact CII XML using Factur-X PDF/A-3 groundwork.
    /// Supply embedded fonts through the PDF options. Validate the exact output with PDF/A and invoice validators before claiming compliance.
    /// </summary>
    public byte[] ToPdfBytes(PdfOptions? options = null) {
        return RenderPdf(options, embedInvoiceXml: true);
    }

    /// <summary>
    /// Renders the captured invoice as a presentation-only PDF without attaching CII XML.
    /// Use <see cref="ToPdfBytes(PdfOptions?)"/> when a Factur-X/ZUGFeRD hybrid document is required.
    /// </summary>
    public byte[] ToPresentationPdfBytes(PdfOptions? options = null) {
        return RenderPdf(options, embedInvoiceXml: false);
    }

    private byte[] RenderPdf(PdfOptions? options, bool embedInvoiceXml) {
        PdfOptions configured = options?.Clone() ?? new PdfOptions();
        configured.UseTextFallbacks(PdfTextFallbackFeatures.MultilingualFonts);
        if (configured.TextShapingProvider == null) configured.UseManagedTextShaping();
        if (embedInvoiceXml) {
            configured.UseFacturX(_xml, relationship: PdfAssociatedFileRelationship.Alternative);
            configured.SetEmbeddedFileModificationDate("factur-x.xml", _capturedAt);
        }
        PdfDocument document = PdfDocument.Create(configured);
        document.Meta(title: DocumentTitle, author: _invoice.Seller.Name);
        Compose(document.Content);
        return document.ToBytes();
    }
}
