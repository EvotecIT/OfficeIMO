namespace OfficeIMO.Invoicing;

/// <summary>Editable semantic invoice shared by XML writers and presentation adapters.</summary>
public sealed class Invoice {
    /// <summary>Seller-assigned invoice number (BT-1).</summary>
    public string Number { get; set; } = string.Empty;
    /// <summary>Issue date, without a time or timezone (BT-2).</summary>
    public DateTime IssueDate { get; set; }
    /// <summary>UNCL 1001 document type, normally 380 for an invoice or 381 for a credit note.</summary>
    public string TypeCode { get; set; } = "380";
    /// <summary>ISO 4217 invoice currency (BT-5).</summary>
    public string Currency { get; set; } = "EUR";
    /// <summary>Business process identifier (BT-23), distinct from the selected guideline.</summary>
    public string? BusinessProcessId { get; set; }
    /// <summary>Buyer routing reference, including a Leitweg-ID where required (BT-10).</summary>
    public string? BuyerReference { get; set; }
    /// <summary>Seller (BG-4).</summary>
    public InvoiceParty Seller { get; set; } = new InvoiceParty();
    /// <summary>Buyer (BG-7).</summary>
    public InvoiceParty Buyer { get; set; } = new InvoiceParty();
    /// <summary>Payee when different from the seller (BG-10).</summary>
    public InvoiceParty? Payee { get; set; }
    /// <summary>Seller's tax representative (BG-11).</summary>
    public InvoiceParty? TaxRepresentative { get; set; }
    /// <summary>Delivery recipient, address, location identifier and date (BG-13).</summary>
    public InvoiceDelivery? Delivery { get; set; }
    /// <summary>Invoice period (BG-14).</summary>
    public InvoicePeriod? Period { get; set; }
    /// <summary>Tax point date (BT-7).</summary>
    public DateTime? TaxPointDate { get; set; }
    /// <summary>Tax point date code when the actual date is not known (BT-8).</summary>
    public string? TaxPointDateCode { get; set; }
    /// <summary>Project reference (BT-11).</summary>
    public string? ProjectReference { get; set; }
    /// <summary>Contract reference (BT-12).</summary>
    public string? ContractReference { get; set; }
    /// <summary>Purchase order reference (BT-13).</summary>
    public string? PurchaseOrderReference { get; set; }
    /// <summary>Sales order reference (BT-14).</summary>
    public string? SalesOrderReference { get; set; }
    /// <summary>Receiving advice reference (BT-15).</summary>
    public string? ReceivingAdviceReference { get; set; }
    /// <summary>Despatch advice reference (BT-16).</summary>
    public string? DespatchAdviceReference { get; set; }
    /// <summary>Tender or lot reference (BT-17).</summary>
    public string? TenderReference { get; set; }
    /// <summary>Buyer accounting reference (BT-19).</summary>
    public string? AccountingReference { get; set; }
    /// <summary>Invoiced object identifier (BT-18).</summary>
    public InvoiceIdentifier? ObjectIdentifier { get; set; }
    /// <summary>Free text notes, optionally classified by subject (BG-1).</summary>
    public IList<InvoiceNote> Notes { get; } = new List<InvoiceNote>();
    /// <summary>Preceding invoice references (BG-3).</summary>
    public IList<InvoiceReference> PrecedingInvoices { get; } = new List<InvoiceReference>();
    /// <summary>Additional supporting documents (BG-24).</summary>
    public IList<InvoiceSupportingDocument> SupportingDocuments { get; } = new List<InvoiceSupportingDocument>();
    /// <summary>Invoice lines (BG-25).</summary>
    public IList<InvoiceLine> Lines { get; } = new List<InvoiceLine>();
    /// <summary>Document-level allowances and charges (BG-20/BG-21).</summary>
    public IList<InvoiceAllowanceCharge> AllowancesAndCharges { get; } = new List<InvoiceAllowanceCharge>();
    /// <summary>Payment instructions (BG-16).</summary>
    public InvoicePayment? Payment { get; set; }
    /// <summary>Payment terms (BT-20).</summary>
    public string? PaymentTerms { get; set; }
    /// <summary>Payment due date (BT-9).</summary>
    public DateTime? DueDate { get; set; }
    /// <summary>Amount already paid (BT-113).</summary>
    public decimal PrepaidAmount { get; set; }
    /// <summary>Explicit rounding adjustment to the amount due (BT-114).</summary>
    public decimal RoundingAmount { get; set; }
    /// <summary>Accounting currency for VAT reporting (BT-6).</summary>
    public string? TaxCurrency { get; set; }
    /// <summary>VAT total in the accounting currency (BT-111), supplied with an explicit tax currency.</summary>
    public decimal? TaxAmountInAccountingCurrency { get; set; }
    /// <summary>Totals declared by a source document, retained for validation rather than silently recalculated on import.</summary>
    public InvoiceDeclaredTotals? DeclaredTotals { get; set; }
    /// <summary>VAT breakdowns declared by a source document (BG-23).</summary>
    public IList<InvoiceDeclaredTax> DeclaredTaxes { get; } = new List<InvoiceDeclaredTax>();
}
