namespace OfficeIMO.Invoicing.Pdf;

/// <summary>Identifies a generated label in the visible invoice layout.</summary>
public enum InvoicePdfText {
    /// <summary>Invoice document heading.</summary>
    Invoice,
    /// <summary>Credit-note document heading.</summary>
    CreditNote,
    /// <summary>Seller party heading.</summary>
    Seller,
    /// <summary>Buyer party heading.</summary>
    Buyer,
    /// <summary>Invoice line item heading.</summary>
    Item,
    /// <summary>Invoiced quantity label.</summary>
    Quantity,
    /// <summary>Net unit price label.</summary>
    NetPrice,
    /// <summary>Value-added tax label.</summary>
    Vat,
    /// <summary>Line net amount label.</summary>
    NetAmount,
    /// <summary>VAT category label.</summary>
    VatCategory,
    /// <summary>Taxable amount label.</summary>
    TaxableAmount,
    /// <summary>VAT amount label.</summary>
    VatAmount,
    /// <summary>Tax-exemption label.</summary>
    Exemption,
    /// <summary>Issue-date label.</summary>
    Issued,
    /// <summary>Payment due-date label.</summary>
    Due,
    /// <summary>Buyer reference label.</summary>
    BuyerReference,
    /// <summary>Invoice currency label.</summary>
    Currency,
    /// <summary>Document-level charges and allowances heading.</summary>
    DocumentAdjustments,
    /// <summary>Charge label.</summary>
    Charge,
    /// <summary>Allowance label.</summary>
    Allowance,
    /// <summary>Sum of invoice-line net amounts label.</summary>
    LineNetTotal,
    /// <summary>Total allowances label.</summary>
    Allowances,
    /// <summary>Total charges label.</summary>
    Charges,
    /// <summary>Total excluding VAT label.</summary>
    TotalExcludingVat,
    /// <summary>Total VAT label.</summary>
    VatTotal,
    /// <summary>Total including VAT label.</summary>
    TotalIncludingVat,
    /// <summary>Prepaid amount label.</summary>
    Prepaid,
    /// <summary>Rounding amount label.</summary>
    Rounding,
    /// <summary>Amount due for payment label.</summary>
    AmountDue,
    /// <summary>Start of a range or period label.</summary>
    From,
    /// <summary>End of a bounded period label.</summary>
    Until,
    /// <summary>Destination or end label.</summary>
    To,
    /// <summary>Generic identifier label.</summary>
    Identifier,
    /// <summary>Legal registration identifier label.</summary>
    LegalRegistration,
    /// <summary>Electronic address label.</summary>
    ElectronicAddress,
    /// <summary>Referenced order-line label.</summary>
    OrderLine,
    /// <summary>Buyer accounting reference label.</summary>
    AccountingReference,
    /// <summary>Invoiced-object identifier label.</summary>
    ObjectIdentifier,
    /// <summary>Standard item identifier label.</summary>
    StandardItem,
    /// <summary>Seller-assigned item identifier label.</summary>
    SellerItem,
    /// <summary>Buyer-assigned item identifier label.</summary>
    BuyerItem,
    /// <summary>Country-of-origin label.</summary>
    Origin,
    /// <summary>Item classification label.</summary>
    Classification,
    /// <summary>Version label.</summary>
    Version,
    /// <summary>Invoice or line period label.</summary>
    Period,
    /// <summary>Gross unit price label.</summary>
    GrossPrice,
    /// <summary>Unit-price discount label.</summary>
    PriceDiscount,
    /// <summary>Charge or allowance reason-code label.</summary>
    ReasonCode,
    /// <summary>Charge or allowance base amount label.</summary>
    Base,
    /// <summary>Charge or allowance percentage label.</summary>
    Percentage,
    /// <summary>Purchase-order reference label.</summary>
    PurchaseOrder,
    /// <summary>Sales-order reference label.</summary>
    SalesOrder,
    /// <summary>Contract reference label.</summary>
    Contract,
    /// <summary>Project reference label.</summary>
    Project,
    /// <summary>Despatch-advice reference label.</summary>
    DespatchAdvice,
    /// <summary>Receiving-advice reference label.</summary>
    ReceivingAdvice,
    /// <summary>Tender or lot reference label.</summary>
    Tender,
    /// <summary>Invoiced-object reference label.</summary>
    InvoicedObject,
    /// <summary>Business-process identifier label.</summary>
    BusinessProcess,
    /// <summary>Invoice document-type label.</summary>
    DocumentType,
    /// <summary>VAT tax-point date label.</summary>
    TaxPoint,
    /// <summary>Preceding-invoice reference label.</summary>
    PrecedingInvoice,
    /// <summary>VAT accounting currency label.</summary>
    VatAccountingCurrency,
    /// <summary>Document references heading.</summary>
    References,
    /// <summary>Delivery information heading.</summary>
    Delivery,
    /// <summary>Delivery location label.</summary>
    Location,
    /// <summary>Payee party heading.</summary>
    Payee,
    /// <summary>Seller tax representative heading.</summary>
    TaxRepresentative,
    /// <summary>Payment instructions heading.</summary>
    Payment,
    /// <summary>Payment-means label.</summary>
    PaymentMeans,
    /// <summary>Payment reference label.</summary>
    Reference,
    /// <summary>Payment account label.</summary>
    Account,
    /// <summary>Direct-debit mandate label.</summary>
    Mandate,
    /// <summary>Creditor identifier label.</summary>
    Creditor,
    /// <summary>Debited account label.</summary>
    DebitedAccount,
    /// <summary>Payment-card label.</summary>
    Card,
    /// <summary>Masked card-number ending label.</summary>
    Ending,
    /// <summary>Invoice notes heading.</summary>
    Notes,
    /// <summary>Supporting documents heading.</summary>
    SupportingDocuments
}
