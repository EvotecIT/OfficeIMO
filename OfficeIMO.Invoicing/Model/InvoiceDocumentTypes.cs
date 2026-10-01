namespace OfficeIMO.Invoicing;

/// <summary>Document type codes with consistent CII, UBL and PDF model mappings. Each output still requires its explicit release/profile checks and standards validation.</summary>
public static class InvoiceDocumentTypes {
    /// <summary>Commercial invoice.</summary>
    public const string Invoice = "380";
    /// <summary>Credit note.</summary>
    public const string CreditNote = "381";
    /// <summary>Partial invoice.</summary>
    public const string PartialInvoice = "326";
    /// <summary>Invoice correcting an earlier invoice.</summary>
    public const string CorrectedInvoice = "384";
    /// <summary>Invoice for payment before delivery or service.</summary>
    public const string PrepaymentInvoice = "386";
    /// <summary>Invoice issued by the buyer on the seller's behalf.</summary>
    public const string SelfBilledInvoice = "389";
    /// <summary>Whether a code has a supported UBL root mapping and visible invoice heading.</summary>
    public static bool HasPresentationMapping(string? code) => code is Invoice or CreditNote or PartialInvoice or CorrectedInvoice or PrepaymentInvoice or SelfBilledInvoice;
}
