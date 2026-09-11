namespace OfficeIMO.Invoicing;

/// <summary>Payment instructions, credit transfer accounts, card or direct debit data.</summary>
public sealed class InvoicePayment {
    /// <summary>UNCL 4461 payment means code (BT-81), for example 58 for SEPA credit transfer.</summary>
    public string MeansCode { get; set; } = string.Empty;
    /// <summary>Payment means text (BT-82).</summary>
    public string? MeansText { get; set; }
    /// <summary>Remittance information (BT-83).</summary>
    public string? Reference { get; set; }
    /// <summary>Credit transfer accounts (BG-17).</summary>
    public IList<InvoiceBankAccount> Accounts { get; } = new List<InvoiceBankAccount>();
    /// <summary>Card primary account number; only the last four to six digits (BT-87).</summary>
    public string? CardNumber { get; set; }
    /// <summary>Card holder name (BT-88).</summary>
    public string? CardHolder { get; set; }
    /// <summary>Direct debit mandate reference (BT-89).</summary>
    public string? MandateReference { get; set; }
    /// <summary>Bank-assigned creditor identifier (BT-90).</summary>
    public string? CreditorIdentifier { get; set; }
    /// <summary>Debited account identifier (BT-91).</summary>
    public string? DebitedAccount { get; set; }
}

/// <summary>Payee credit transfer account.</summary>
public sealed class InvoiceBankAccount {
    /// <summary>IBAN or local account identifier (BT-84).</summary>
    public string Identifier { get; set; } = string.Empty;
    /// <summary>Whether CII should emit IBANID rather than a proprietary account identifier.</summary>
    public bool IsIban { get; set; } = true;
    /// <summary>Account name (BT-85).</summary>
    public string? Name { get; set; }
    /// <summary>Payment service provider identifier, such as a BIC (BT-86).</summary>
    public string? ProviderIdentifier { get; set; }
}
