namespace OfficeIMO.Invoicing;

/// <summary>Explicit syntax and guideline for deterministic XML emission. Validation release selection is separate.</summary>
public sealed class InvoiceXmlOptions {
    /// <summary>Creates an output contract; defaults to CII with the EN 16931 guideline.</summary>
    public InvoiceXmlOptions(InvoiceSyntax syntax = InvoiceSyntax.Cii, InvoiceProfile profile = InvoiceProfile.En16931) {
        if (syntax != InvoiceSyntax.Cii && syntax != InvoiceSyntax.Ubl) throw new ArgumentOutOfRangeException(nameof(syntax));
        if (profile != InvoiceProfile.En16931 && profile != InvoiceProfile.XRechnung && profile != InvoiceProfile.PeppolBis)
            throw new NotSupportedException("Typed authoring supports EN 16931, XRechnung and Peppol BIS. Other Factur-X levels are available for declaration inspection only.");
        if (profile == InvoiceProfile.PeppolBis && syntax != InvoiceSyntax.Ubl)
            throw new NotSupportedException("Peppol BIS authoring requires UBL.");
        Syntax = syntax; Profile = profile;
    }
    /// <summary>Output syntax.</summary>
    public InvoiceSyntax Syntax { get; }
    /// <summary>Output guideline.</summary>
    public InvoiceProfile Profile { get; }
}
