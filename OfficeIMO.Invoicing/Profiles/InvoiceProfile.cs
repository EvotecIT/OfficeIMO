namespace OfficeIMO.Invoicing;

/// <summary>Declared electronic invoice profiles. Recognition alone does not establish compliance.</summary>
public enum InvoiceProfile {
    /// <summary>Factur-X header subset.</summary>
    Minimum,
    /// <summary>Factur-X basic profile without invoice lines.</summary>
    BasicWithoutLines,
    /// <summary>Factur-X basic profile including core line information.</summary>
    Basic,
    /// <summary>The EN 16931 semantic invoice model.</summary>
    En16931,
    /// <summary>Factur-X extended profile.</summary>
    Extended,
    /// <summary>German XRechnung 3.0 usage specification.</summary>
    XRechnung,
    /// <summary>Peppol BIS Billing 3.0 usage specification.</summary>
    PeppolBis
}
