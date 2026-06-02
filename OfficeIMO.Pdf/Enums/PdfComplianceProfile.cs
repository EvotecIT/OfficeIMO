namespace OfficeIMO.Pdf;

/// <summary>
/// Requested generated-PDF compliance profile.
/// </summary>
public enum PdfComplianceProfile {
    /// <summary>No formal archival, accessibility, or e-invoice profile is requested.</summary>
    None = 0,
    /// <summary>PDF/A-2b archival profile.</summary>
    PdfA2B,
    /// <summary>PDF/A-2u archival profile with Unicode text mapping requirements.</summary>
    PdfA2U,
    /// <summary>PDF/A-2a archival profile with accessibility structure requirements.</summary>
    PdfA2A,
    /// <summary>PDF/A-3b archival profile.</summary>
    PdfA3B,
    /// <summary>PDF/A-3u archival profile with Unicode text mapping requirements.</summary>
    PdfA3U,
    /// <summary>PDF/A-3a archival profile with accessibility structure requirements.</summary>
    PdfA3A,
    /// <summary>PDF/UA-1 accessibility profile.</summary>
    PdfUa1,
    /// <summary>Factur-X e-invoice profile, built on PDF/A-3 plus embedded EN 16931 XML.</summary>
    FacturX,
    /// <summary>ZUGFeRD e-invoice profile, built on PDF/A-3 plus embedded EN 16931 XML.</summary>
    Zugferd
}
