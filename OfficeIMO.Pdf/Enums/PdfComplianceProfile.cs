namespace OfficeIMO.Pdf;

/// <summary>
/// Requested generated-PDF compliance profile.
/// </summary>
public enum PdfComplianceProfile {
    /// <summary>No formal archival, accessibility, or e-invoice profile is requested.</summary>
    None = 0,
    /// <summary>PDF/A-2b archival profile.</summary>
    PdfA2B = 1,
    /// <summary>PDF/A-2u archival profile with Unicode text mapping requirements.</summary>
    PdfA2U = 2,
    /// <summary>PDF/A-2a archival profile with accessibility structure requirements.</summary>
    PdfA2A = 3,
    /// <summary>PDF/A-3b archival profile.</summary>
    PdfA3B = 4,
    /// <summary>PDF/A-3u archival profile with Unicode text mapping requirements.</summary>
    PdfA3U = 5,
    /// <summary>PDF/A-3a archival profile with accessibility structure requirements.</summary>
    PdfA3A = 6,
    /// <summary>PDF/UA-1 accessibility profile.</summary>
    PdfUa1 = 7,
    /// <summary>Factur-X e-invoice profile, built on PDF/A-3 plus embedded EN 16931 XML.</summary>
    FacturX = 8,
    /// <summary>ZUGFeRD e-invoice profile, built on PDF/A-3 plus embedded EN 16931 XML.</summary>
    Zugferd = 9,
    /// <summary>PDF/A-4 archival profile.</summary>
    PdfA4 = 10,
    /// <summary>PDF/A-4e archival profile for engineering documents.</summary>
    PdfA4E = 11,
    /// <summary>PDF/A-4f archival profile with embedded-file support.</summary>
    PdfA4F = 12,
    /// <summary>PDF/UA-2 accessibility profile.</summary>
    PdfUa2 = 13
}
