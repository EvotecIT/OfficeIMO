namespace OfficeIMO.Pdf;

/// <summary>Exact-artifact page-box and font-embedding evidence for print-production preflight.</summary>
public sealed class PdfPrintProductionStructureEvidence {
    internal PdfPrintProductionStructureEvidence(
        int pageCount,
        int validProductionPageBoxCount,
        int invalidProductionPageBoxCount,
        int fontResourceCount,
        int unembeddedFontResourceCount,
        int uninspectableFontResourceCount) {
        PageCount = pageCount;
        ValidProductionPageBoxCount = validProductionPageBoxCount;
        InvalidProductionPageBoxCount = invalidProductionPageBoxCount;
        FontResourceCount = fontResourceCount;
        UnembeddedFontResourceCount = unembeddedFontResourceCount;
        UninspectableFontResourceCount = uninspectableFontResourceCount;
    }

    /// <summary>Number of pages inspected.</summary>
    public int PageCount { get; }

    /// <summary>Pages with a MediaBox, exactly one TrimBox or ArtBox, and any explicit BleedBox in valid nesting order.</summary>
    public int ValidProductionPageBoxCount { get; }

    /// <summary>Pages missing or violating the required production boundary-box relationship.</summary>
    public int InvalidProductionPageBoxCount { get; }

    /// <summary>Distinct font resource dictionaries discovered in the artifact object graph.</summary>
    public int FontResourceCount { get; }

    /// <summary>Font resources without a self-contained font program or Type3 character procedures.</summary>
    public int UnembeddedFontResourceCount { get; }

    /// <summary>Font resource dictionaries that could not be inspected completely.</summary>
    public int UninspectableFontResourceCount { get; }

    /// <summary>True when every page has valid print boxes and every font resource is inspectable and embedded.</summary>
    public bool IsComplete =>
        PageCount > 0 &&
        ValidProductionPageBoxCount == PageCount &&
        InvalidProductionPageBoxCount == 0 &&
        UnembeddedFontResourceCount == 0 &&
        UninspectableFontResourceCount == 0;
}
