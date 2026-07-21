namespace OfficeIMO.Pdf;

/// <summary>Renderer compatibility diagnostics for one source page without producing image bytes.</summary>
public sealed class PdfRenderCompatibilityPage {
    internal PdfRenderCompatibilityPage(int pageNumber, IReadOnlyList<PdfRenderCapabilityDiagnostic> diagnostics) {
        PageNumber = pageNumber;
        Diagnostics = diagnostics;
    }

    /// <summary>One-based page number.</summary>
    public int PageNumber { get; }

    /// <summary>Simplified or unsupported features discovered on this page.</summary>
    public IReadOnlyList<PdfRenderCapabilityDiagnostic> Diagnostics { get; }

    /// <summary>True when no known simplification or omission was detected.</summary>
    public bool IsExactForManagedRenderer => Diagnostics.Count == 0;

    /// <summary>True when at least one feature would be omitted.</summary>
    public bool HasUnsupportedFeatures => Diagnostics.Any(static diagnostic => diagnostic.SupportLevel == PdfRenderSupportLevel.Unsupported);
}

/// <summary>
/// Document-level producer interoperability assessment backed by the generated render capability registry.
/// </summary>
public sealed class PdfRenderCompatibilityReport {
    private readonly PdfRenderCapabilityManifest _manifest;

    internal PdfRenderCompatibilityReport(IReadOnlyList<PdfRenderCompatibilityPage> pages) {
        Pages = pages;
        _manifest = PdfRenderCapabilities.Current;
    }

    /// <summary>The single generated capability manifest used to interpret every diagnostic.</summary>
    public PdfRenderCapabilityManifest Manifest => _manifest;

    /// <summary>Per-page compatibility evidence.</summary>
    public IReadOnlyList<PdfRenderCompatibilityPage> Pages { get; }

    /// <summary>Total simplified and unsupported occurrences.</summary>
    public int DiagnosticCount => Pages.Sum(static page => page.Diagnostics.Count);

    /// <summary>True when at least one feature would be approximated.</summary>
    public bool HasSimplifications => Pages.Any(static page => page.Diagnostics.Any(
        static diagnostic => diagnostic.SupportLevel == PdfRenderSupportLevel.Simplified));

    /// <summary>True when at least one feature would be omitted.</summary>
    public bool HasUnsupportedFeatures => Pages.Any(static page => page.HasUnsupportedFeatures);

    /// <summary>True when the assessment found no known renderer approximation or omission.</summary>
    public bool IsExactForManagedRenderer => DiagnosticCount == 0;

    /// <summary>Stable distinct capability identifiers encountered across the document.</summary>
    public IReadOnlyList<string> CapabilityCodes => Pages
        .SelectMany(static page => page.Diagnostics)
        .Select(static diagnostic => diagnostic.Code)
        .Distinct(StringComparer.Ordinal)
        .OrderBy(static code => code, StringComparer.Ordinal)
        .ToArray();
}
