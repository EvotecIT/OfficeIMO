namespace OfficeIMO.Html;

/// <summary>Observed geometry, diagnostics, and executed checks for one hash-bound gallery artifact.</summary>
public sealed class HtmlCapabilityGalleryArtifactEvidence {
    /// <summary>Creates an immutable snapshot of observations from the artifact's producing operation.</summary>
    public HtmlCapabilityGalleryArtifactEvidence(
        int pageCount,
        int? pageNumber,
        double? width,
        double? height,
        string dimensionUnit,
        IEnumerable<HtmlDiagnostic> diagnostics,
        IEnumerable<HtmlCapabilityGalleryCheck> checks) {
        if (pageCount < 1) throw new ArgumentOutOfRangeException(nameof(pageCount));
        if (pageNumber.HasValue && (pageNumber < 1 || pageNumber > pageCount)) throw new ArgumentOutOfRangeException(nameof(pageNumber));
        if (width.HasValue && (!(width > 0) || double.IsInfinity(width.Value))) throw new ArgumentOutOfRangeException(nameof(width));
        if (height.HasValue && (!(height > 0) || double.IsInfinity(height.Value))) throw new ArgumentOutOfRangeException(nameof(height));
        PageCount = pageCount;
        PageNumber = pageNumber;
        Width = width;
        Height = height;
        DimensionUnit = dimensionUnit ?? throw new ArgumentNullException(nameof(dimensionUnit));
        Diagnostics = (diagnostics ?? throw new ArgumentNullException(nameof(diagnostics))).ToList().AsReadOnly();
        Checks = (checks ?? throw new ArgumentNullException(nameof(checks))).ToList().AsReadOnly();
    }

    /// <summary>Total pages in the rendered source document.</summary>
    public int PageCount { get; }
    /// <summary>One-based source page, or null for a multi-page artifact.</summary>
    public int? PageNumber { get; }
    /// <summary>Observed width, or null when pages may have different dimensions.</summary>
    public double? Width { get; }
    /// <summary>Observed height, or null when pages may have different dimensions.</summary>
    public double? Height { get; }
    /// <summary>Unit used for width and height, such as px or pt.</summary>
    public string DimensionUnit { get; }
    /// <summary>Diagnostics collected by this artifact's layout and encoding operation.</summary>
    public IReadOnlyList<HtmlDiagnostic> Diagnostics { get; }
    /// <summary>Checks that actually ran; caller-declared expectations are recorded separately.</summary>
    public IReadOnlyList<HtmlCapabilityGalleryCheck> Checks { get; }
    /// <summary>Whether this artifact's producing operation reported fidelity loss.</summary>
    public bool HasLoss => Diagnostics.Any(diagnostic => diagnostic.LossKind != OfficeConversionLossKind.None);
}

/// <summary>Result of an executed check, rather than an unexecuted capability expectation.</summary>
public sealed class HtmlCapabilityGalleryCheck {
    /// <summary>Creates a named check with its observed result and explanation.</summary>
    public HtmlCapabilityGalleryCheck(string name, bool passed, string detail) {
        Name = name ?? throw new ArgumentNullException(nameof(name));
        Passed = passed;
        Detail = detail ?? throw new ArgumentNullException(nameof(detail));
    }
    /// <summary>Stable check name.</summary>
    public string Name { get; }
    /// <summary>Whether the executed check passed.</summary>
    public bool Passed { get; }
    /// <summary>What was checked and any failure detail.</summary>
    public string Detail { get; }
}
