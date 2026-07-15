namespace OfficeIMO.Pdf;

/// <summary>Controls how an imported PDF page Form XObject fits its target rectangle.</summary>
public enum PdfPageOverlayFit {
    /// <summary>Keep the source page at its visual size.</summary>
    None,
    /// <summary>Preserve aspect ratio and fit entirely inside the target rectangle.</summary>
    Contain,
    /// <summary>Preserve aspect ratio and cover the target rectangle, clipping overflow.</summary>
    Cover,
    /// <summary>Scale independently in both dimensions to fill the target rectangle.</summary>
    Stretch
}
