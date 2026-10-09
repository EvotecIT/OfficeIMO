using OfficeIMO.Drawing;

namespace OfficeIMO.DjVu;

/// <summary>Bounded page rendering settings, captured at the start of an operation.</summary>
public sealed class DjVuRenderOptions {
    /// <summary>Output resolution. Null uses the page's declared DPI.</summary>
    public int? Dpi { get; set; }
    /// <summary>Native, unrotated page region with a bottom-left origin. Null renders the whole page.</summary>
    public DjVuRectangle? Region { get; set; }
    /// <summary>Apply the page's declared display rotation.</summary>
    public bool ApplyRotation { get; set; } = true;
    /// <summary>Output device gamma; the conventional display value is 2.2.</summary>
    public double Gamma { get; set; } = 2.2;
    /// <summary>Background for a bitonal page without a background layer.</summary>
    public OfficeColor Background { get; set; } = OfficeColor.White;
    /// <summary>Sampling used when changing the composed page's resolution.</summary>
    public OfficeRasterResamplingMode Resampling { get; set; } = OfficeRasterResamplingMode.Area;
    /// <summary>Maximum pixels in the returned image.</summary>
    public long MaxPixels { get; set; } = 32L * 1024 * 1024;
    /// <summary>Maximum retained codec and raster working bytes in this operation, also capped by read options.</summary>
    public long MaxBytes { get; set; } = 256L * 1024 * 1024;

    /// <summary>Creates an independent, validated settings snapshot.</summary>
    public DjVuRenderOptions Clone() => Snapshot();

    internal DjVuRenderOptions Snapshot() {
        var copy = (DjVuRenderOptions)MemberwiseClone();
        if (copy.Dpi.HasValue && (copy.Dpi.Value < 1 || copy.Dpi.Value > 6000)) throw new ArgumentOutOfRangeException(nameof(Dpi));
        if (copy.Gamma < 0.3 || copy.Gamma > 5 || double.IsNaN(copy.Gamma)) throw new ArgumentOutOfRangeException(nameof(Gamma));
        if (copy.MaxPixels <= 0 || copy.MaxPixels > 50_000_000) throw new ArgumentOutOfRangeException(nameof(MaxPixels));
        if (copy.MaxBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxBytes));
        if (!Enum.IsDefined(typeof(OfficeRasterResamplingMode), copy.Resampling)) throw new ArgumentOutOfRangeException(nameof(Resampling));
        if (copy.Background.A != 255) throw new ArgumentException("DjVu page rendering requires an opaque background.", nameof(Background));
        return copy;
    }
}

/// <summary>A complete selected-page raster and the geometry used to produce it.</summary>
public sealed class DjVuRenderResult : IOfficeConversionReport {
    internal DjVuRenderResult(OfficeRasterImage image, int dpi, DjVuRectangle sourceRegion, int rotation, List<OfficeConversionFidelityDiagnostic> diagnostics) {
        Image = image; Dpi = dpi; SourceRegion = sourceRegion; Rotation = rotation; FidelityDiagnostics = diagnostics.AsReadOnly();
    }
    /// <summary>Owned top-to-bottom RGBA raster. Caller pixel edits do not affect the source document.</summary>
    public OfficeRasterImage Image { get; }
    /// <summary>Requested output resolution.</summary>
    public int Dpi { get; }
    /// <summary>Selected region in unrotated native page pixels.</summary>
    public DjVuRectangle SourceRegion { get; }
    /// <summary>Clockwise rotation applied to the output raster.</summary>
    public int Rotation { get; }
    /// <summary>Explicit preservation qualifications, including non-painted annotations.</summary>
    public IReadOnlyList<OfficeConversionFidelityDiagnostic> FidelityDiagnostics { get; }
    /// <inheritdoc />
    public bool HasLoss => FidelityDiagnostics.Any(d => d.LossKind != OfficeConversionLossKind.None);
    /// <inheritdoc />
    public void RequireNoLoss() {
        if (HasLoss) throw new OfficeConversionException("DjVu rendering reported fidelity loss. Inspect FidelityDiagnostics.", this);
    }
}
