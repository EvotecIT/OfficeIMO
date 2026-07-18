using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Output format for managed PDF page batches.</summary>
public enum PdfPageRenderFormat {
    /// <summary>Dependency-free raster PNG.</summary>
    Png,
    /// <summary>UTF-8 vector SVG.</summary>
    Svg
}

/// <summary>Options for bounded page-range and thumbnail rendering.</summary>
public sealed class PdfPageRenderOptions : OfficeImageExportOptions {
    /// <summary>Creates legacy page-render options with fail-fast pixel-limit behavior.</summary>
    public PdfPageRenderOptions() {
        RasterOverflowBehavior = OfficeRasterOverflowBehavior.Throw;
    }

    /// <summary>Batch output format.</summary>
    public PdfPageRenderFormat Format { get; set; } = PdfPageRenderFormat.Png;
    /// <summary>Optional target DPI, converted from PDF's 72 points per inch.</summary>
    public double? Dpi { get; set; }
    /// <summary>PNG background color.</summary>
    public OfficeColor Background {
        get => BackgroundColor;
        set => BackgroundColor = value;
    }
    /// <summary>Optional maximum thumbnail width or height in pixels.</summary>
    public int? ThumbnailMaxDimension { get; set; }
    /// <summary>Maximum pages rendered by one call.</summary>
    public int MaxPages { get; set; } = 100;
    /// <summary>Maximum output pixels for one page.</summary>
    public long MaxPixelsPerPage {
        get => MaximumRasterPixels;
        set => MaximumRasterPixels = value;
    }
    /// <summary>Continues a batch and returns a failed per-page report when rendering fails.</summary>
    public bool ContinueOnError { get; set; } = true;
    internal double GetScale(OfficeDrawing drawing) {
        double scale = Dpi.HasValue ? Dpi.Value / 72D : Scale;
        if (ThumbnailMaxDimension.HasValue) {
            double thumbnailScale = ThumbnailMaxDimension.Value / Math.Max(drawing.Width, drawing.Height);
            scale = Math.Min(scale, thumbnailScale);
        }

        return scale;
    }

    internal void Validate() {
        ValidateImageExportOptions();
        if (Format < PdfPageRenderFormat.Png || Format > PdfPageRenderFormat.Svg) throw new ArgumentOutOfRangeException(nameof(Format));
        if (Dpi.HasValue && !IsPositiveFinite(Dpi.Value)) throw new ArgumentOutOfRangeException(nameof(Dpi), Dpi, "DPI must be positive and finite.");
        if (ThumbnailMaxDimension.HasValue && ThumbnailMaxDimension.Value <= 0) throw new ArgumentOutOfRangeException(nameof(ThumbnailMaxDimension));
        if (MaxPages <= 0) throw new ArgumentOutOfRangeException(nameof(MaxPages));
    }

    private static bool IsPositiveFinite(double value) => value > 0D && !double.IsNaN(value) && !double.IsInfinity(value);
}

/// <summary>Per-page managed render result.</summary>
public sealed class PdfPageRenderResult {
    private readonly byte[]? _bytes;

    internal PdfPageRenderResult(
        int pageNumber,
        PdfPageRenderFormat format,
        byte[]? bytes,
        int width,
        int height,
        TimeSpan elapsed,
        IReadOnlyList<PdfRenderCapabilityDiagnostic> capabilityDiagnostics,
        IReadOnlyList<string>? errors = null) {
        PageNumber = pageNumber;
        Format = format;
        _bytes = bytes == null ? null : (byte[])bytes.Clone();
        Width = width;
        Height = height;
        Elapsed = elapsed;
        CapabilityDiagnostics = capabilityDiagnostics.ToArray();
        var diagnostics = new List<string>(capabilityDiagnostics.Count + (errors?.Count ?? 0));
        for (int i = 0; i < capabilityDiagnostics.Count; i++) diagnostics.Add(capabilityDiagnostics[i].Code + ": " + capabilityDiagnostics[i].Message);
        if (errors != null) diagnostics.AddRange(errors);
        Diagnostics = diagnostics.Count == 0 ? Array.Empty<string>() : diagnostics.AsReadOnly();
    }

    /// <summary>One-based source page number.</summary>
    public int PageNumber { get; }
    /// <summary>Requested output format.</summary>
    public PdfPageRenderFormat Format { get; }
    /// <summary>Rendered bytes, or null when the page failed.</summary>
    public byte[]? Bytes => _bytes == null ? null : (byte[])_bytes.Clone();
    /// <summary>Output width in pixels or SVG user units.</summary>
    public int Width { get; }
    /// <summary>Output height in pixels or SVG user units.</summary>
    public int Height { get; }
    /// <summary>Elapsed render time.</summary>
    public TimeSpan Elapsed { get; }
    /// <summary>Stable per-page diagnostics.</summary>
    public IReadOnlyList<string> Diagnostics { get; }
    /// <summary>Typed skipped or simplified operator/resource diagnostics backed by the generated capability manifest.</summary>
    public IReadOnlyList<PdfRenderCapabilityDiagnostic> CapabilityDiagnostics { get; }
    /// <summary>True when output bytes were produced.</summary>
    public bool Succeeded => _bytes is not null;
}
