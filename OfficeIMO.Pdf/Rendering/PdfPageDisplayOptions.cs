namespace OfficeIMO.Pdf;

/// <summary>
/// Bounded PNG display settings. No source-content, font, codec, or text-shaping callbacks are exposed by the viewing path.
/// </summary>
public sealed class PdfPageDisplayOptions {
    /// <summary>Pixels per PDF point. Defaults to one.</summary>
    public double Scale { get; set; } = 1D;
    /// <summary>Optional maximum pixel dimension for a thumbnail or constrained viewport.</summary>
    public int? MaximumDimension { get; set; }
    /// <summary>Maximum raster pixels for the page. Defaults to 16 million pixels.</summary>
    public long MaximumPixels { get; set; } = 16_000_000;
    /// <summary>Maximum encoded PNG bytes. Defaults to 64 MiB.</summary>
    public long MaximumOutputBytes { get; set; } = 64L * 1024L * 1024L;
    /// <summary>Optional cooperative deadline for parsing and rasterization.</summary>
    public TimeSpan? Timeout { get; set; }

    internal PdfPageRenderOptions ToRenderOptions() {
        var options = new PdfPageRenderOptions {
            Format = PdfPageRenderFormat.Png, Scale = Scale, ThumbnailMaxDimension = MaximumDimension,
            MaxPixelsPerPage = MaximumPixels, MaxOutputBytesPerPage = MaximumOutputBytes,
            MaxTotalOutputBytes = MaximumOutputBytes, MaxPages = 1, ContinueOnError = false
        };
        if (Timeout.HasValue) options.RenderTimeout = Timeout.Value;
        options.Validate();
        return options;
    }
}
