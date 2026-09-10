namespace OfficeIMO.Pdf;

/// <summary>Bounded raster settings for printing the managed renderer's page appearance.</summary>
public sealed class PdfPagePrintOptions {
    /// <summary>Output pixels per inch, from 72 to 600. Defaults to 150.</summary>
    public double Dpi { get; set; } = 150;
    /// <summary>Source-to-paper placement scale. The source is rasterized at this multiple of the paper DPI.</summary>
    /// <remarks>Documents permitting only low-quality printing retain a maximum source resolution of 150 DPI.</remarks>
    public double PageScale { get; set; } = 1;
    /// <summary>Maximum pixels in one rendered page.</summary>
    public long MaximumPixels { get; set; } = 16_000_000;
    /// <summary>Maximum encoded PNG bytes per page.</summary>
    public long MaximumOutputBytes { get; set; } = 64 * 1024 * 1024;

    internal PdfPageRenderOptions ToRenderOptions() {
        if (Dpi < 72 || Dpi > 600 || double.IsNaN(Dpi)) throw new ArgumentOutOfRangeException(nameof(Dpi));
        if (PageScale <= 0 || double.IsNaN(PageScale) || double.IsInfinity(PageScale)) throw new ArgumentOutOfRangeException(nameof(PageScale));
        return new PdfPageDisplayOptions {
            Scale = Dpi / 72 * PageScale, MaximumPixels = MaximumPixels, MaximumOutputBytes = MaximumOutputBytes
        }.ToRenderOptions();
    }
}
