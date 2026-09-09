namespace OfficeIMO.Pdf;

/// <summary>Bounded raster settings for printing the managed renderer's page appearance.</summary>
public sealed class PdfPagePrintOptions {
    /// <summary>Output pixels per inch, from 72 to 600. Defaults to 150.</summary>
    public double Dpi { get; set; } = 150;
    /// <summary>Maximum pixels in one rendered page.</summary>
    public long MaximumPixels { get; set; } = 16_000_000;
    /// <summary>Maximum encoded PNG bytes per page.</summary>
    public long MaximumOutputBytes { get; set; } = 64 * 1024 * 1024;

    internal PdfPageRenderOptions ToRenderOptions() {
        if (Dpi < 72 || Dpi > 600 || double.IsNaN(Dpi)) throw new ArgumentOutOfRangeException(nameof(Dpi));
        return new PdfPageDisplayOptions {
            Scale = Dpi / 72, MaximumPixels = MaximumPixels, MaximumOutputBytes = MaximumOutputBytes
        }.ToRenderOptions();
    }
}
