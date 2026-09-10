using OfficeIMO.Pdf;
using OfficeIMO.Drawing;

namespace OfficeIMO.Workflows;

/// <summary>One immutable raster print sheet at the planned physical paper size.</summary>
public sealed class PdfRenderedPrintSheet {
    private readonly byte[] _png;
    private readonly int _width, _height;
    internal PdfRenderedPrintSheet(PdfPrintSheet plan, byte[] png, int width, int height) {
        Plan = plan; _png = png; _width = width; _height = height;
    }
    /// <summary>Physical paper and source-page placement.</summary>
    public PdfPrintSheet Plan { get; }
    /// <summary>Returns a copy of the exact pixels used for preview and delivery.</summary>
    public byte[] GetPng() => (byte[])_png.Clone();
    internal byte[] Png => _png;
    internal OfficeRasterImage Decode(CancellationToken token) {
        if (!OfficeRasterImageDecoder.TryDecode(_png, new OfficeRasterDecodeOptions {
            MaximumDecodedPixels = (long)_width * _height, MaximumEncodedBytes = _png.Length, CancellationToken = token
        }, out OfficeRasterImage? raster, out _) || raster is null || raster.Width != _width || raster.Height != _height)
            throw new InvalidOperationException("A prepared sheet could not be decoded at its validated dimensions.");
        return raster;
    }
}

/// <summary>Reviewed raster sheets, detached from later source-file or workspace changes.</summary>
public sealed class PdfPreparedPrintDocument {
    internal PdfPreparedPrintDocument(string sourcePath, PdfPrintPlan plan, IReadOnlyList<PdfRenderedPrintSheet> sheets,
        IReadOnlyList<string> diagnostics) {
        SourcePath = sourcePath; Plan = plan; Sheets = Array.AsReadOnly(sheets.ToArray());
        Diagnostics = Array.AsReadOnly(diagnostics.Distinct().ToArray());
    }
    /// <summary>Original source identity; never used to reopen content during delivery.</summary>
    public string SourcePath { get; }
    /// <summary>Plan that generated these pixels.</summary>
    public PdfPrintPlan Plan { get; }
    /// <summary>Ordered physical sheets.</summary>
    public IReadOnlyList<PdfRenderedPrintSheet> Sheets { get; }
    /// <summary>Managed-renderer limitations reported while preparing the sheets.</summary>
    public IReadOnlyList<string> Diagnostics { get; }
}

/// <summary>Resource limits for preparing reviewable print sheets.</summary>
public sealed class PdfPrintRenderOptions {
    /// <summary>Raster resolution. Defaults to 150 DPI.</summary>
    public double Dpi { get; set; } = 150;
    /// <summary>Maximum selected source pages.</summary>
    public int MaximumPages { get; set; } = 100;
    /// <summary>Maximum pixels in a source-page or sheet raster, up to the shared decoder limit of 50 million.</summary>
    public long MaximumPixelsPerImage { get; set; } = 16_000_000;
    /// <summary>Maximum retained encoded sheet bytes.</summary>
    public long MaximumOutputBytes { get; set; } = 128L * 1024 * 1024;
}
