using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Pixel alignment used when compared pages have different rendered dimensions.</summary>
public enum PdfVisualPageAlignment {
    /// <summary>Aligns page images at their top-left corners.</summary>
    TopLeft,
    /// <summary>Centers each page image in the comparison canvas.</summary>
    Center
}

/// <summary>Pixel rectangle excluded from visual comparison.</summary>
public readonly struct PdfPixelRegion {
    /// <summary>Creates an ignored pixel rectangle.</summary>
    public PdfPixelRegion(int x, int y, int width, int height) {
        Guard.NonNegative(x, nameof(x));
        Guard.NonNegative(y, nameof(y));
        Guard.PositiveInteger(width, nameof(width));
        Guard.PositiveInteger(height, nameof(height));
        X = x; Y = y; Width = width; Height = height;
    }

    /// <summary>Left pixel.</summary>
    public int X { get; }
    /// <summary>Top pixel.</summary>
    public int Y { get; }
    /// <summary>Region width.</summary>
    public int Width { get; }
    /// <summary>Region height.</summary>
    public int Height { get; }
    internal bool Contains(int x, int y) => x >= X && y >= Y && x < X + Width && y < Y + Height;
}

/// <summary>Options for dependency-free rendered PDF comparison.</summary>
public sealed class PdfVisualComparisonOptions {
    private readonly List<PdfPixelRegion> _ignoredRegions = new();
    /// <summary>Render scale.</summary>
    public double Scale { get; set; } = 1D;
    /// <summary>Maximum per-channel difference treated as equal, from 0 through 255.</summary>
    public byte ChannelTolerance { get; set; }
    /// <summary>Maximum changed-pixel ratio accepted as equal, from 0 through 1.</summary>
    public double AllowedDifferenceRatio { get; set; }
    /// <summary>Alignment for pages with differing dimensions.</summary>
    public PdfVisualPageAlignment Alignment { get; set; } = PdfVisualPageAlignment.TopLeft;
    /// <summary>Raster background used for both inputs.</summary>
    public OfficeColor Background { get; set; } = OfficeColor.White;
    /// <summary>Ignored comparison regions in output pixel coordinates.</summary>
    public IList<PdfPixelRegion> IgnoredRegions => _ignoredRegions;
    /// <summary>Maximum pages compared by one call.</summary>
    public int MaxPages { get; set; } = 100;
    /// <summary>Maximum raster pixels allocated for one expected, actual, or diff page image.</summary>
    public long MaxPixelsPerImage { get; set; } = 20_000_000L;
    /// <summary>Maximum aggregate raster pixels allocated across the comparison.</summary>
    public long MaxTotalPixels { get; set; } = 100_000_000L;
    /// <summary>Maximum aggregate PNG bytes retained in the comparison report.</summary>
    public long MaxTotalOutputBytes { get; set; } = 256L * 1024L * 1024L;

    internal void Validate() {
        if (Scale <= 0D || double.IsNaN(Scale) || double.IsInfinity(Scale)) throw new ArgumentOutOfRangeException(nameof(Scale));
        if (AllowedDifferenceRatio < 0D || AllowedDifferenceRatio > 1D || double.IsNaN(AllowedDifferenceRatio)) throw new ArgumentOutOfRangeException(nameof(AllowedDifferenceRatio));
        if (Alignment < PdfVisualPageAlignment.TopLeft || Alignment > PdfVisualPageAlignment.Center) throw new ArgumentOutOfRangeException(nameof(Alignment));
        if (MaxPages <= 0) throw new ArgumentOutOfRangeException(nameof(MaxPages));
        if (MaxPixelsPerImage <= 0) throw new ArgumentOutOfRangeException(nameof(MaxPixelsPerImage));
        if (MaxTotalPixels <= 0) throw new ArgumentOutOfRangeException(nameof(MaxTotalPixels));
        if (MaxTotalOutputBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxTotalOutputBytes));
    }
}
