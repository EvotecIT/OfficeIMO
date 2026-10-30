using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Public snapshot of the effective clipping path applied to an image placement.</summary>
/// <remarks>Geometry uses the source page's default user-space units: one unit is /UserUnit physical points (1/72 inch each), with /UserUnit defaulting to 1. Coordinates are not automatically normalized to physical points. Explicitly named visual bounds use their documented visual coordinate system.</remarks>
public sealed class PdfImageClipInfo {
    internal PdfImageClipInfo(PdfPageClipPath path) {
        OriginalPath = path;
        X = path.X;
        Y = path.Y;
        Width = path.Width;
        Height = path.Height;
        IsRectangle = path.IsRectangle;
        FillRule = path.FillRule;
        IsExact = path.IsExact;
        ContainsTextClipping = path.ContainsTextClipping;
        Commands = Array.AsReadOnly(path.Commands.ToArray());
    }

    internal PdfPageClipPath OriginalPath { get; }

    /// <summary>Left edge in transformed PDF user-space coordinates; crop origin, page rotation, and UserUnit are not normalized.</summary>
    public double X { get; }

    /// <summary>Top edge after flipping transformed PDF user space against the parser page-box height; crop origin, page rotation, and UserUnit are not normalized.</summary>
    public double Y { get; }

    /// <summary>Clip bounds width in transformed, unscaled PDF user-space units.</summary>
    public double Width { get; }

    /// <summary>Clip bounds height in transformed, unscaled PDF user-space units.</summary>
    public double Height { get; }

    /// <summary>True when the effective clip is represented by one rectangle.</summary>
    public bool IsRectangle { get; }

    /// <summary>Fill rule used by a path clip.</summary>
    public OfficeFillRule FillRule { get; }

    /// <summary>True when the engine could retain exact clip geometry rather than a conservative bound.</summary>
    public bool IsExact { get; }

    /// <summary>True when text rendering modes contributed to the effective clipping path.</summary>
    public bool ContainsTextClipping { get; }

    /// <summary>Path commands for a nonrectangular clip. Rectangular clips expose an empty list.</summary>
    public IReadOnlyList<OfficePathCommand> Commands { get; }
}
