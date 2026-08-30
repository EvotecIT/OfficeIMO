using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Public snapshot of the effective clipping path applied to an image placement.</summary>
public sealed class PdfImageClipInfo {
    internal PdfImageClipInfo(PdfPageClipPath path) {
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

    /// <summary>Left edge of the clip bounds in the placement coordinate space.</summary>
    public double X { get; }

    /// <summary>Top edge of the clip bounds in the placement coordinate space.</summary>
    public double Y { get; }

    /// <summary>Clip bounds width in points.</summary>
    public double Width { get; }

    /// <summary>Clip bounds height in points.</summary>
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
