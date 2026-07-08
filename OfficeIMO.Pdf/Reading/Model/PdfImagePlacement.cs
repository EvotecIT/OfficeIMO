using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Placement geometry for one image XObject invocation on a PDF page.
/// </summary>
public sealed class PdfImagePlacement {
    internal PdfImagePlacement(
        int pageNumber,
        string resourceName,
        int objectNumber,
        int directStreamIdentity,
        double a,
        double b,
        double c,
        double d,
        double e,
        double f,
        double x,
        double y,
        double width,
        double height,
        PdfPageClipPath? clipPath = null,
        OfficeColor? imageMaskColor = null,
        double? imageOpacity = null,
        PdfStream? inlineImageStream = null,
        PdfDictionary? inlineImageResources = null,
        double paintOrder = 0D) {
        PageNumber = pageNumber;
        ResourceName = resourceName;
        ObjectNumber = objectNumber;
        DirectStreamIdentity = directStreamIdentity;
        A = a;
        B = b;
        C = c;
        D = d;
        E = e;
        F = f;
        X = x;
        Y = y;
        Width = width;
        Height = height;
        ClipPath = clipPath;
        ImageMaskColor = imageMaskColor ?? OfficeColor.Black;
        ImageOpacity = imageOpacity;
        InlineImageStream = inlineImageStream;
        InlineImageResources = inlineImageResources;
        PaintOrder = paintOrder;
    }

    /// <summary>One-based source page number containing the image invocation.</summary>
    public int PageNumber { get; }

    /// <summary>Image resource name from the page or form XObject resource dictionary.</summary>
    public string ResourceName { get; }

    /// <summary>PDF object number for the image stream, or 0 when the image is direct.</summary>
    public int ObjectNumber { get; }

    /// <summary>Runtime identity for a direct image stream, or 0 when the image is indirect.</summary>
    internal int DirectStreamIdentity { get; }

    /// <summary>Current transformation matrix A component at the image invocation.</summary>
    public double A { get; }

    /// <summary>Current transformation matrix B component at the image invocation.</summary>
    public double B { get; }

    /// <summary>Current transformation matrix C component at the image invocation.</summary>
    public double C { get; }

    /// <summary>Current transformation matrix D component at the image invocation.</summary>
    public double D { get; }

    /// <summary>Current transformation matrix E translation component at the image invocation.</summary>
    public double E { get; }

    /// <summary>Current transformation matrix F translation component at the image invocation.</summary>
    public double F { get; }

    /// <summary>Left edge of the placement bounding box in PDF points.</summary>
    public double X { get; }

    /// <summary>Bottom edge of the placement bounding box in PDF points.</summary>
    public double Y { get; }

    /// <summary>Bounding-box width in PDF points.</summary>
    public double Width { get; }

    /// <summary>Bounding-box height in PDF points.</summary>
    public double Height { get; }

    internal PdfPageClipPath? ClipPath { get; }

    internal OfficeColor ImageMaskColor { get; }

    internal double? ImageOpacity { get; }

    internal PdfStream? InlineImageStream { get; }

    internal PdfDictionary? InlineImageResources { get; }

    internal double PaintOrder { get; }

    /// <summary>True when the placement matrix is axis-aligned within a small tolerance.</summary>
    public bool IsAxisAligned => Math.Abs(B) <= 0.001D && Math.Abs(C) <= 0.001D;
}
