using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Describes a simple shape rendered in a page header or footer.
/// </summary>
public sealed class PdfHeaderFooterShape {
    private readonly ShapeBlock _block;

    /// <summary>Creates a header/footer shape.</summary>
    public PdfHeaderFooterShape(OfficeShape shape, PdfAlign align = PdfAlign.Left) {
        Guard.NotNull(shape, nameof(shape));
        Guard.LeftCenterRightAlign(align, nameof(align), "PDF header/footer shape");

        _block = PdfDocument.CreateShapeBlock(shape, align, 0D, 0D);
    }

    /// <summary>Shape width in PDF points.</summary>
    public double Width => _block.Shape.Width;

    /// <summary>Shape height in PDF points.</summary>
    public double Height => _block.Shape.Height;

    /// <summary>Horizontal placement inside the page content width.</summary>
    public PdfAlign Align => _block.Align;

    /// <summary>Returns a copy of the shape payload.</summary>
    public OfficeShape Shape => _block.Shape.Clone();

    // The payload is private and all public/internal exports return copies.
    internal PdfHeaderFooterShape Clone() => this;

    internal ShapeBlock ToShapeBlock() => new ShapeBlock(_block.Shape, _block.Style);
}
