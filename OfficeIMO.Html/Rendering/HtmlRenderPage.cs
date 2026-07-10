using System.Collections.ObjectModel;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>
/// One continuous or paged HTML render surface in CSS-pixel coordinates.
/// </summary>
public sealed class HtmlRenderPage {
    private readonly ReadOnlyCollection<HtmlRenderVisual> _visuals;

    internal HtmlRenderPage(int pageNumber, double width, double height, IEnumerable<HtmlRenderVisual> visuals) {
        if (pageNumber <= 0) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber));
        }

        if (width <= 0D || height <= 0D || double.IsNaN(width) || double.IsNaN(height) || double.IsInfinity(width) || double.IsInfinity(height)) {
            throw new ArgumentOutOfRangeException(nameof(width), "Rendered page dimensions must be finite positive numbers.");
        }

        PageNumber = pageNumber;
        Width = width;
        Height = height;
        _visuals = new List<HtmlRenderVisual>(visuals ?? throw new ArgumentNullException(nameof(visuals)))
            .OrderBy(item => item.PaintOrder)
            .ToList()
            .AsReadOnly();
    }

    /// <summary>One-based page number.</summary>
    public int PageNumber { get; }

    /// <summary>Page width in CSS pixels.</summary>
    public double Width { get; }

    /// <summary>Page height in CSS pixels.</summary>
    public double Height { get; }

    /// <summary>Ordered backend-neutral visuals on this page.</summary>
    public IReadOnlyList<HtmlRenderVisual> Visuals => _visuals;

    /// <summary>Creates a dependency-free drawing snapshot for PNG or SVG rendering.</summary>
    public OfficeDrawing CreateDrawing() {
        var drawing = new OfficeDrawing(Width, Height);
        foreach (HtmlRenderVisual visual in _visuals) {
            if (visual is HtmlRenderShape shape) {
                drawing.AddShape(shape.Shape.Clone(), shape.X, shape.Y);
            } else if (visual is HtmlRenderText text && text.Text.Length > 0) {
                drawing.AddText(text.Text, text.X, text.Y, text.Width, text.Height, text.Font, text.Color, text.Alignment, text.LineHeight);
            } else if (visual is HtmlRenderImage image) {
                var placement = new OfficeImagePlacement(image.X, image.Y, image.Width, image.Height);
                drawing.AddImage(image.Bytes, image.ContentType, new OfficeImageProjection(placement), image.AlternativeText);
            }
        }

        return drawing;
    }
}
