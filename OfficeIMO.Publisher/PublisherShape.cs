using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher;

/// <summary>An immutable vector-artwork view owned by one document or master page. Text and artwork can share the same publication object identifier.</summary>
public sealed class PublisherShape {
    private readonly OfficeShape? _shape;

    internal PublisherShape(uint id, uint pageId, double x, double y, double width, double height,
        OfficeTransform pageTransform, bool hasPlacement, OfficeShape? shape, string? editLimitation,
        bool isAuthored = false) {
        Id = id; PageId = pageId; X = x; Y = y; Width = width; Height = height;
        PageTransform = pageTransform; HasPlacement = hasPlacement; _shape = shape?.Clone();
        EditLimitation = editLimitation; IsAuthored = isAuthored;
    }

    /// <summary>Stable object identifier scoped to the publication; a text-capable shape shares this identifier with its text frame.</summary>
    public uint Id { get; }
    /// <summary>Owning document or master page. Applied master artwork remains listed on its master.</summary>
    public uint PageId { get; }
    /// <summary>Left edge of the unrotated frame in page-local points.</summary>
    public double X { get; }
    /// <summary>Top edge of the unrotated frame in page-local points.</summary>
    public double Y { get; }
    /// <summary>Logical frame width in points. A vertical line has zero width.</summary>
    public double Width { get; }
    /// <summary>Logical frame height in points. A horizontal line has zero height.</summary>
    public double Height { get; }
    /// <summary>Maps shape-local points to page-local points, including recovered object and ancestor transforms.</summary>
    public OfficeTransform PageTransform { get; }
    /// <summary>Whether usable native or authored page placement is available. False means the coordinate values must not be treated as recovered geometry.</summary>
    public bool HasPlacement { get; }
    /// <summary>Whether bounds and basic solid paint can be edited. Unsupported artwork retains its existing reading and rendering behavior.</summary>
    public bool CanEdit => _shape != null && EditLimitation == null;
    /// <summary>Reason the artwork is inspection-only, or null for supported edits.</summary>
    public string? EditLimitation { get; }

    /// <summary>Creates a detached shared shape descriptor for supported editable artwork. Editing this copy does not change the publication.</summary>
    /// <exception cref="NotSupportedException">The artwork has no qualified editable geometry and paint model. Its rendered approximation remains available through the page drawing.</exception>
    public OfficeShape CreateShape() => _shape?.Clone() ?? throw new NotSupportedException(EditLimitation ?? "The shape has no editable geometry model.");

    internal bool IsAuthored { get; }
    internal PublisherShape WithPaint(OfficeColor? fill, OfficeColor? stroke, double strokeWidth, double fillOpacity, double strokeOpacity) {
        OfficeShape shape = CreateShape();
        shape.FillColor = fill; shape.StrokeColor = stroke; shape.StrokeWidth = strokeWidth;
        shape.FillOpacity = fillOpacity; shape.StrokeOpacity = strokeOpacity;
        return new PublisherShape(Id, PageId, X, Y, Width, Height, PageTransform, HasPlacement, shape, EditLimitation, IsAuthored);
    }
    internal PublisherShape WithPlacement(double x, double y, double width, double height, OfficeTransform transform, bool hasPlacement) =>
        new(Id, PageId, x, y, width, height, transform, hasPlacement,
            _shape == null ? null : PublisherShapeGeometry.Resize(_shape, width, height), EditLimitation, IsAuthored);
}
