using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher;

/// <summary>A picture placement owned by one document or master page. The embedded asset is separate from this placement.</summary>
public sealed class PublisherPicture {
    private readonly OfficeShape? _paint;

    internal PublisherPicture(uint id, uint pageId, int? imageId, double x, double y, double width, double height,
        OfficeTransform pageTransform, OfficeImageSourceCrop crop, bool hasPlacement, bool canEdit, OfficeShape? paint = null,
        string? editLimitation = null, bool isAuthored = false) {
        Id = id; PageId = pageId; ImageId = imageId; X = x; Y = y; Width = width; Height = height;
        PageTransform = pageTransform; Crop = crop; HasPlacement = hasPlacement; CanEdit = canEdit;
        _paint = paint?.Clone();
        EditLimitation = editLimitation;
        IsAuthored = isAuthored;
    }

    /// <summary>Stable publication object identifier, scoped to this document.</summary>
    public uint Id { get; }
    /// <summary>Owning document or master page identifier. Inherited pictures remain owned by their master.</summary>
    public uint PageId { get; }
    /// <summary>Embedded asset identifier in <see cref="PublisherDocument.Images"/>, or null when no usable image reference was recovered. A non-null reference can still be unresolved.</summary>
    public int? ImageId { get; }
    /// <summary>Left edge of the unrotated frame in page-local points.</summary>
    public double X { get; }
    /// <summary>Top edge of the unrotated frame in page-local points.</summary>
    public double Y { get; }
    /// <summary>Frame width in points.</summary>
    public double Width { get; }
    /// <summary>Frame height in points.</summary>
    public double Height { get; }
    /// <summary>Maps picture-local points into page-local points, including object and group transforms.</summary>
    public OfficeTransform PageTransform { get; }
    /// <summary>Normalized source-edge crop of the current placement. Native crop fallback diagnostics remain in the read report.</summary>
    public OfficeImageSourceCrop Crop { get; }
    /// <summary>Whether this current picture has usable page placement.</summary>
    public bool HasPlacement { get; }
    /// <summary>Whether image replacement, bounds and source-crop edits are supported for this placement. Hidden, grouped, transformed, effected, custom-clipped or unresolved pictures are inspection-only.</summary>
    public bool CanEdit { get; }
    /// <summary>Reason this placement is inspection-only, or null when the supported picture edits are available.</summary>
    public string? EditLimitation { get; }
    internal bool IsAuthored { get; }

    internal OfficeShape CreatePaint() => _paint?.Clone() ?? throw new InvalidOperationException("The picture has no semantic paint model.");
    internal PublisherPicture WithImage(int imageId) => new(Id, PageId, imageId, X, Y, Width, Height, PageTransform, Crop, HasPlacement, CanEdit, _paint, EditLimitation, IsAuthored);
    internal PublisherPicture WithCrop(OfficeImageSourceCrop crop) =>
        new(Id, PageId, ImageId, X, Y, Width, Height, PageTransform, crop, HasPlacement, CanEdit, _paint, EditLimitation, IsAuthored);
    internal PublisherPicture WithBounds(double x, double y, double width, double height) {
        OfficeShape paint = CreatePaint(); paint.Width = width; paint.Height = height;
        return new PublisherPicture(Id, PageId, ImageId, x, y, width, height,
            OfficeTransform.Translate(x, y), Crop, true, CanEdit, paint, EditLimitation, IsAuthored);
    }

    // Mixed image/text objects take current placement from their frame owner.
    // This creates a view; it does not introduce independently editable geometry.
    internal PublisherPicture WithPlacement(double x, double y, double width, double height, OfficeTransform transform, bool hasPlacement) =>
        new(Id, PageId, ImageId, x, y, width, height, transform, Crop, hasPlacement, CanEdit, _paint, EditLimitation, IsAuthored);
}
