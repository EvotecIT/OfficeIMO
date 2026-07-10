using System.Collections.ObjectModel;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>
/// One continuous or paged HTML render surface in CSS-pixel coordinates.
/// </summary>
public sealed class HtmlRenderPage {
    private readonly ReadOnlyCollection<HtmlRenderVisual> _visuals;
    private readonly OfficeFontFaceCollection _fonts;

    internal HtmlRenderPage(int pageNumber, double width, double height, IEnumerable<HtmlRenderVisual> visuals, string? pageName = null, OfficeFontFaceCollection? fonts = null) {
        if (pageNumber <= 0) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber));
        }

        if (width <= 0D || height <= 0D || double.IsNaN(width) || double.IsNaN(height) || double.IsInfinity(width) || double.IsInfinity(height)) {
            throw new ArgumentOutOfRangeException(nameof(width), "Rendered page dimensions must be finite positive numbers.");
        }

        PageNumber = pageNumber;
        Width = width;
        Height = height;
        PageName = pageName == null || string.IsNullOrWhiteSpace(pageName) ? null : pageName.Trim();
        _visuals = new List<HtmlRenderVisual>(visuals ?? throw new ArgumentNullException(nameof(visuals)))
            .OrderBy(item => item.PaintOrder)
            .ToList()
            .AsReadOnly();
        // The renderer passes one operation-scoped snapshot to every page. Public access still clones it.
        _fonts = fonts ?? new OfficeFontFaceCollection();
    }

    /// <summary>One-based page number.</summary>
    public int PageNumber { get; }

    /// <summary>Page width in CSS pixels.</summary>
    public double Width { get; }

    /// <summary>Page height in CSS pixels.</summary>
    public double Height { get; }

    /// <summary>CSS named-page identifier selected for this page, or <see langword="null"/> for the generic page master.</summary>
    public string? PageName { get; }

    /// <summary>Ordered backend-neutral visuals on this page.</summary>
    public IReadOnlyList<HtmlRenderVisual> Visuals => _visuals;

    /// <summary>Independent snapshot of scoped font faces used by this page.</summary>
    public OfficeFontFaceCollection Fonts => _fonts.Clone();

    /// <summary>Creates a dependency-free drawing snapshot for PNG or SVG rendering.</summary>
    public OfficeDrawing CreateDrawing() {
        var drawing = new OfficeDrawing(Width, Height);
        drawing.Fonts.AddRange(_fonts);
        foreach (HtmlRenderVisual visual in _visuals) {
            AddVisual(drawing, visual, Width, Height, _fonts);
        }

        return drawing;
    }

    private static void AddVisual(
        OfficeDrawing drawing,
        HtmlRenderVisual visual,
        double surfaceWidth,
        double surfaceHeight,
        OfficeFontFaceCollection fonts) {
        if (visual is HtmlRenderShape shape) {
            drawing.AddShape(shape.Shape.Clone(), shape.X, shape.Y);
        } else if (visual is HtmlRenderText text && text.Text.Length > 0) {
            drawing.AddText(text.Text, text.X, text.Y, text.Width, text.Height, text.Font, text.Color, text.Alignment, text.LineHeight);
        } else if (visual is HtmlRenderImage image) {
            var placement = new OfficeImagePlacement(image.X, image.Y, image.Width, image.Height);
            drawing.AddImage(image.EncodedBytes, image.ContentType, new OfficeImageProjection(placement, image.SourceCrop), image.AlternativeText);
        } else if (visual is HtmlRenderDrawing vector) {
            OfficeTransform transform = OfficeTransform.Scale(
                    vector.Width / vector.InnerDrawing.Width,
                    vector.Height / vector.InnerDrawing.Height)
                .Then(OfficeTransform.Translate(vector.X, vector.Y));
            drawing.AddEffectDrawing(vector.InnerDrawing, transform);
        } else if (visual is HtmlRenderImagePattern imagePattern) {
            drawing.AddImagePattern(
                imagePattern.EncodedBytes,
                imagePattern.ContentType,
                imagePattern.Pattern,
                imagePattern.MaximumTileCount);
        } else if (visual is HtmlRenderClipGroup group) {
            AddClipGroup(drawing, group, surfaceWidth, surfaceHeight, fonts);
        } else if (visual is HtmlRenderPathClipGroup pathClipGroup) {
            AddPathClipGroup(drawing, pathClipGroup, surfaceWidth, surfaceHeight, fonts);
        } else if (visual is HtmlRenderEffectGroup effectGroup) {
            AddEffectGroup(drawing, effectGroup, surfaceWidth, surfaceHeight, fonts);
        }
    }

    private static void AddEffectGroup(
        OfficeDrawing drawing,
        HtmlRenderEffectGroup group,
        double surfaceWidth,
        double surfaceHeight,
        OfficeFontFaceCollection fonts) {
        double nestedWidth = Math.Max(surfaceWidth, MaximumRight(group.Visuals));
        double nestedHeight = Math.Max(surfaceHeight, MaximumBottom(group.Visuals));
        var nested = new OfficeDrawing(Math.Max(0.01D, nestedWidth), Math.Max(0.01D, nestedHeight));
        nested.Fonts.AddRange(fonts);
        foreach (HtmlRenderVisual child in group.Visuals) AddVisual(nested, child, nested.Width, nested.Height, fonts);
        drawing.AddEffectDrawing(nested, group.Transform, group.Opacity);
    }

    private static void AddClipGroup(
        OfficeDrawing drawing,
        HtmlRenderClipGroup group,
        double surfaceWidth,
        double surfaceHeight,
        OfficeFontFaceCollection fonts) {
        double left = group.ClipHorizontal ? Math.Max(0D, group.ClipX) : 0D;
        double top = group.ClipVertical ? Math.Max(0D, group.ClipY) : 0D;
        double right = group.ClipHorizontal ? Math.Min(surfaceWidth, group.ClipX + group.ClipWidth) : surfaceWidth;
        double bottom = group.ClipVertical ? Math.Min(surfaceHeight, group.ClipY + group.ClipHeight) : surfaceHeight;
        if (right <= left + 0.0001D || bottom <= top + 0.0001D) return;

        double nestedWidth = Math.Max(surfaceWidth, MaximumRight(group.Visuals));
        double nestedHeight = Math.Max(surfaceHeight, MaximumBottom(group.Visuals));
        var nested = new OfficeDrawing(Math.Max(0.01D, nestedWidth), Math.Max(0.01D, nestedHeight));
        nested.Fonts.AddRange(fonts);
        foreach (HtmlRenderVisual child in group.Visuals) AddVisual(nested, child, nested.Width, nested.Height, fonts);
        drawing.AddClippedDrawing(
            nested,
            left,
            top,
            OfficeClipPath.Rectangle(right - left, bottom - top),
            -left,
            -top);
    }

    private static void AddPathClipGroup(
        OfficeDrawing drawing,
        HtmlRenderPathClipGroup group,
        double surfaceWidth,
        double surfaceHeight,
        OfficeFontFaceCollection fonts) {
        double nestedWidth = Math.Max(surfaceWidth, MaximumRight(group.Visuals));
        double nestedHeight = Math.Max(surfaceHeight, MaximumBottom(group.Visuals));
        var nested = new OfficeDrawing(Math.Max(0.01D, nestedWidth), Math.Max(0.01D, nestedHeight));
        nested.Fonts.AddRange(fonts);
        foreach (HtmlRenderVisual child in group.Visuals) AddVisual(nested, child, nested.Width, nested.Height, fonts);
        drawing.AddClippedDrawing(nested, group.ClipX, group.ClipY, group.ClipPath, -group.ClipX, -group.ClipY);
    }

    private static double MaximumRight(IEnumerable<HtmlRenderVisual> visuals) => visuals
        .Select(visual => visual is HtmlRenderClipGroup clipGroup
            ? Math.Max(visual.X + visual.Width, MaximumRight(clipGroup.Visuals))
            : visual is HtmlRenderPathClipGroup pathClipGroup
                ? Math.Max(visual.X + visual.Width, MaximumRight(pathClipGroup.Visuals))
            : visual is HtmlRenderEffectGroup effectGroup
                ? Math.Max(visual.X + visual.Width, MaximumRight(effectGroup.Visuals))
                : visual.X + visual.Width)
        .DefaultIfEmpty(0.01D)
        .Max();

    private static double MaximumBottom(IEnumerable<HtmlRenderVisual> visuals) => visuals
        .Select(visual => visual is HtmlRenderClipGroup clipGroup
            ? Math.Max(visual.Y + visual.Height, MaximumBottom(clipGroup.Visuals))
            : visual is HtmlRenderPathClipGroup pathClipGroup
                ? Math.Max(visual.Y + visual.Height, MaximumBottom(pathClipGroup.Visuals))
            : visual is HtmlRenderEffectGroup effectGroup
                ? Math.Max(visual.Y + visual.Height, MaximumBottom(effectGroup.Visuals))
                : visual.Y + visual.Height)
        .DefaultIfEmpty(0.01D)
        .Max();
}
