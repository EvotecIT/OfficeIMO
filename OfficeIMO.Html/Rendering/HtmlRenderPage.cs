using System.Collections.ObjectModel;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>
/// One continuous or paged HTML render surface in CSS-pixel coordinates.
/// </summary>
public sealed class HtmlRenderPage {
    private readonly ReadOnlyCollection<HtmlRenderVisual> _visuals;
    private readonly ReadOnlyCollection<HtmlRenderVisual> _scene;
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
        _scene = new List<HtmlRenderVisual>(visuals ?? throw new ArgumentNullException(nameof(visuals)))
            .OrderBy(item => item.PaintOrder)
            .ToList()
            .AsReadOnly();
        _visuals = FlattenSemanticGroups(_scene).ToList().AsReadOnly();
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

    /// <summary>Ordered backend-neutral scene including paint-neutral semantic ownership groups.</summary>
    public IReadOnlyList<HtmlRenderVisual> Scene => _scene;

    /// <summary>Independent snapshot of scoped font faces used by this page.</summary>
    public OfficeFontFaceCollection Fonts => _fonts.Clone();

    /// <summary>Creates a dependency-free drawing snapshot for PNG or SVG rendering.</summary>
    public OfficeDrawing CreateDrawing() => CreateDrawing(CancellationToken.None);

    /// <summary>Creates a dependency-free drawing snapshot for PNG or SVG rendering with cooperative cancellation.</summary>
    public OfficeDrawing CreateDrawing(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        var drawing = new OfficeDrawing(Width, Height);
        drawing.Fonts.AddRange(_fonts);
        foreach (HtmlRenderVisual visual in _scene) {
            cancellationToken.ThrowIfCancellationRequested();
            AddVisual(drawing, visual, Width, Height, _fonts, cancellationToken);
        }

        cancellationToken.ThrowIfCancellationRequested();
        return drawing;
    }

    private static IEnumerable<HtmlRenderVisual> FlattenSemanticGroups(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            if (visual is HtmlRenderSemanticGroup semanticGroup) {
                foreach (HtmlRenderVisual child in FlattenSemanticGroups(semanticGroup.Visuals)) yield return child;
            } else if (visual is HtmlRenderLogicalTextGroup logicalTextGroup) {
                foreach (HtmlRenderVisual child in FlattenSemanticGroups(logicalTextGroup.Visuals)) yield return child;
            } else {
                yield return visual;
            }
        }
    }

    private static void AddVisual(
        OfficeDrawing drawing,
        HtmlRenderVisual visual,
        double surfaceWidth,
        double surfaceHeight,
        OfficeFontFaceCollection fonts,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (visual is HtmlRenderShape shape) {
            AddShape(drawing, shape, surfaceWidth, surfaceHeight, fonts);
        } else if (visual is HtmlRenderText text && text.Text.Length > 0) {
            if (text.TextAdvanceWidth.HasValue) {
                drawing.AddPositionedText(
                    text.Text,
                    text.X,
                    text.Y,
                    text.Width,
                    text.Height,
                    text.Font,
                    text.Color,
                    text.Alignment,
                    text.LineHeight,
                    textAdvanceWidth: text.TextAdvanceWidth.Value);
            } else {
                drawing.AddText(text.Text, text.X, text.Y, text.Width, text.Height, text.Font, text.Color, text.Alignment, text.LineHeight);
            }
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
            AddClipGroup(drawing, group, surfaceWidth, surfaceHeight, fonts, cancellationToken);
        } else if (visual is HtmlRenderPathClipGroup pathClipGroup) {
            AddPathClipGroup(drawing, pathClipGroup, surfaceWidth, surfaceHeight, fonts, cancellationToken);
        } else if (visual is HtmlRenderEffectGroup effectGroup) {
            AddEffectGroup(drawing, effectGroup, surfaceWidth, surfaceHeight, fonts, cancellationToken);
        } else if (visual is HtmlRenderSemanticGroup semanticGroup) {
            foreach (HtmlRenderVisual child in semanticGroup.Visuals) AddVisual(drawing, child, surfaceWidth, surfaceHeight, fonts, cancellationToken);
        } else if (visual is HtmlRenderLogicalTextGroup logicalTextGroup) {
            foreach (HtmlRenderVisual child in logicalTextGroup.Visuals) AddVisual(drawing, child, surfaceWidth, surfaceHeight, fonts, cancellationToken);
        }
    }

    private static void AddShape(
        OfficeDrawing drawing,
        HtmlRenderShape visual,
        double surfaceWidth,
        double surfaceHeight,
        OfficeFontFaceCollection fonts) {
        OfficeShape shape = visual.Shape.Clone();
        if (visual.X >= 0D &&
            visual.Y >= 0D &&
            visual.X + shape.Width <= surfaceWidth &&
            visual.Y + shape.Height <= surfaceHeight) {
            drawing.AddShape(shape, visual.X, visual.Y);
            return;
        }

        double right = visual.X + shape.Width;
        double bottom = visual.Y + shape.Height;
        if (right <= 0D ||
            bottom <= 0D ||
            visual.X >= surfaceWidth ||
            visual.Y >= surfaceHeight) {
            return;
        }

        double minimumLeft = Math.Min(0D, visual.X);
        double minimumTop = Math.Min(0D, visual.Y);
        double shiftX = -minimumLeft;
        double shiftY = -minimumTop;
        double nestedWidth =
            Math.Max(surfaceWidth, right) - minimumLeft;
        double nestedHeight =
            Math.Max(surfaceHeight, bottom) - minimumTop;
        var nested = new OfficeDrawing(
            Math.Max(0.01D, nestedWidth),
            Math.Max(0.01D, nestedHeight));
        nested.Fonts.AddRange(fonts);
        nested.AddShape(
            shape,
            visual.X + shiftX,
            visual.Y + shiftY);
        drawing.AddClippedDrawing(
            nested,
            0D,
            0D,
            OfficeClipPath.Rectangle(surfaceWidth, surfaceHeight),
            -shiftX,
            -shiftY);
    }

    private static void AddEffectGroup(
        OfficeDrawing drawing,
        HtmlRenderEffectGroup group,
        double surfaceWidth,
        double surfaceHeight,
        OfficeFontFaceCollection fonts,
        CancellationToken cancellationToken) {
        double nestedWidth = Math.Max(surfaceWidth, MaximumRight(group.Visuals));
        double nestedHeight = Math.Max(surfaceHeight, MaximumBottom(group.Visuals));
        var nested = new OfficeDrawing(Math.Max(0.01D, nestedWidth), Math.Max(0.01D, nestedHeight));
        nested.Fonts.AddRange(fonts);
        foreach (HtmlRenderVisual child in group.Visuals) {
            cancellationToken.ThrowIfCancellationRequested();
            AddVisual(nested, child, nested.Width, nested.Height, fonts, cancellationToken);
        }
        drawing.AddEffectDrawing(nested, group.Transform, group.Opacity);
    }

    private static void AddClipGroup(
        OfficeDrawing drawing,
        HtmlRenderClipGroup group,
        double surfaceWidth,
        double surfaceHeight,
        OfficeFontFaceCollection fonts,
        CancellationToken cancellationToken) {
        double left = group.ClipHorizontal ? Math.Max(0D, group.ClipX) : 0D;
        double top = group.ClipVertical ? Math.Max(0D, group.ClipY) : 0D;
        double right = group.ClipHorizontal ? Math.Min(surfaceWidth, group.ClipX + group.ClipWidth) : surfaceWidth;
        double bottom = group.ClipVertical ? Math.Min(surfaceHeight, group.ClipY + group.ClipHeight) : surfaceHeight;
        if (right <= left + 0.0001D || bottom <= top + 0.0001D) return;

        double minimumLeft = Math.Min(0D, MinimumLeft(group.Visuals));
        double minimumTop = Math.Min(0D, MinimumTop(group.Visuals));
        double shiftX = -minimumLeft;
        double shiftY = -minimumTop;
        double nestedWidth = Math.Max(surfaceWidth, MaximumRight(group.Visuals)) - minimumLeft;
        double nestedHeight = Math.Max(surfaceHeight, MaximumBottom(group.Visuals)) - minimumTop;
        var nested = new OfficeDrawing(Math.Max(0.01D, nestedWidth), Math.Max(0.01D, nestedHeight));
        nested.Fonts.AddRange(fonts);
        foreach (HtmlRenderVisual child in group.Visuals) {
            cancellationToken.ThrowIfCancellationRequested();
            AddVisual(nested, child.Translate(shiftX, shiftY, child.PaintOrder), nested.Width, nested.Height, fonts, cancellationToken);
        }
        drawing.AddClippedDrawing(
            nested,
            left,
            top,
            OfficeClipPath.Rectangle(right - left, bottom - top),
            -left - shiftX,
            -top - shiftY);
    }

    private static void AddPathClipGroup(
        OfficeDrawing drawing,
        HtmlRenderPathClipGroup group,
        double surfaceWidth,
        double surfaceHeight,
        OfficeFontFaceCollection fonts,
        CancellationToken cancellationToken) {
        double nestedWidth = Math.Max(surfaceWidth, MaximumRight(group.Visuals));
        double nestedHeight = Math.Max(surfaceHeight, MaximumBottom(group.Visuals));
        var nested = new OfficeDrawing(Math.Max(0.01D, nestedWidth), Math.Max(0.01D, nestedHeight));
        nested.Fonts.AddRange(fonts);
        foreach (HtmlRenderVisual child in group.Visuals) {
            cancellationToken.ThrowIfCancellationRequested();
            AddVisual(nested, child, nested.Width, nested.Height, fonts, cancellationToken);
        }
        drawing.AddClippedDrawing(nested, group.ClipX, group.ClipY, group.ClipPath, -group.ClipX, -group.ClipY);
    }

    private static double MaximumRight(IEnumerable<HtmlRenderVisual> visuals) => visuals
        .Select(visual => visual is HtmlRenderClipGroup clipGroup
            ? Math.Max(visual.X + visual.Width, MaximumRight(clipGroup.Visuals))
            : visual is HtmlRenderPathClipGroup pathClipGroup
                ? Math.Max(visual.X + visual.Width, MaximumRight(pathClipGroup.Visuals))
            : visual is HtmlRenderEffectGroup effectGroup
                ? Math.Max(visual.X + visual.Width, MaximumRight(effectGroup.Visuals))
            : visual is HtmlRenderSemanticGroup semanticGroup
                ? Math.Max(visual.X + visual.Width, MaximumRight(semanticGroup.Visuals))
            : visual is HtmlRenderLogicalTextGroup logicalTextGroup
                ? Math.Max(visual.X + visual.Width, MaximumRight(logicalTextGroup.Visuals))
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
            : visual is HtmlRenderSemanticGroup semanticGroup
                ? Math.Max(visual.Y + visual.Height, MaximumBottom(semanticGroup.Visuals))
            : visual is HtmlRenderLogicalTextGroup logicalTextGroup
                ? Math.Max(visual.Y + visual.Height, MaximumBottom(logicalTextGroup.Visuals))
                : visual.Y + visual.Height)
        .DefaultIfEmpty(0.01D)
        .Max();

    private static double MinimumLeft(IEnumerable<HtmlRenderVisual> visuals) => visuals
        .Select(visual => visual is HtmlRenderClipGroup clipGroup
            ? Math.Min(visual.X, MinimumLeft(clipGroup.Visuals))
            : visual is HtmlRenderPathClipGroup pathClipGroup
                ? Math.Min(visual.X, MinimumLeft(pathClipGroup.Visuals))
                : visual is HtmlRenderEffectGroup effectGroup
                    ? Math.Min(visual.X, MinimumLeft(effectGroup.Visuals))
                : visual is HtmlRenderSemanticGroup semanticGroup
                    ? Math.Min(visual.X, MinimumLeft(semanticGroup.Visuals))
                : visual is HtmlRenderLogicalTextGroup logicalTextGroup
                    ? Math.Min(visual.X, MinimumLeft(logicalTextGroup.Visuals))
                    : visual.X)
        .DefaultIfEmpty(0D)
        .Min();

    private static double MinimumTop(IEnumerable<HtmlRenderVisual> visuals) => visuals
        .Select(visual => visual is HtmlRenderClipGroup clipGroup
            ? Math.Min(visual.Y, MinimumTop(clipGroup.Visuals))
            : visual is HtmlRenderPathClipGroup pathClipGroup
                ? Math.Min(visual.Y, MinimumTop(pathClipGroup.Visuals))
                : visual is HtmlRenderEffectGroup effectGroup
                    ? Math.Min(visual.Y, MinimumTop(effectGroup.Visuals))
                : visual is HtmlRenderSemanticGroup semanticGroup
                    ? Math.Min(visual.Y, MinimumTop(semanticGroup.Visuals))
                : visual is HtmlRenderLogicalTextGroup logicalTextGroup
                    ? Math.Min(visual.Y, MinimumTop(logicalTextGroup.Visuals))
                    : visual.Y)
        .DefaultIfEmpty(0D)
        .Min();
}
