namespace OfficeIMO.Drawing;

internal sealed class OfficeColorGlyphContours {
    internal OfficeColorGlyphContours(List<List<OfficePoint>> contours, OfficeColor color) {
        Contours = contours;
        Color = color;
    }

    /// <summary>A CBDT or sbix PNG glyph placed in output coordinates, y down; its only contour is its box.</summary>
    internal OfficeColorGlyphContours(OfficeBitmapGlyphImage bitmap, double left, double top, double width, double height) {
        // The box as one contour, so the transforms applied to glyph contours also move the image.
        Contours = new List<List<OfficePoint>> {
            new List<OfficePoint> { new(left, top), new(left + width, top), new(left + width, top + height), new(left, top + height) }
        };
        Color = OfficeColor.FromRgba(0, 0, 0, 0);
        Bitmap = bitmap;
        BitmapLeft = left; BitmapTop = top; BitmapWidth = width; BitmapHeight = height;
    }

    internal List<List<OfficePoint>> Contours { get; }
    internal OfficeColor Color { get; }
    internal OfficeBitmapGlyphImage? Bitmap { get; }
    internal double BitmapLeft { get; }
    internal double BitmapTop { get; }
    internal double BitmapWidth { get; }
    internal double BitmapHeight { get; }
}

public sealed partial class OfficeTrueTypeFont {
    /// <summary>Whether this program has palette layers that shaped glyph substitution can select.</summary>
    internal bool HasColorGlyphs => _colorGlyphs != null;

    private readonly bool _hasUnreadColorTable;

    /// <summary>Whether the font has CBDT, sbix, SVG or unparsed COLR data.</summary>
    internal bool HasUnreadColorTable => _hasUnreadColorTable;
    private readonly OfficeOpenTypeBitmapGlyphs? _bitmapGlyphs;

    /// <summary>Whether this program has CBDT or sbix PNG glyphs.</summary>
    internal bool HasBitmapGlyphs => _bitmapGlyphs != null;

    /// <summary>Returns the PNG artwork of a CBDT or sbix glyph in font units, or false when it has none.</summary>
    internal bool TryGetBitmapGlyph(int glyphId, out OfficeBitmapGlyphImage? image) {
        image = null;
        return _bitmapGlyphs != null && _bitmapGlyphs.TryGetImage(glyphId, out image);
    }

    /// <summary>Whether a glyph is painted from COLR layers or a readable PNG rather than its plain outline.</summary>
    internal bool IsPaintedColorGlyph(int glyphId) =>
        _colorGlyphs?.HasColorGlyph(glyphId) == true || _bitmapGlyphs?.TryGetImage(glyphId, out _) == true;

    /// <summary>Whether a glyph has a readable CBDT or sbix PNG and no COLR layers.</summary>
    internal bool IsBitmapColorGlyph(int glyphId) =>
        _colorGlyphs?.HasColorGlyph(glyphId) != true && _bitmapGlyphs?.TryGetImage(glyphId, out _) == true;

    private bool HasPaintedColorGlyphs => _colorGlyphs != null || _bitmapGlyphs != null;

    /// <summary>
    /// Whether a mapped glyph depends on a bitmap, SVG or unreadable color table: its outline is
    /// empty, so painting it as an ordinary glyph would leave a silent blank.
    /// </summary>
    internal bool ContainsUnpaintableColorGlyph(string text) {
        if (!_hasUnreadColorTable) return false;
        for (int index = 0; index < text.Length;) {
            int glyph = ReadMappedGlyph(text, ref index, out int scalar);
            if (!char.IsWhiteSpace(char.ConvertFromUtf32(scalar), 0) && IsUnpaintableColorGlyph(glyph)) return true;
        }
        return false;
    }

    /// <summary>Whether a glyph, such as a shaped ligature, has only SVG or unreadable color artwork.</summary>
    internal bool IsUnpaintableColorGlyph(int glyph) =>
        _hasUnreadColorTable && glyph > 0 && glyph < _numGlyphs && !IsPaintedColorGlyph(glyph)
        && HasEmptyOutline((ushort)glyph);

    // The (xMin, yMin) of a glyph's outline bounding box, or null when it has no outline.
    private (int XMin, int YMin)? GetOutlineOrigin(int glyph) {
        if (glyph <= 0 || glyph >= _numGlyphs || HasEmptyOutline((ushort)glyph)) return null;
        int header = _glyf + GlyphOffset((ushort)glyph);
        return (unchecked((short)ReadUInt16(_data, header + 2)), unchecked((short)ReadUInt16(_data, header + 4)));
    }

    private bool HasEmptyOutline(ushort glyph) {
        int start = GlyphOffset(glyph);
        return start == GlyphOffset((ushort)(glyph + 1)) || ReadUInt16(_data, _glyf + start) == 0;
    }

    internal bool TryGetColorTextContours(
        string text,
        double x,
        double y,
        double fontSize,
        string? palette,
        OfficeColor foreground,
        int maximumPointCount,
        CancellationToken cancellationToken,
        out List<OfficeColorGlyphContours> paintedLayers) {
        paintedLayers = new List<OfficeColorGlyphContours>();
        if (!HasPaintedColorGlyphs || string.IsNullOrEmpty(text)) return false;

        var glyphs = new List<(ushort Glyph, int Scalar)>();
        var indexes = new List<int>();
        for (int textIndex = 0; textIndex < text.Length;) {
            cancellationToken.ThrowIfCancellationRequested();
            int sourceIndex = textIndex;
            int glyph = ReadMappedGlyph(text, ref textIndex, out int scalar);
            if (glyph > 0 || glyph == 0 && (HasPaintedNotdefMapping(scalar) ||
                OfficeMissingSpaceMetrics.IsEmSpace(scalar) && MapGlyph(scalar) == 0)) {
                indexes.Add(sourceIndex);
                glyphs.Add((checked((ushort)glyph), scalar));
            }
        }
        if (glyphs.Count == 0 || !glyphs.Exists(item => IsPaintedColorGlyph(item.Glyph))) return false;

        var glyphIds = new List<int>(glyphs.Count);
        var scalars = new List<int>(glyphs.Count);
        foreach ((ushort glyph, int scalar) in glyphs) {
            glyphIds.Add(glyph);
            scalars.Add(scalar);
        }
        bool[]? missingSpaces = FindMissingEmSpaces(glyphIds, scalars);
        OfficeOpenTypeGlyphPositioning[] positioning = _kerning.PositionRun(glyphIds, scalars, synthesizedSpaces: missingSpaces);
        bool[]? boundaries = TrackingBoundaries(text, indexes);
        var positioned = new PositionedGlyph[glyphs.Count];
        OfficeTrueTypeVariations.WorkBudget? variationBudget = _variations?.CreateWorkBudget();
        for (int index = 0; index < glyphs.Count; index++) {
            positioned[index] = new PositionedGlyph(
                glyphs[index].Glyph,
                checked(TextAdvanceWidth(glyphs[index].Glyph, OfficeMissingSpaceMetrics.IsSynthesized(missingSpaces, index),
                    variationBudget, cancellationToken) + positioning[index].XAdvance),
                0,
                positioning[index].XPlacement,
                0);
            positioned[index].TrackingBoundary = boundaries == null || boundaries[index];
        }
        return TryGetPositionedColorContours(positioned, OfficeTextDirection.LeftToRight, x, y, fontSize, palette, foreground,
            maximumPointCount, cancellationToken, out paintedLayers, missingSpaces);
    }

    internal bool TryGetShapedColorTextContours(
        string text,
        OfficeTextShapingResult shapingResult,
        double x,
        double y,
        double fontSize,
        string? palette,
        OfficeColor foreground,
        int maximumPointCount,
        CancellationToken cancellationToken,
        out List<OfficeColorGlyphContours> paintedLayers) {
        paintedLayers = new List<OfficeColorGlyphContours>();
        if (!HasPaintedColorGlyphs) return false;
        ShapedTextRun run = CreateShapedTextRun(text, shapingResult, cancellationToken: cancellationToken);
        if (!Array.Exists(run._glyphs, glyph => IsPaintedColorGlyph(glyph.GlyphId))) return false;
        return TryGetPositionedColorContours(run._glyphs, run.Direction, x, y, fontSize, palette, foreground, maximumPointCount, cancellationToken, out paintedLayers);
    }

    private bool TryGetPositionedColorContours(
        PositionedGlyph[] glyphs,
        OfficeTextDirection direction,
        double x,
        double y,
        double fontSize,
        string? palette,
        OfficeColor foreground,
        int maximumPointCount,
        CancellationToken cancellationToken,
        out List<OfficeColorGlyphContours> paintedLayers,
        bool[]? synthesizedSpaces = null) {
        paintedLayers = new List<OfficeColorGlyphContours>();
        if (!HasPaintedColorGlyphs || glyphs.Length == 0) return false;
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maximumPointCount);

        double scale = ScaleFor(fontSize);
        if (direction == OfficeTextDirection.TopToBottom) {
            double verticalCursor = y;
            int verticalPointCount = 0;
            OfficeTrueTypeVariations.WorkBudget? verticalVariationBudget = _variations?.CreateWorkBudget();
            for (int index = 0; index < glyphs.Length; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                PositionedGlyph glyph = glyphs[index];
                double glyphX = x + glyph.OffsetX * scale;
                double glyphBaseline = verticalCursor - glyph.OffsetY * scale;
                IReadOnlyList<OfficeColorGlyphLayer> layers;
                if (OfficeMissingSpaceMetrics.IsSynthesized(synthesizedSpaces, index)) {
                    layers = Array.Empty<OfficeColorGlyphLayer>();
                } else if (!TryGetColorLayers(glyph.GlyphId, palette, foreground, out layers)) {
                    layers = AddBitmapGlyph(paintedLayers, glyph.GlyphId, scale, glyphX, glyphBaseline)
                        ? Array.Empty<OfficeColorGlyphLayer>()
                        : new[] { new OfficeColorGlyphLayer(glyph.GlyphId, foreground) };
                }
                foreach (OfficeColorGlyphLayer layer in layers) {
                    cancellationToken.ThrowIfCancellationRequested();
                    List<List<OfficePoint>> contours = ReadGlyphContours(
                        checked((ushort)layer.GlyphId),
                        LayerPlacement(layer, scale, glyphX, glyphBaseline),
                        0,
                        verticalVariationBudget,
                        maximumPointCount,
                        ref verticalPointCount,
                        cancellationToken,
                        attachmentPoints: null);
                    if (contours.Count > 0 && layer.Color.A > 0) paintedLayers.Add(new OfficeColorGlyphContours(contours, layer.Color));
                }
                verticalCursor -= glyph.AdvanceHeight * scale;
            }
            return paintedLayers.Count > 0;
        }
        long totalAdvance = 0;
        int trackingCount = 0;
        foreach (PositionedGlyph glyph in glyphs) {
            totalAdvance = checked(totalAdvance + glyph.AdvanceWidth);
            if (glyph.TrackingBoundary) trackingCount++;
        }
        bool negativeDirection = totalAdvance < 0;
        double tracking = (negativeDirection ? -1D : 1D) * HorizontalTracking(fontSize);
        double cursor = negativeDirection ? x - (totalAdvance * scale + tracking * trackingCount) : x;
        double baseline = y + _ascender * scale;
        int pointCount = 0;
        OfficeTrueTypeVariations.WorkBudget? variationBudget = _variations?.CreateWorkBudget();
        for (int index = 0; index < glyphs.Length; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            PositionedGlyph glyph = glyphs[index];
            if (negativeDirection) cursor += glyph.AdvanceWidth * scale + (glyph.TrackingBoundary ? tracking : 0D);
            double glyphX = cursor + glyph.OffsetX * scale;
            double glyphBaseline = baseline - glyph.OffsetY * scale;
            IReadOnlyList<OfficeColorGlyphLayer> layers;
            if (OfficeMissingSpaceMetrics.IsSynthesized(synthesizedSpaces, index)) {
                layers = Array.Empty<OfficeColorGlyphLayer>();
            } else if (!TryGetColorLayers(glyph.GlyphId, palette, foreground, out layers)) {
                layers = AddBitmapGlyph(paintedLayers, glyph.GlyphId, scale, glyphX, glyphBaseline)
                    ? Array.Empty<OfficeColorGlyphLayer>()
                    : new[] { new OfficeColorGlyphLayer(glyph.GlyphId, foreground) };
            }
            foreach (OfficeColorGlyphLayer layer in layers) {
                cancellationToken.ThrowIfCancellationRequested();
                List<List<OfficePoint>> contours = ReadGlyphContours(
                    checked((ushort)layer.GlyphId),
                    LayerPlacement(layer, scale, glyphX, glyphBaseline),
                    0,
                    variationBudget,
                    maximumPointCount,
                    ref pointCount,
                    cancellationToken,
                    attachmentPoints: null);
                if (contours.Count > 0 && layer.Color.A > 0) paintedLayers.Add(new OfficeColorGlyphContours(contours, layer.Color));
            }
            if (!negativeDirection) cursor += glyph.AdvanceWidth * scale + (glyph.TrackingBoundary ? tracking : 0D);
        }
        return paintedLayers.Count > 0;
    }

    /// <summary>
    /// Returns a glyph's color layers as outlines in font units, y up, with the glyph origin at 0,0.
    /// Layers that use the text color are flagged so a consumer can substitute its own fill; their color carries the paint alpha.
    /// </summary>
    internal bool TryGetGlyphColorLayers(int glyphId, OfficeColor foreground, int maximumPointCount,
        CancellationToken cancellationToken, out List<(List<List<OfficePoint>> Contours, OfficeColor Color, bool Foreground)> layers) {
        layers = new List<(List<List<OfficePoint>> Contours, OfficeColor Color, bool Foreground)>();
        if (_colorGlyphs == null || glyphId <= 0 || glyphId >= _numGlyphs
            || !_colorGlyphs.TryGetLayers(glyphId, null, foreground, out IReadOnlyList<OfficeColorGlyphLayer> resolved)) return false;
        int pointCount = 0;
        OfficeTrueTypeVariations.WorkBudget? variationBudget = _variations?.CreateWorkBudget();
        foreach (OfficeColorGlyphLayer layer in resolved) {
            cancellationToken.ThrowIfCancellationRequested();
            if (layer.Color.A == 0) continue;
            var placement = new FontTransform(1D, 0D, 0D, 1D, 0D, 0D);
            if (layer.Transform is LayerTransform transform)
                placement = placement.Compose(transform.Xx, transform.Xy, transform.Yx, transform.Yy, transform.Dx, transform.Dy);
            List<List<OfficePoint>> contours = ReadGlyphContours(checked((ushort)layer.GlyphId), placement, 0, variationBudget,
                maximumPointCount, ref pointCount, cancellationToken, attachmentPoints: null);
            if (contours.Count > 0) layers.Add((contours, layer.Color, layer.IsForeground));
        }
        // A glyph whose layers are all transparent is still COLR artwork: it paints nothing, not its base outline.
        return true;
    }

    /// <summary>Returns a glyph's plain outline in font units, y up, with the glyph origin at 0,0.</summary>
    internal bool TryGetGlyphOutline(int glyphId, int maximumPointCount, CancellationToken cancellationToken,
        out List<List<OfficePoint>> contours) {
        contours = new List<List<OfficePoint>>();
        if (glyphId <= 0 || glyphId >= _numGlyphs) return false;
        int pointCount = 0;
        contours = ReadGlyphContours(checked((ushort)glyphId), new FontTransform(1D, 0D, 0D, 1D, 0D, 0D), 0,
            _variations?.CreateWorkBudget(), maximumPointCount, ref pointCount, cancellationToken, attachmentPoints: null);
        return contours.Count > 0;
    }

    private bool TryGetColorLayers(int glyphId, string? palette, OfficeColor foreground, out IReadOnlyList<OfficeColorGlyphLayer> layers) {
        layers = Array.Empty<OfficeColorGlyphLayer>();
        return _colorGlyphs != null && _colorGlyphs.TryGetLayers(glyphId, palette, foreground, out layers);
    }

    // Output coordinates are y down with the baseline at glyphBaseline; the PNG box is in font units, y up.
    private bool AddBitmapGlyph(List<OfficeColorGlyphContours> paintedLayers, int glyphId, double scale, double glyphX, double glyphBaseline) {
        if (!TryGetBitmapGlyph(glyphId, out OfficeBitmapGlyphImage? image) || image == null) return false;
        paintedLayers.Add(new OfficeColorGlyphContours(image, glyphX + image.Left * scale,
            glyphBaseline - (image.Bottom + image.Height) * scale, image.Width * scale, image.Height * scale));
        return true;
    }

    private static FontTransform LayerPlacement(OfficeColorGlyphLayer layer, double scale, double glyphX, double glyphBaseline) {
        var placement = new FontTransform(scale, 0D, 0D, -scale, glyphX, glyphBaseline);
        return layer.Transform is LayerTransform transform
            ? placement.Compose(transform.Xx, transform.Xy, transform.Yx, transform.Yy, transform.Dx, transform.Dy)
            : placement;
    }
}
