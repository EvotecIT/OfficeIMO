using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

internal sealed class OfficeColorGlyphContours {
    internal OfficeColorGlyphContours(List<List<OfficePoint>> contours, OfficeColor color) {
        Contours = contours;
        Color = color;
    }

    internal List<List<OfficePoint>> Contours { get; }
    internal OfficeColor Color { get; }
}

public sealed partial class OfficeTrueTypeFont {
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
        if (_colorGlyphs == null || string.IsNullOrEmpty(text)) return false;

        var glyphs = new List<(ushort Glyph, int Scalar)>();
        for (int textIndex = 0; textIndex < text.Length;) {
            cancellationToken.ThrowIfCancellationRequested();
            int glyph = ReadMappedGlyph(text, ref textIndex, out int scalar);
            if (glyph > 0) glyphs.Add((checked((ushort)glyph), scalar));
        }
        if (glyphs.Count == 0 || !glyphs.Exists(item => _colorGlyphs.HasColorGlyph(item.Glyph))) return false;

        var glyphIds = new List<int>(glyphs.Count);
        var scalars = new List<int>(glyphs.Count);
        foreach ((ushort glyph, int scalar) in glyphs) {
            glyphIds.Add(glyph);
            scalars.Add(scalar);
        }
        OfficeOpenTypeGlyphPositioning[] positioning = _kerning.PositionRun(glyphIds, scalars);
        var positioned = new PositionedGlyph[glyphs.Count];
        OfficeTrueTypeVariations.WorkBudget? variationBudget = _variations?.CreateWorkBudget();
        for (int index = 0; index < glyphs.Count; index++) {
            positioned[index] = new PositionedGlyph(
                glyphs[index].Glyph,
                checked(AdvanceWidth(glyphs[index].Glyph, variationBudget, cancellationToken) + positioning[index].XAdvance),
                positioning[index].XPlacement,
                0);
        }
        return TryGetPositionedColorContours(positioned, x, y, fontSize, palette, foreground, maximumPointCount, cancellationToken, out paintedLayers);
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
        if (_colorGlyphs == null) return false;
        ShapedTextRun run = CreateShapedTextRun(text, shapingResult, cancellationToken: cancellationToken);
        if (!Array.Exists(run._glyphs, glyph => _colorGlyphs.HasColorGlyph(glyph.GlyphId))) return false;
        return TryGetPositionedColorContours(run._glyphs, x, y, fontSize, palette, foreground, maximumPointCount, cancellationToken, out paintedLayers);
    }

    private bool TryGetPositionedColorContours(
        PositionedGlyph[] glyphs,
        double x,
        double y,
        double fontSize,
        string? palette,
        OfficeColor foreground,
        int maximumPointCount,
        CancellationToken cancellationToken,
        out List<OfficeColorGlyphContours> paintedLayers) {
        paintedLayers = new List<OfficeColorGlyphContours>();
        if (_colorGlyphs == null || glyphs.Length == 0) return false;
        if (maximumPointCount <= 0) throw new ArgumentOutOfRangeException(nameof(maximumPointCount));

        double scale = ScaleFor(fontSize);
        long totalAdvance = 0;
        foreach (PositionedGlyph glyph in glyphs) totalAdvance = checked(totalAdvance + glyph.AdvanceWidth);
        bool negativeDirection = totalAdvance < 0;
        double cursor = negativeDirection ? x - totalAdvance * scale : x;
        double baseline = y + _ascender * scale;
        int pointCount = 0;
        OfficeTrueTypeVariations.WorkBudget? variationBudget = _variations?.CreateWorkBudget();
        for (int index = 0; index < glyphs.Length; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            PositionedGlyph glyph = glyphs[index];
            if (negativeDirection) cursor += glyph.AdvanceWidth * scale;
            double glyphX = cursor + glyph.OffsetX * scale;
            double glyphBaseline = baseline - glyph.OffsetY * scale;
            IReadOnlyList<OfficeColorGlyphLayer> layers;
            if (!_colorGlyphs.TryGetLayers(glyph.GlyphId, palette, foreground, out layers)) {
                layers = new[] { new OfficeColorGlyphLayer(glyph.GlyphId, foreground) };
            }
            foreach (OfficeColorGlyphLayer layer in layers) {
                cancellationToken.ThrowIfCancellationRequested();
                List<List<OfficePoint>> contours = ReadGlyphContours(
                    checked((ushort)layer.GlyphId),
                    new FontTransform(scale, 0D, 0D, -scale, glyphX, glyphBaseline),
                    0,
                    variationBudget,
                    maximumPointCount,
                    ref pointCount,
                    cancellationToken,
                    attachmentPoints: null);
                if (contours.Count > 0 && layer.Color.A > 0) paintedLayers.Add(new OfficeColorGlyphContours(contours, layer.Color));
            }
            if (!negativeDirection) cursor += glyph.AdvanceWidth * scale;
        }
        return paintedLayers.Count > 0;
    }
}
