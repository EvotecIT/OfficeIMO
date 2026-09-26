using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeTrueTypeFont {
    private byte[]? _fontDataForShaping;

    internal byte[] FontDataForShaping {
        get {
            byte[]? existing = _fontDataForShaping;
            if (existing != null) return existing;
            var snapshot = (byte[])_data.Clone();
            return System.Threading.Interlocked.CompareExchange(
                ref _fontDataForShaping,
                snapshot,
                null) ?? snapshot;
        }
    }

    internal int UnitsPerEm => _unitsPerEm;

    int IOfficeFontProgram.UnitsPerEm => UnitsPerEm;

    internal bool TryGetGlyphMetrics(int scalar, out int glyphId, out int advanceWidth) {
        ushort mapped = MapGlyph(scalar);
        glyphId = mapped;
        advanceWidth = mapped == 0 ? 0 : AdvanceWidth(mapped);
        return mapped != 0;
    }

    internal int GetKerningAdjustment(
        int leftGlyphId,
        int rightGlyphId,
        int leftScalar,
        int rightScalar) {
        if ((uint)leftGlyphId > ushort.MaxValue || (uint)rightGlyphId > ushort.MaxValue) return 0;
        return Kerning((ushort)leftGlyphId, (ushort)rightGlyphId, leftScalar, rightScalar);
    }

    internal OfficeOpenTypeGlyphPositioning[] PositionGlyphRun(
        IReadOnlyList<int> glyphs,
        IReadOnlyList<int> scalars,
        System.Threading.CancellationToken cancellationToken = default) => _kerning.PositionRun(glyphs, scalars, cancellationToken);

    bool IOfficeColorFontProgram.HasColorGlyph(int glyphId) => _colorGlyphs?.HasColorGlyph(glyphId) == true;

    bool IOfficeColorFontProgram.TryGetColorLayers(
        int glyphId,
        string? palette,
        OfficeColor foreground,
        out IReadOnlyList<OfficeColorGlyphLayer> layers) {
        if (_colorGlyphs != null) return _colorGlyphs.TryGetLayers(glyphId, palette, foreground, out layers);
        layers = Array.Empty<OfficeColorGlyphLayer>();
        return false;
    }

    byte[] IOfficeFontProgram.GetFontDataForShaping() => (byte[])FontDataForShaping.Clone();

    bool IOfficeFontProgram.HasGlyphs(string text) => HasGlyphs(text);

    IReadOnlyList<double> IOfficeFontProgram.MeasureTextElements(
        IReadOnlyList<string> elements,
        double fontSize) => MeasureTextElements(elements, fontSize);

    double IOfficeFontProgram.LineSpacingRatio => LineSpacingRatio;

    bool IOfficeFontProgram.TryGetGlyphMetrics(
        int scalar,
        out int glyphId,
        out int advanceWidth) => TryGetGlyphMetrics(scalar, out glyphId, out advanceWidth);

    /// <inheritdoc />
    public bool IsOpenTypeCff => false;

    /// <inheritdoc />
    public bool ProvidesComplexTextLayout => false;

    double IOfficeFontProgram.MeasureShapedText(
        string text,
        OfficeTextShapingResult result,
        double fontSize) {
        OfficeTrueTypeVariations.WorkBudget? workBudget = _variations?.CreateWorkBudget();
        return CreateShapedTextRun(text, result, workBudget, CancellationToken.None).Measure(fontSize);
    }

    List<List<OfficePoint>> IOfficeFontProgram.GetShapedTextContours(
        string text,
        OfficeTextShapingResult result,
        double x,
        double y,
        double fontSize) {
        OfficeTrueTypeVariations.WorkBudget? workBudget = _variations?.CreateWorkBudget();
        return CreateShapedTextRun(text, result, workBudget, CancellationToken.None).GetContours(
            x,
            y,
            fontSize,
            variationWorkBudget: workBudget);
    }

    List<List<OfficePoint>> IOfficeBoundedFontProgram.GetShapedTextContoursBounded(
        string text,
        OfficeTextShapingResult result,
        double x,
        double y,
        double fontSize,
        int maximumPointCount,
        CancellationToken cancellationToken) {
        OfficeTrueTypeVariations.WorkBudget? workBudget = _variations?.CreateWorkBudget();
        return CreateShapedTextRun(text, result, workBudget, cancellationToken).GetContours(
            x,
            y,
            fontSize,
            maximumPointCount,
            cancellationToken,
            workBudget);
    }

    internal ShapedTextRun CreateShapedTextRun(
        string text,
        OfficeTextShapingResult result,
        OfficeTrueTypeVariations.WorkBudget? variationWorkBudget = null,
        CancellationToken cancellationToken = default) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        if (result == null) throw new ArgumentNullException(nameof(result));
        if (text.Length > 0 && result.Glyphs.Count == 0) {
            throw new ArgumentException(
                "Drawing text shaping provider returned no glyphs for non-empty text.",
                nameof(result));
        }
        variationWorkBudget ??= _variations?.CreateWorkBudget();

        var glyphs = new PositionedGlyph[result.Glyphs.Count];
        for (int index = 0; index < result.Glyphs.Count; index++) {
            OfficeShapedGlyph glyph = result.Glyphs[index];
            if (glyph.GlyphId <= 0 || glyph.GlyphId >= _numGlyphs) {
                throw new ArgumentException(
                    "Drawing text shaping provider returned glyph id " +
                    glyph.GlyphId.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                    ", which is outside the selected font glyph range.",
                    nameof(result));
            }
            if (glyph.TextIndex < 0 ||
                glyph.TextIndex > text.Length ||
                glyph.UnicodeText.Length > text.Length - glyph.TextIndex ||
                !string.Equals(
                    text.Substring(glyph.TextIndex, glyph.UnicodeText.Length),
                    glyph.UnicodeText,
                    StringComparison.Ordinal)) {
                throw new ArgumentException(
                    "Drawing text shaping provider returned a Unicode mapping outside the source text.",
                    nameof(result));
            }

            glyphs[index] = new PositionedGlyph(
                (ushort)glyph.GlyphId,
                checked(
                    (glyph.AdvanceWidth ?? AdvanceWidth(
                        (ushort)glyph.GlyphId,
                        variationWorkBudget,
                        cancellationToken)) +
                    result.GetAdvanceAdjustment(index)),
                glyph.AdvanceHeight ?? 0,
                glyph.OffsetX,
                glyph.OffsetY);
        }

        return new ShapedTextRun(this, glyphs, result.Direction);
    }

    internal sealed class ShapedTextRun {
        private readonly OfficeTrueTypeFont _font;
        internal readonly PositionedGlyph[] _glyphs;
        private readonly long _advanceWidth;
        private readonly long _advanceHeight;

        internal ShapedTextRun(OfficeTrueTypeFont font, PositionedGlyph[] glyphs, OfficeTextDirection direction) {
            _font = font;
            _glyphs = glyphs;
            Direction = direction;
            long width = 0L;
            long height = 0L;
            for (int index = 0; index < glyphs.Length; index++) {
                width = checked(width + glyphs[index].AdvanceWidth);
                height = checked(height + glyphs[index].AdvanceHeight);
            }
            _advanceWidth = width;
            _advanceHeight = height;
        }

        internal OfficeTextDirection Direction { get; }

        internal double Measure(double fontSize) => Math.Abs(
            (Direction == OfficeTextDirection.TopToBottom ? _advanceHeight : _advanceWidth) * _font.ScaleFor(fontSize));

        internal List<List<OfficePoint>> GetContours(
            double x,
            double y,
            double fontSize,
            int maximumPointCount = int.MaxValue,
            CancellationToken cancellationToken = default,
            OfficeTrueTypeVariations.WorkBudget? variationWorkBudget = null) {
            if (maximumPointCount <= 0) throw new ArgumentOutOfRangeException(nameof(maximumPointCount));
            var contours = new List<List<OfficePoint>>();
            double scale = _font.ScaleFor(fontSize);
            if (Direction == OfficeTextDirection.TopToBottom) {
                double cursorY = y;
                int verticalPointCount = 0;
                variationWorkBudget ??= _font._variations?.CreateWorkBudget();
                for (int index = 0; index < _glyphs.Length; index++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    PositionedGlyph glyph = _glyphs[index];
                    double glyphX = x + (glyph.OffsetX * scale);
                    double glyphBaseline = cursorY - (glyph.OffsetY * scale);
                    List<List<OfficePoint>> glyphContours = _font.ReadGlyphContours(
                        glyph.GlyphId,
                        new FontTransform(scale, 0D, 0D, -scale, glyphX, glyphBaseline),
                        0,
                        variationWorkBudget,
                        maximumPointCount,
                        ref verticalPointCount,
                        cancellationToken,
                        attachmentPoints: null);
                    contours.AddRange(glyphContours);
                    cursorY -= glyph.AdvanceHeight * scale;
                }
                return contours;
            }
            bool negativeDirection = _advanceWidth < 0L;
            double cursor = negativeDirection ? x - (_advanceWidth * scale) : x;
            double baseline = y + (_font._ascender * scale);
            int pointCount = 0;
            variationWorkBudget ??= _font._variations?.CreateWorkBudget();
            for (int index = 0; index < _glyphs.Length; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                PositionedGlyph glyph = _glyphs[index];
                if (negativeDirection) {
                    cursor += glyph.AdvanceWidth * scale;
                }
                double glyphX = cursor + (glyph.OffsetX * scale);
                double glyphBaseline = baseline - (glyph.OffsetY * scale);
                List<List<OfficePoint>> glyphContours = _font.ReadGlyphContours(
                    glyph.GlyphId,
                    new FontTransform(scale, 0D, 0D, -scale, glyphX, glyphBaseline),
                    0,
                    variationWorkBudget,
                    maximumPointCount,
                    ref pointCount,
                    cancellationToken,
                    attachmentPoints: null);
                contours.AddRange(glyphContours);
                if (!negativeDirection) {
                    cursor += glyph.AdvanceWidth * scale;
                }
            }
            return contours;
        }
    }

    internal readonly struct PositionedGlyph {
        internal PositionedGlyph(ushort glyphId, int advanceWidth, int advanceHeight, int offsetX, int offsetY) {
            GlyphId = glyphId;
            AdvanceWidth = advanceWidth;
            AdvanceHeight = advanceHeight;
            OffsetX = offsetX;
            OffsetY = offsetY;
        }

        internal ushort GlyphId { get; }
        internal int AdvanceWidth { get; }
        internal int AdvanceHeight { get; }
        internal int OffsetX { get; }
        internal int OffsetY { get; }
    }
}
