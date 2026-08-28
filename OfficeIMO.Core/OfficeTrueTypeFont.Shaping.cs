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
                glyph.AdvanceWidth ?? AdvanceWidth(
                    (ushort)glyph.GlyphId,
                    variationWorkBudget,
                    cancellationToken),
                glyph.OffsetX,
                glyph.OffsetY);
        }

        return new ShapedTextRun(this, glyphs);
    }

    internal sealed class ShapedTextRun {
        private readonly OfficeTrueTypeFont _font;
        private readonly PositionedGlyph[] _glyphs;
        private readonly long _advanceWidth;

        internal ShapedTextRun(OfficeTrueTypeFont font, PositionedGlyph[] glyphs) {
            _font = font;
            _glyphs = glyphs;
            long width = 0L;
            for (int index = 0; index < glyphs.Length; index++) {
                width = checked(width + glyphs[index].AdvanceWidth);
            }
            _advanceWidth = width;
        }

        internal double Measure(double fontSize) => Math.Abs(_advanceWidth * _font.ScaleFor(fontSize));

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
        internal PositionedGlyph(ushort glyphId, int advanceWidth, int offsetX, int offsetY) {
            GlyphId = glyphId;
            AdvanceWidth = advanceWidth;
            OffsetX = offsetX;
            OffsetY = offsetY;
        }

        internal ushort GlyphId { get; }
        internal int AdvanceWidth { get; }
        internal int OffsetX { get; }
        internal int OffsetY { get; }
    }
}
