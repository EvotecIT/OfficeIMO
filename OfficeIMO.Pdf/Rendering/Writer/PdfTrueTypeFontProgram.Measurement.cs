using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal sealed partial class PdfTrueTypeFontProgram {
    private int MeasureDefaultLatinAdvanceWidth1000(string text, PdfTextShapingOptions options) {
        if (!PdfShortTextCache<PdfMeasuredText>.IsEligible(text, options))
            return PdfExternalTextShaper.MeasureDefaultLatinAdvanceWidth1000(text, this, options);
        if (_shortMeasurements.TryGet(text, options, out PdfMeasuredText cached)) {
            ReplayMeasuredGlyphUsage(cached);
            return cached.Advance;
        }
        // Text is usually measured before it is drawn. Shaping the run here, when drawing would produce the
        // same default Latin run, lets the draw and its paint inspection reuse it instead of shaping again.
        if (_shortGlyphRuns.TryGet(text, options, out PdfGlyphRun shaped)) {
            ReplayGlyphRunUsage(shaped);
            return shaped.TotalAdvanceWidth1000;
        }
        if (PdfExternalTextShaper.TryShapeCacheableDefaultLatinText(text, this, options, out shaped)) {
            _shortGlyphRuns.Add(text, options, shaped, shaped.Glyphs.Count);
            return shaped.TotalAdvanceWidth1000;
        }
        var usage = new List<(int, string)>(text.Length);
        int advance = PdfExternalTextShaper.MeasureDefaultLatinAdvanceWidth1000(text, this, options,
            (glyphId, unicode) => usage.Add((glyphId, unicode)));
        _shortMeasurements.Add(text, options, new PdfMeasuredText(advance, usage), usage.Count);
        return advance;
    }

    private void ReplayMeasuredGlyphUsage(PdfMeasuredText measured) {
        lock (_usageLock) {
            foreach (var glyph in measured.Usage) {
                if (glyph.GlyphId < 0) continue;
                RecordNormalizedGlyphUsage(glyph.GlyphId, OfficeArabicTextShaper.ToLogicalText(glyph.Unicode));
            }
        }
    }

    /// <summary>Measures canonical substituted tokens, preserving usage and failure order under one usage lock.</summary>
    internal int MeasureDefaultLatinTokens(List<OfficeOpenTypeSubstitution.GlyphToken> tokens, OfficeOpenTypeGlyphPositioning[]? positions,
        bool recordGlyphUsage, Action<int, string>? observeUsage, CancellationToken cancellationToken) {
        if (recordGlyphUsage) {
            lock (_usageLock) {
                return MeasureDefaultLatinTokensCore(tokens, positions, recordGlyphUsage, observeUsage, cancellationToken);
            }
        }
        return MeasureDefaultLatinTokensCore(tokens, positions, recordGlyphUsage, observeUsage, cancellationToken);
    }

    private int MeasureDefaultLatinTokensCore(List<OfficeOpenTypeSubstitution.GlyphToken> tokens, OfficeOpenTypeGlyphPositioning[]? positions,
        bool recordGlyphUsage, Action<int, string>? observeUsage, CancellationToken cancellationToken) {
        int advance = 0;
        for (int index = 0; index < tokens.Count; index++) {
            OfficeOpenTypeSubstitution.GlyphToken token = tokens[index];
            cancellationToken.ThrowIfCancellationRequested();
            // Kerning uses the drawing path's per-glyph rounding so wrapping matches painting.
            if (positions != null && positions[index].XAdvance != 0)
                advance = checked(advance + PdfExternalTextShaper.ScaleToPdfUnits(positions[index].XAdvance, UnitsPerEm));
            PdfExternalTextShaper.ValidateMeasuredGlyph(token.GlyphId, GlyphCount);
            int width = GetGlyphWidth1000(token.GlyphId);
            if (recordGlyphUsage) {
                RecordNormalizedGlyphUsage(token.GlyphId, OfficeArabicTextShaper.ToLogicalText(token.UnicodeText));
            }
            observeUsage?.Invoke(token.GlyphId, token.UnicodeText);
            advance = checked(advance + width);
        }
        return advance;
    }
}
