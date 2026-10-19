using OfficeIMO.Drawing;
using System.Threading;

namespace OfficeIMO.Pdf;

internal sealed partial class PdfTrueTypeFontProgram {
    private int MeasureDefaultLatinAdvanceWidth1000(string text, PdfTextShapingOptions options) {
        if (!PdfShortTextCache<PdfMeasuredText>.IsEligible(text, options))
            return PdfExternalTextShaper.MeasureDefaultLatinAdvanceWidth1000(text, this, options);
        if (_shortMeasurements.TryGet(text, options, out PdfMeasuredText cached)) {
            ReplayMeasuredGlyphUsage(cached);
            return cached.Advance;
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
