using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Splits painted runs whose text cannot name their glyphs into one visual span per glyph.
/// </summary>
/// <remarks>
/// A PDF run paints shaped glyphs whose ToUnicode text can be a multi-letter cluster (a Devanagari
/// conjunct, a Latin or Arabic ligature), nothing at all, or a letter that no single-glyph cmap entry
/// can name. Complex-script runs and runs containing such glyphs are split. Drawing the run as
/// text would re-map those letters to different glyphs or none. Each glyph is projected at its own
/// origin so the drawing can name that exact glyph. Only runs painted with an embedded font program
/// are split. Visual projection only.
/// </remarks>
internal static class PdfPaintedGlyphRuns {
    /// <summary>
    /// Stands in for a painted glyph whose ToUnicode text is empty or U+0000 (a noncharacter reserved
    /// for internal use). It keeps the glyph and its geometry in visual spans so the drawing can name
    /// that exact glyph; placeholders that cannot be resolved are removed before drawing.
    /// </summary>
    internal const char UndecodedGlyph = '\uFDD0';

    internal static readonly string UndecodedGlyphText = UndecodedGlyph.ToString();

    internal static void SplitComplexRuns(List<PdfTextSpan> spans, System.Threading.CancellationToken cancellationToken = default) {
        for (int index = spans.Count - 1; index >= 0; index--) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfTextSpan span = spans[index];
            // Only an embedded program can draw each exact glyph; substituted fonts keep the whole run
            // so a substitute shaper still sees its joining and cluster context.
            if (span.Text.Length < 2 || span.DrawingFontFamily == null || !HasGlyphGeometry(span) ||
                !OfficeManagedTextShaper.RequiresComplexLayout(span.Text) && span.Text.IndexOf(UndecodedGlyph) < 0 &&
                !HasMultiCharacterGlyph(span) ||
                !PdfTextAdvanceProjection.TryGetResolvedDirection(span, cancellationToken, out double direction)) continue;
            double radians = span.RotationDegrees * Math.PI / 180D;
            double alongX = Math.Cos(radians);
            double alongY = Math.Sin(radians);
            var glyphs = new List<PdfTextSpan>(span.GlyphCharacterLengths!.Count);
            int characterOffset = 0;
            double offset = 0D;
            for (int glyph = 0; glyph < span.GlyphCharacterLengths.Count; glyph++) {
                glyphs.Add(span.WithPaintedGlyph(glyph, characterOffset, span.X + alongX * offset, span.Y + alongY * offset));
                int end = characterOffset + span.GlyphCharacterLengths[glyph];
                for (; characterOffset < end; characterOffset++) offset += span.CharacterAdvances![characterOffset] * direction;
            }
            spans.RemoveAt(index);
            spans.InsertRange(index, glyphs);
        }
    }

    // A ligature glyph decodes to several letters; drawing those letters would replace the glyph.
    private static bool HasMultiCharacterGlyph(PdfTextSpan span) {
        foreach (int length in span.GlyphCharacterLengths!) {
            if (length > 1) return true;
        }
        return false;
    }

    private static bool HasGlyphGeometry(PdfTextSpan span) {
        IReadOnlyList<int>? lengths = span.GlyphCharacterLengths;
        IReadOnlyList<byte[]>? bytes = span.GlyphBytes;
        IReadOnlyList<double>? painted = span.GlyphPaintedAdvances;
        IReadOnlyList<double>? advances = span.CharacterAdvances;
        if (lengths == null || bytes == null || painted == null || advances == null || lengths.Count < 2 ||
            bytes.Count != lengths.Count || painted.Count != lengths.Count || advances.Count != span.Text.Length) return false;
        int total = 0;
        for (int index = 0; index < lengths.Count; index++) {
            if (lengths[index] <= 0 || !(painted[index] >= 0D) || double.IsInfinity(painted[index])) return false;
            total += lengths[index];
        }
        return total == span.Text.Length;
    }
}
