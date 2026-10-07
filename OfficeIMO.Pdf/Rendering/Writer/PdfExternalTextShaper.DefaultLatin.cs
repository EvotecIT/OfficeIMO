using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfExternalTextShaper {
    // Default Latin shaping has nominal advances and no positioning. Project the shared
    // substituted tokens directly, retaining each source index and continuation cluster.
    private static bool TryShapeDefaultLatinText(string text, string fontName, byte[] fontData,
        bool isCff, int unitsPerEm, int glyphCount, Func<int, int> getGlyphWidth1000,
        Action<int, string>? recordGlyphUsage, PdfTextShapingOptions options, out PdfGlyphRun glyphRun) {
        var request = new OfficeTextShapingRequest(text, fontName, fontData, isCff, unitsPerEm,
            options.Direction, options.Language, default, fontCollectionIndex: null, variationCoordinates: null,
            cloneFontData: false, applyDefaultLatinLigatures: true);
        if (!OfficeManagedTextShapingProvider.Instance.TryShapeDefaultLatinTokens(request, out var tokens, out OfficeTextDirection direction)) {
            glyphRun = null!;
            return false;
        }

        var glyphs = new List<PdfGlyphInfo>(tokens.Count);
        foreach (OfficeOpenTypeSubstitution.GlyphToken token in tokens) {
            ValidateMeasuredGlyph(token.GlyphId, glyphCount);
            if (token.TextIndex < 0 || token.TextIndex > text.Length) {
                throw new ArgumentException("PDF text shaping provider returned a text index outside the source text.", nameof(text));
            }
            int width = getGlyphWidth1000(token.GlyphId);
            recordGlyphUsage?.Invoke(token.GlyphId, token.UnicodeText);
            glyphs.Add(new PdfGlyphInfo(token.GlyphId, token.UnicodeText, token.TextIndex,
                width, width, 0, 0, 0, token.ClusterStart));
        }

        bool includeActualText = direction != OfficeTextDirection.LeftToRight || OfficeManagedTextShaper.RequiresComplexLayout(text);
        glyphRun = new PdfGlyphRun(glyphs, Array.Empty<PdfTextEncodingDiagnostic>(), includeActualText ? text : null,
            direction, preserveGlyphUnicode: !includeActualText, isAutomaticallyShaped: true);
        options.ProviderShapedTextRecorder?.Invoke(text, fontName, isCff, true);
        return true;
    }
}
