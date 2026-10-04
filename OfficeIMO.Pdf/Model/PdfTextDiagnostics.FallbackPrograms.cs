using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfTextDiagnostics {
    private readonly struct EmbeddedFontFallbackProgram {
        public EmbeddedFontFallbackProgram(PdfEmbeddedFontFallbackCandidate candidate) {
            FontName = candidate.FontName;
            _candidate = candidate;
            _trueTypeFont = null;
            _cffFont = null;
            _unicodeRanges = candidate.UnicodeRanges;
        }

        public EmbeddedFontFallbackProgram(
            string fontName,
            PdfTrueTypeFontProgram font,
            OfficeFontUnicodeRangeSet unicodeRanges) {
            FontName = fontName;
            _candidate = null;
            _trueTypeFont = font;
            _cffFont = null;
            _unicodeRanges = unicodeRanges;
        }

        public EmbeddedFontFallbackProgram(
            string fontName,
            PdfOpenTypeCffFontProgram font,
            OfficeFontUnicodeRangeSet unicodeRanges) {
            FontName = fontName;
            _candidate = null;
            _trueTypeFont = null;
            _cffFont = font;
            _unicodeRanges = unicodeRanges;
        }

        private readonly PdfEmbeddedFontFallbackCandidate? _candidate;
        private readonly PdfTrueTypeFontProgram? _trueTypeFont;
        private readonly PdfOpenTypeCffFontProgram? _cffFont;
        private readonly OfficeFontUnicodeRangeSet _unicodeRanges;

        public string FontName { get; }

        public bool TryGetGlyphId(int unicodeScalar, out int glyphId) {
            if (!_unicodeRanges.Contains(unicodeScalar)) {
                glyphId = 0;
                return false;
            }
            return TryGetGlyphIdIgnoringUnicodeRanges(unicodeScalar, out glyphId);
        }

        public bool TryGetLigatureGlyphId(
            string text,
            int textIndex,
            int textLength,
            int ligatureScalar,
            out int glyphId) {
            int end = textIndex + textLength;
            for (int index = textIndex; index < end;) {
                int scalar = ReadScalar(text, ref index);
                if (!_unicodeRanges.Contains(scalar)) {
                    glyphId = 0;
                    return false;
                }
            }

            return TryGetGlyphIdIgnoringUnicodeRanges(ligatureScalar, out glyphId);
        }

        public bool TryGetGlyphIdIgnoringUnicodeRanges(int unicodeScalar, out int glyphId) {
            // Priority scanning stops as soon as a font covers the scalar. Keep later
            // candidates unparsed, including large collections or unusable unused data.
            if (_candidate != null) {
                return FallbackProgramCache.GetValue(_candidate, CreateFallbackProgram).Program
                    .TryGetGlyphIdIgnoringUnicodeRanges(unicodeScalar, out glyphId);
            }

            if (_trueTypeFont != null) {
                return _trueTypeFont.TryGetGlyphId(unicodeScalar, out glyphId);
            }

            if (_cffFont != null) {
                return _cffFont.TryGetGlyphId(unicodeScalar, out glyphId);
            }

            glyphId = 0;
            return false;
        }
    }
}
