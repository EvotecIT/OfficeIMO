using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeManagedTextShapingProvider {
    private const int MaximumNormalizationInput = 1_000_000;
    private const int MaximumNormalizedClusterScalars = 256;

    // Normalization changes the font glyph sequence, never the original searchable string.
    // One glyph owns each complete original grapheme; all other glyphs in that cluster
    // carry empty Unicode continuations, including retained non-painting controls.
    private static bool TryMapTypographyTokens(OfficeTextShapingRequest request, IOfficeFontProgram font,
        OfficeOpenTypeReader reader, int[] scalars, int[] starts, byte[] levels,
        out List<OfficeOpenTypeSubstitution.GlyphToken> tokens) {
        tokens = new List<OfficeOpenTypeSubstitution.GlyphToken>();
        if (request.Text.Length > MaximumNormalizationInput || scalars.Length != starts.Length
            || starts.Length != levels.Length || reader.UnitsPerEm != font.UnitsPerEm) return false;
        var blankGlyphs = new Dictionary<int, bool>();
        int sourceStart = 0;
        foreach (string original in OfficeTextElements.Enumerate(request.Text)) {
            request.CancellationToken.ThrowIfCancellationRequested();
            if (original.Length > MaximumNormalizedClusterScalars * 2) return false;
            int sourceScalar = Array.BinarySearch(starts, sourceStart);
            if (sourceScalar < 0) return false;
            List<int>? normalized = NormalizeTypographyCluster(original, font, request);
            if (normalized == null || normalized.Count == 0 || normalized.Count > MaximumNormalizedClusterScalars) return false;
            bool first = true;
            foreach (int scalar in normalized) {
                request.CancellationToken.ThrowIfCancellationRequested();
                int glyphId;
                if (OfficeManagedTextShaper.IsNonPaintingControl(scalar)) {
                    if (!TryMapBlankControl(font, scalar, request, blankGlyphs, out glyphId)) return false;
                } else {
                    int mapped = (levels[sourceScalar] & 1) == 1 && scalar <= char.MaxValue
                        ? OfficeBidiMirroring.Mirror((char)scalar) : scalar;
                    if (!font.TryGetGlyphMetrics(mapped, out glyphId, out _)) return false;
                }
                tokens.Add(new OfficeOpenTypeSubstitution.GlyphToken(glyphId, first ? original : string.Empty,
                    sourceStart, scalar, isUnicodeContinuation: !first, clusterStart: sourceStart));
                first = false;
            }
            sourceStart = checked(sourceStart + original.Length);
        }
        return sourceStart == request.Text.Length && tokens.Count > 0;
    }

    private static List<int>? NormalizeTypographyCluster(string original, IOfficeFontProgram font,
        OfficeTextShapingRequest request) {
        string decomposed;
        try { decomposed = original.Normalize(NormalizationForm.FormD); }
        catch (ArgumentException) { return null; } // Ill-formed UTF-16 is a declined public shaping request.
        int[] nfd = OfficeUnicodeBidi.ToScalars(decomposed, out _);
        if (nfd.Length > MaximumNormalizedClusterScalars) return null;
        var sequence = new List<int>(nfd);
        CanonicalOrder(sequence, request);
        bool arabic = false;
        foreach (int scalar in sequence) {
            request.CancellationToken.ThrowIfCancellationRequested();
            arabic |= OfficeTextElements.IsJoiningScriptScalar(scalar) ||
                OfficeUnicodeTextData.GetCanonicalCombiningClass(scalar) is >= 27 and <= 35;
        }
        if (arabic) {
            ReorderArabicMarks(sequence, request);
            ComposeArabicGlyphSequence(sequence, font, request);
        } else {
            string composed;
            try { composed = original.Normalize(NormalizationForm.FormC); }
            catch (ArgumentException) { return null; }
            var nfc = new List<int>(OfficeUnicodeBidi.ToScalars(composed, out _));
            if (CanMapNormalizedScalars(nfc, font, request)) return nfc;
        }
        return CanMapNormalizedScalars(sequence, font, request) ? sequence : null;
    }

    // Use the pinned UCD classes; runtime normalization supplies canonical decomposition,
    // but no probing character or platform-specific ordering heuristic supplies CCC values.
    private static void CanonicalOrder(List<int> sequence, OfficeTextShapingRequest request) {
        for (int index = 1; index < sequence.Count; index++) {
            request.CancellationToken.ThrowIfCancellationRequested();
            byte combining = OfficeUnicodeTextData.GetCanonicalCombiningClass(sequence[index]);
            if (combining == 0) continue;
            int current = index;
            while (current > 0) {
                byte previous = OfficeUnicodeTextData.GetCanonicalCombiningClass(sequence[current - 1]);
                if (previous == 0 || previous <= combining) break;
                int swap = sequence[current - 1]; sequence[current - 1] = sequence[current]; sequence[current] = swap;
                current--;
            }
        }
    }

    // Arabic transient mark ordering follows Unicode 16.0.0 UAX #53 revision 10:
    // https://www.unicode.org/reports/tr53/tr53-10.html#Specification_of_AMTRA
    // CCC-zero characters, including CGJ, delimit each independently reordered sequence.
    private static void ReorderArabicMarks(List<int> sequence, OfficeTextShapingRequest request) {
        for (int start = 0; start < sequence.Count;) {
            request.CancellationToken.ThrowIfCancellationRequested();
            if (OfficeUnicodeTextData.GetCanonicalCombiningClass(sequence[start]) == 0) { start++; continue; }
            int end = start + 1;
            while (end < sequence.Count && OfficeUnicodeTextData.GetCanonicalCombiningClass(sequence[end]) != 0) end++;
            var below = new List<int>(); var above = new List<int>();
            var shadda = new List<int>(); var rest = new List<int>();
            bool belowPrefix = true, abovePrefix = true;
            for (int index = start; index < end; index++) {
                int scalar = sequence[index];
                byte combining = OfficeUnicodeTextData.GetCanonicalCombiningClass(scalar);
                if (combining == 33) shadda.Add(scalar);
                else if (combining == 220 && belowPrefix && OfficeUnicodeTextData.IsModifierCombiningMark(scalar)) below.Add(scalar);
                else if (combining == 230 && abovePrefix && OfficeUnicodeTextData.IsModifierCombiningMark(scalar)) above.Add(scalar);
                else {
                    rest.Add(scalar);
                    if (combining == 220) belowPrefix = false;
                    if (combining == 230) abovePrefix = false;
                }
            }
            int output = start;
            foreach (List<int> group in new[] { below, above, shadda, rest }) foreach (int scalar in group) sequence[output++] = scalar;
            start = end;
        }
    }

    // Font-directed composition follows the reference shaper's Arabic render classes:
    // a lower-class intervening mark does not block a covered canonical composite.
    // This deliberately differs from UAX #53's adjacent-only recomposition guidance;
    // it is native-compatible Arabic normalization, not complete AMTRA conformance.
    // https://github.com/harfbuzz/harfbuzz/blob/14.2.1/src/hb-ot-shape-normalize.cc
    private static void ComposeArabicGlyphSequence(List<int> sequence, IOfficeFontProgram font,
        OfficeTextShapingRequest request) {
        List<byte> renderClasses = ArabicRenderClasses(sequence, request);
        int starter = -1;
        for (int index = 0; index < sequence.Count;) {
            request.CancellationToken.ThrowIfCancellationRequested();
            if (renderClasses[index] == 0) {
                starter = OfficeManagedTextShaper.IsNonPaintingControl(sequence[index]) ? -1 : index;
                index++;
                continue;
            }
            if (starter < 0 || (starter != index - 1 && renderClasses[index - 1] >= renderClasses[index])) {
                index++;
                continue;
            }
            string pair = char.ConvertFromUtf32(sequence[starter]) + char.ConvertFromUtf32(sequence[index]);
            int[] composed = OfficeUnicodeBidi.ToScalars(pair.Normalize(NormalizationForm.FormC), out _);
            if (composed.Length != 1 || !font.TryGetGlyphMetrics(composed[0], out _, out _)) {
                index++;
                continue;
            }
            sequence[starter] = composed[0];
            sequence.RemoveAt(index);
            renderClasses.RemoveAt(index);
        }
    }

    // These are rendering-only classes. The generated canonical Unicode data remains
    // unchanged. Only leading MCM groups moved by ReorderArabicMarks receive 25/26;
    // a later MCM behind a same-class non-MCM keeps its original blocking class.
    private static List<byte> ArabicRenderClasses(List<int> sequence, OfficeTextShapingRequest request) {
        var classes = new List<byte>(sequence.Count);
        foreach (int scalar in sequence) {
            request.CancellationToken.ThrowIfCancellationRequested();
            byte combining = OfficeUnicodeTextData.GetCanonicalCombiningClass(scalar);
            classes.Add(combining is >= 27 and <= 32 ? (byte)(combining + 1) : combining == 33 ? (byte)27 : combining);
        }
        for (int start = 0; start < sequence.Count;) {
            request.CancellationToken.ThrowIfCancellationRequested();
            if (classes[start] == 0) { start++; continue; }
            int end = start + 1;
            while (end < sequence.Count && classes[end] != 0) end++;
            int index = start;
            while (index < end && classes[index] == 220 && OfficeUnicodeTextData.IsModifierCombiningMark(sequence[index]))
                classes[index++] = 25;
            while (index < end && classes[index] == 230 && OfficeUnicodeTextData.IsModifierCombiningMark(sequence[index]))
                classes[index++] = 26;
            start = end;
        }
        return classes;
    }

    private static bool CanMapNormalizedScalars(List<int> sequence, IOfficeFontProgram font,
        OfficeTextShapingRequest request) {
        foreach (int scalar in sequence) {
            request.CancellationToken.ThrowIfCancellationRequested();
            if (!OfficeManagedTextShaper.IsNonPaintingControl(scalar) && !font.TryGetGlyphMetrics(scalar, out _, out _)) return false;
        }
        return sequence.Count <= MaximumNormalizedClusterScalars;
    }

    private static bool TryMapBlankControl(IOfficeFontProgram font, int scalar, OfficeTextShapingRequest request,
        Dictionary<int, bool> blankGlyphs, out int glyphId) {
        // HarfBuzz also replaces default ignorables with the selected font's blank space.
        // A hostile font can draw its space: verify the actual bounded contours first.
        if (font.TryGetGlyphMetrics(0x20, out glyphId, out _) && IsBlankGlyph(font, glyphId, request, blankGlyphs)) return true;
        return font.TryGetGlyphMetrics(scalar, out glyphId, out _) && IsBlankGlyph(font, glyphId, request, blankGlyphs);
    }

    private static bool IsBlankGlyph(IOfficeFontProgram font, int glyphId, OfficeTextShapingRequest request,
        Dictionary<int, bool> blankGlyphs) {
        if (blankGlyphs.TryGetValue(glyphId, out bool empty)) return empty;
        if (font is not IOfficeBoundedFontProgram bounded) return false;
        request.CancellationToken.ThrowIfCancellationRequested();
        var result = new OfficeTextShapingResult(new[] { new OfficeShapedGlyph(glyphId, " ", 0,
            advanceWidth: 0, advanceHeight: 0, offsetX: 0, offsetY: 0) }, OfficeTextDirection.LeftToRight);
        try {
            List<List<OfficePoint>> contours = bounded.GetShapedTextContoursBounded(" ", result, 0D, 0D,
                font.UnitsPerEm, maximumPointCount: 1, request.CancellationToken);
            empty = true;
            foreach (List<OfficePoint> contour in contours) if (contour.Count != 0) { empty = false; break; }
        } catch (InvalidDataException) { empty = false; }
        catch (InvalidOperationException exception) when ((font is OfficeTrueTypeFont or OfficeOpenTypeCffFont)
            && exception.Message == "Font outline expansion exceeded the configured point budget.") {
            // The owned one-point proof intentionally rejects inked candidates.
            // Other provider failures and cancellation retain their normal behavior.
            empty = false;
        }
        blankGlyphs.Add(glyphId, empty);
        return empty;
    }
}
