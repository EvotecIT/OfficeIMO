using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Drawing;

/// <summary>Normalizes legacy symbol-font list markers for Unicode text renderers.</summary>
/// <remarks>
/// Word stores symbol-font markers as single-byte codes (for example <c>§</c> in Wingdings), either
/// directly or in the U+F020–U+F0FF private-use range. Word draws the glyph at that code in the symbol
/// font. Unicode renderers draw the equivalent Unicode character in the paragraph's text font instead,
/// so output never depends on a symbol font being installed and never encodes a code the font lacks.
/// </remarks>
internal static class OfficeTextListMarkerNormalizer {
    private const string FallbackBullet = "\u2022";

    // Adobe Symbol encoding, limited to the codes Word and other producers use for markers.
    private static readonly Dictionary<int, string> SymbolGlyphs = new() {
        [0x2A] = "\u2217", [0x2D] = "\u2212", [0x6F] = "\u03bf",
        [0xA7] = "\u2663", [0xA8] = "\u2666", [0xA9] = "\u2665", [0xAA] = "\u2660",
        [0xAE] = "\u2192", [0xB0] = "\u00b0", [0xB7] = "\u2022", [0xD7] = "\u22c5",
        [0xDE] = "\u21d2", [0xE0] = "\u25ca"
    };

    // Unicode equivalents of the Wingdings glyphs used as bullets and check marks.
    private static readonly Dictionary<int, string> WingdingsGlyphs = new() {
        [0x46] = "\u261e", [0x4A] = "\u263a", [0x4C] = "\u2639",
        [0x6C] = "\u25cf", [0x6D] = "\u274d", [0x6E] = "\u25a0", [0x6F] = "\u25a1",
        [0x70] = "\u25fb", [0x71] = "\u2751", [0x72] = "\u2752", [0x73] = "\u2b27",
        [0x74] = "\u29eb", [0x75] = "\u25c6", [0x76] = "\u2756", [0x77] = "\u2b25",
        [0x78] = "\u2327", [0x7A] = "\u2318",
        [0x9E] = "\u00b7", [0x9F] = "\u2022", [0xA1] = "\u25cb", [0xA4] = "\u25c9",
        [0xA7] = "\u25aa", [0xA8] = "\u25fb", [0xAB] = "\u2605",
        [0xD8] = "\u27a2", [0xE8] = "\u2794", [0xF0] = "\u21e8",
        [0xFB] = "\u2717", [0xFC] = "\u2714", [0xFD] = "\u2612", [0xFE] = "\u2611"
    };

    internal static (string Marker, bool UseTextFont) Normalize(string marker, string? fontFamily) {
        string trimmed = marker.Trim();
        if (trimmed.Length == 0) return (string.Empty, false);
        Dictionary<int, string>? glyphs = ResolveSymbolGlyphs(fontFamily);
        bool symbolFont = glyphs != null || IsUnmappedSymbolFont(fontFamily);
        if (!symbolFont) {
            // A private-use code outside a known symbol font still means a symbol glyph.
            return trimmed == "\uf0b7" ? ("\u2022", true) : (marker, false);
        }

        var builder = new StringBuilder(trimmed.Length);
        foreach (char character in trimmed) {
            int code = character >= 0xF020 && character <= 0xF0FF ? character - 0xF000 : character;
            if (code > 0xFF) {
                // Characters beyond the single-byte symbol range are already Unicode.
                builder.Append(character);
            } else if (glyphs != null && glyphs.TryGetValue(code, out string? glyph)) {
                builder.Append(glyph);
            } else if (trimmed.Length == 1) {
                // A single symbol-font code is a bullet glyph with no portable
                // meaning without the font; a bullet keeps the list readable.
                builder.Append(FallbackBullet);
            } else {
                // Numbered markers keep their digits and punctuation, which
                // symbol fonts place at the ASCII codes.
                builder.Append(code < 0x80 ? (char)code : character);
            }
        }
        return (builder.ToString(), true);
    }

    private static Dictionary<int, string>? ResolveSymbolGlyphs(string? fontFamily) {
        if (string.Equals(fontFamily, "Symbol", StringComparison.OrdinalIgnoreCase)) return SymbolGlyphs;
        if (string.Equals(fontFamily, "Wingdings", StringComparison.OrdinalIgnoreCase)) return WingdingsGlyphs;
        return null;
    }

    private static bool IsUnmappedSymbolFont(string? fontFamily) =>
        string.Equals(fontFamily, "Wingdings 2", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(fontFamily, "Wingdings 3", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(fontFamily, "Webdings", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(fontFamily, "Marlett", StringComparison.OrdinalIgnoreCase);
}
