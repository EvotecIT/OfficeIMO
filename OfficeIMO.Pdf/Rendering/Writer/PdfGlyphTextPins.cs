using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Chooses composite-font character codes whose ToUnicode entries carry the source text.</summary>
internal interface IPdfGlyphTextPinner {
    /// <summary>
    /// Returns true with the character code to paint when its final ToUnicode entry will be exactly
    /// <paramref name="unicodeText"/>. Otherwise the caller must describe the glyph through ActualText.
    /// </summary>
    bool TryResolveCode(int glyphId, string unicodeText, out int code);
}

/// <summary>
/// Deterministic per-document ToUnicode ownership for a composite font whose CIDs are glyph ids.
/// </summary>
/// <remarks>
/// A glyph reachable through the font cmap maps to its lowest printable scalar (logical Arabic,
/// decomposed Latin ligature), independent of usage order. Other glyphs (ligatures, contextual
/// forms) keep the first emitted text. When fonts map several characters to one glyph, every other
/// character can use an alias code (glyph count + scalar) whose CIDToGIDMap entry selects that
/// nominal glyph. Codes are a pure function of the character, so independent document contexts
/// that share one font resource never disagree.
/// </remarks>
internal sealed class PdfGlyphTextPins : IPdfGlyphTextPinner {
    private readonly PdfCanonicalGlyphText _canonical;
    private readonly int _glyphCount;
    private readonly bool _allowAliases;
    private readonly Dictionary<int, string> _pins = new();
    private readonly SortedDictionary<int, int> _aliases = new();
    private readonly object _lock = new();

    internal PdfGlyphTextPins(PdfCanonicalGlyphText canonical, int glyphCount, bool allowAliases) {
        _canonical = canonical ?? throw new ArgumentNullException(nameof(canonical));
        _glyphCount = glyphCount;
        _allowAliases = allowAliases;
    }

    public bool TryResolveCode(int glyphId, string unicodeText, out int code) {
        code = glyphId;
        if (glyphId <= 0 || string.IsNullOrEmpty(unicodeText)) return false;
        unicodeText = OfficeArabicTextShaper.ToLogicalText(unicodeText);
        lock (_lock) {
            if (!_pins.TryGetValue(glyphId, out string? pinned)) {
                pinned = _canonical.TryGet(glyphId, out string canonical) ? canonical : unicodeText;
                _pins.Add(glyphId, pinned);
            }
            if (string.Equals(pinned, unicodeText, StringComparison.Ordinal)) return true;
            if (!_allowAliases || !TryGetSingleScalar(unicodeText, out int scalar) ||
                !_canonical.IsNominalGlyph(scalar, glyphId) || (long)_glyphCount + scalar > 0xFFFF) return false;
            code = _glyphCount + scalar;
            _aliases[code] = glyphId;
            return true;
        }
    }

    internal bool TryGetPinned(int glyphId, out string unicodeText) {
        lock (_lock) {
            if (_pins.TryGetValue(glyphId, out string? pinned)) {
                unicodeText = pinned;
                return true;
            }
        }
        unicodeText = string.Empty;
        return false;
    }

    internal IReadOnlyList<(int Code, int GlyphId, string UnicodeText)> GetAliases() {
        lock (_lock) {
            var result = new List<(int, int, string)>(_aliases.Count);
            foreach (KeyValuePair<int, int> alias in _aliases)
                result.Add((alias.Key, alias.Value, OfficeArabicTextShaper.ToLogicalText(char.ConvertFromUtf32(alias.Key - _glyphCount))));
            return result;
        }
    }

    // Snapshot first: option-context merges must not hold two pin locks in opposing order.
    internal void MergeFrom(PdfGlyphTextPins other) {
        if (ReferenceEquals(this, other)) return;
        KeyValuePair<int, string>[] pins;
        KeyValuePair<int, int>[] aliases;
        lock (other._lock) {
            pins = other._pins.ToArray();
            aliases = other._aliases.ToArray();
        }
        lock (_lock) {
            foreach (KeyValuePair<int, string> pin in pins) {
                if (!_pins.ContainsKey(pin.Key)) _pins.Add(pin.Key, pin.Value);
            }
            foreach (KeyValuePair<int, int> alias in aliases) _aliases[alias.Key] = alias.Value;
        }
    }

    internal void Clear() {
        lock (_lock) {
            _pins.Clear();
            _aliases.Clear();
        }
    }

    private static bool TryGetSingleScalar(string text, out int scalar) {
        scalar = 0;
        if (text.Length == 1 && !char.IsSurrogate(text[0])) scalar = text[0];
        else if (text.Length == 2 && char.IsSurrogatePair(text[0], text[1])) scalar = char.ConvertToUtf32(text[0], text[1]);
        else return false;
        return true;
    }
}

/// <summary>Immutable glyph-to-text map derived from a font cmap, shared by document forks.</summary>
internal sealed class PdfCanonicalGlyphText {
    private readonly Dictionary<int, int> _cmap;
    private Dictionary<int, string>? _texts;
    private readonly object _lock = new();

    internal PdfCanonicalGlyphText(Dictionary<int, int> cmap) {
        _cmap = cmap ?? throw new ArgumentNullException(nameof(cmap));
    }

    internal bool IsNominalGlyph(int scalar, int glyphId) =>
        _cmap.TryGetValue(scalar, out int mapped) && mapped == glyphId;

    internal bool TryGet(int glyphId, out string text) {
        Dictionary<int, string> texts = _texts ?? Build();
        if (texts.TryGetValue(glyphId, out string? value)) {
            text = value;
            return true;
        }
        text = string.Empty;
        return false;
    }

    private Dictionary<int, string> Build() {
        lock (_lock) {
            if (_texts != null) return _texts;
            var lowest = new Dictionary<int, int>(); // preferred scalar per glyph
            foreach (KeyValuePair<int, int> entry in _cmap) {
                if (entry.Value <= 0 || !IsTextScalar(entry.Key)) continue;
                if (!lowest.TryGetValue(entry.Value, out int scalar) || Prefer(entry.Key, scalar)) lowest[entry.Value] = entry.Key;
            }
            var texts = new Dictionary<int, string>(lowest.Count);
            foreach (KeyValuePair<int, int> entry in lowest) texts.Add(entry.Key, ToSourceText(entry.Value));
            _texts = texts;
            return texts;
        }
    }

    // Ordinary characters win over presentation forms; Arabic Presentation Forms-B (the base letters)
    // win over Forms-A (extended and language-specific letters), so a shared yeh glyph reads as U+064A.
    private static bool Prefer(int candidate, int current) {
        int candidateRank = PresentationRank(candidate), currentRank = PresentationRank(current);
        return candidateRank != currentRank ? candidateRank < currentRank : candidate < current;
    }

    private static int PresentationRank(int scalar) =>
        scalar >= 0xFE70 && scalar <= 0xFEFF ? 1 : scalar >= 0xFB00 && scalar <= 0xFDFF ? 2 : 0;

    // Presentation forms describe how text is drawn; extraction should recover the source letters.
    private static string ToSourceText(int scalar) {
        string text = char.ConvertFromUtf32(scalar);
        if (scalar >= 0xFB00 && scalar <= 0xFB06) {
            string decomposed = text.Normalize(System.Text.NormalizationForm.FormKD);
            if (decomposed.Length > 1) return decomposed;
        }
        return OfficeArabicTextShaper.ToLogicalText(text);
    }

    private static bool IsTextScalar(int scalar) =>
        scalar >= 0x20 && !(scalar >= 0x7F && scalar <= 0x9F) && !(scalar >= 0xD800 && scalar <= 0xDFFF) &&
        (scalar & 0xFFFE) != 0xFFFE && !(scalar >= 0xFDD0 && scalar <= 0xFDEF) && scalar <= 0x10FFFF &&
        scalar != 0x00AD && !(scalar >= 0x200B && scalar <= 0x200F) && !(scalar >= 0x2028 && scalar <= 0x202E) &&
        !(scalar >= 0x2060 && scalar <= 0x206F) && scalar != 0xFEFF && !(scalar >= 0xFE00 && scalar <= 0xFE0F);
}
