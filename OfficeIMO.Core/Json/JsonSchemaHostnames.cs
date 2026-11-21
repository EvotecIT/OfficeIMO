using System.Globalization;
using System.Text;

namespace OfficeIMO.Json;

/// <summary>
/// The <c>hostname</c> and <c>idn-hostname</c> formats: RFC 1123 labels, and IDNA2008 labels (RFC 5890-5893) whose
/// A-labels decode as Punycode (RFC 3492). Code point validity, contextual rules and the Bidi rule are approximated
/// from Unicode general categories and script ranges rather than the full IDNA2008 tables.
/// </summary>
internal static class JsonSchemaHostnames {
    private static readonly char[] IdnSeparators = { '.', '。', '．', '｡' };

    public static bool IsHostname(string value) => IsName(value, new[] { '.' }, international: false);

    /// <summary>Checks an internationalized host name after the UTS 46 mapping, approximated as removing ignored code points, NFKC and lowercasing.</summary>
    public static bool IsIdnHostname(string value) => !HasUnpairedSurrogate(value) && IsName(MapUts46(value), IdnSeparators, international: true);

    private static bool HasUnpairedSurrogate(string value) {
        for (int i = 0; i < value.Length; i++) {
            if (char.IsHighSurrogate(value[i]) && i + 1 < value.Length && char.IsLowSurrogate(value[i + 1])) i++;
            else if (char.IsSurrogate(value[i])) return true;
        }
        return false;
    }

    private static string MapUts46(string value) {
        var text = new StringBuilder(value.Length);
        foreach (char c in value) {
            if (c is '\u00AD' or '\u034F' or '\u200B' or '\u2060' or '\uFEFF' || c >= '\u180B' && c <= '\u180F' || c >= '\uFE00' && c <= '\uFE0F') continue;
            text.Append(c);
        }
        return text.ToString().Normalize(NormalizationForm.FormKC).ToLowerInvariant();
    }

    private static bool IsName(string value, char[] separators, bool international) {
        if (value.Length == 0) return false;
        string[] labels = value.Split(separators);
        var unicodeLabels = new List<string>(labels.Length);
        int asciiLength = labels.Length - 1;
        foreach (string label in labels) {
            if (label.Length == 0) return false;
            string? ascii = ToAsciiLabel(label, international, out string? unicode);
            if (ascii is null || ascii.Length > 63) return false;
            asciiLength += ascii.Length;
            unicodeLabels.Add(unicode ?? label);
        }
        if (asciiLength > 253) return false;
        // The Bidi rule applies to every label once any label is right-to-left.
        if (unicodeLabels.Any(IsRightToLeftLabel)) {
            foreach (string label in unicodeLabels) {
                if (!SatisfiesBidiRule(label)) return false;
            }
        }
        return true;
    }

    /// <summary>Validates one label and returns its ASCII form, or null when it is invalid.</summary>
    private static string? ToAsciiLabel(string label, bool international, out string? unicode) {
        unicode = null;
        bool ascii = label.All(c => c < 0x80);
        if (!ascii) {
            if (!international || !IsValidULabel(label)) return null;
            unicode = label;
            string? encoded = Punycode.Encode(label);
            return encoded is null ? null : "xn--" + encoded;
        }
        if (label[0] == '-' || label[label.Length - 1] == '-') return null;
        foreach (char c in label) {
            if (!(IsAsciiLetterOrDigit(c) || c == '-')) return null;
        }
        if (label.Length >= 4 && label[2] == '-' && label[3] == '-') {
            if (!label.StartsWith("xn--", StringComparison.OrdinalIgnoreCase)) return null;
            string? decoded = Punycode.Decode(label.Substring(4));
            // An A-label must decode to a non-ASCII U-label that encodes back to the same text.
            if (decoded is null || decoded.All(c => c < 0x80) || !IsValidULabel(decoded)) return null;
            if (!string.Equals("xn--" + Punycode.Encode(decoded), label, StringComparison.OrdinalIgnoreCase)) return null;
            unicode = decoded;
        }
        return label;
    }

    private static bool IsValidULabel(string label) {
        if (label.Length == 0 || label[0] == '-' || label[label.Length - 1] == '-') return false;
        if (label.Length >= 4 && label[2] == '-' && label[3] == '-') return false;
        if (!label.IsNormalized(NormalizationForm.FormC)) return false;
        int[] codePoints = ToCodePoints(label);
        if (IsMark(codePoints[0])) return false;
        bool arabicIndic = false, extendedArabicIndic = false;
        for (int i = 0; i < codePoints.Length; i++) {
            int cp = codePoints[i];
            if (cp >= 0x0660 && cp <= 0x0669) arabicIndic = true;
            if (cp >= 0x06F0 && cp <= 0x06F9) extendedArabicIndic = true;
            switch (cp) {
                case 0x00B7: // MIDDLE DOT: between two 'l'.
                    if (i == 0 || i == codePoints.Length - 1 || codePoints[i - 1] != 'l' || codePoints[i + 1] != 'l') return false;
                    continue;
                case 0x0375: // GREEK KERAIA: followed by Greek.
                    if (i == codePoints.Length - 1 || !IsGreek(codePoints[i + 1])) return false;
                    continue;
                case 0x05F3:
                case 0x05F4: // HEBREW GERESH and GERSHAYIM: preceded by Hebrew.
                    if (i == 0 || !IsHebrew(codePoints[i - 1])) return false;
                    continue;
                case 0x30FB: // KATAKANA MIDDLE DOT: the label has Hiragana, Katakana or Han.
                    if (!codePoints.Any(c => c != 0x30FB && IsHiraganaKatakanaOrHan(c))) return false;
                    continue;
                case 0x200D: // ZERO WIDTH JOINER: after a virama.
                    if (i == 0 || !IsVirama(codePoints[i - 1])) return false;
                    continue;
                case 0x200C: // ZERO WIDTH NON-JOINER: after a virama, or between joining letters.
                    if (i > 0 && IsVirama(codePoints[i - 1])) continue;
                    if (!HasJoiningNeighbour(codePoints, i, -1) || !HasJoiningNeighbour(codePoints, i, 1)) return false;
                    continue;
            }
            if (!IsPermitted(cp)) return false;
        }
        return !(arabicIndic && extendedArabicIndic);
    }

    /// <summary>Approximates IDNA2008 PVALID: lowercase letters, other letters, marks and digits that NFKC and case folding leave unchanged.</summary>
    private static bool IsPermitted(int cp) {
        switch (cp) {
            case 0x00DF: case 0x03C2: case 0x06FD: case 0x06FE: case 0x0F0B: case 0x3007:
                return true;
            case 0x0640: case 0x07FA: case 0x302E: case 0x302F: case 0x3031: case 0x3032: case 0x3033: case 0x3034: case 0x3035: case 0x303B:
                return false;
        }
        if (cp == '-') return true;
        if (cp >= 0x0660 && cp <= 0x0669 || cp >= 0x06F0 && cp <= 0x06F9) return true;
        string text = char.ConvertFromUtf32(cp);
        UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(text, 0);
        if (category is not (UnicodeCategory.LowercaseLetter or UnicodeCategory.OtherLetter or UnicodeCategory.ModifierLetter
            or UnicodeCategory.NonSpacingMark or UnicodeCategory.SpacingCombiningMark or UnicodeCategory.DecimalDigitNumber)) return false;
        return text.Normalize(NormalizationForm.FormKC) == text && text.ToLowerInvariant() == text;
    }

    private static bool HasJoiningNeighbour(int[] codePoints, int index, int step) {
        for (int i = index + step; i >= 0 && i < codePoints.Length; i += step) {
            if (IsMark(codePoints[i])) continue;
            return IsJoiningLetter(codePoints[i]);
        }
        return false;
    }

    private static bool IsJoiningLetter(int cp) => cp >= 0x0620 && cp <= 0x064A || cp >= 0x066E && cp <= 0x06D3 || cp >= 0x06FA && cp <= 0x06FC
        || cp >= 0x0710 && cp <= 0x072F || cp >= 0x074D && cp <= 0x077F || cp >= 0x08A0 && cp <= 0x08C9 || cp >= 0xA840 && cp <= 0xA872;

    private static bool IsVirama(int cp) => cp is 0x094D or 0x09CD or 0x0A4D or 0x0ACD or 0x0B4D or 0x0BCD or 0x0C4D or 0x0CCD
        or 0x0D3B or 0x0D3C or 0x0D4D or 0x0DCA or 0x0E3A or 0x0EBA or 0x0F84 or 0x1039 or 0x103A or 0x1714 or 0x1734
        or 0x17D2 or 0x1A60 or 0x1B44 or 0x1BAA or 0x1BAB or 0x1BF2 or 0x1BF3 or 0x2D7F or 0xA806 or 0xA8C4 or 0xA953
        or 0xA9C0 or 0xAAF6 or 0xABED or 0x10A3F or 0x11046 or 0x1107F or 0x110B9 or 0x11133 or 0x11134 or 0x111C0
        or 0x11235 or 0x112EA or 0x1134D or 0x11442 or 0x114C2 or 0x115BF or 0x1163F or 0x116B6 or 0x1172B or 0x11839;

    private static bool IsGreek(int cp) => cp >= 0x0370 && cp <= 0x03FF && cp != 0x0375 && cp != 0x037E || cp >= 0x1F00 && cp <= 0x1FFF;

    private static bool IsHebrew(int cp) => cp >= 0x0591 && cp <= 0x05F4 || cp >= 0xFB1D && cp <= 0xFB4F;

    private static bool IsHiraganaKatakanaOrHan(int cp) => cp >= 0x3040 && cp <= 0x30FF || cp >= 0x31F0 && cp <= 0x31FF
        || cp >= 0x3400 && cp <= 0x4DBF || cp >= 0x4E00 && cp <= 0x9FFF || cp >= 0xF900 && cp <= 0xFAFF || cp >= 0x20000 && cp <= 0x3134F;

    private static bool IsMark(int cp) {
        UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(char.ConvertFromUtf32(cp), 0);
        return category is UnicodeCategory.NonSpacingMark or UnicodeCategory.SpacingCombiningMark or UnicodeCategory.EnclosingMark;
    }

    #region Bidi rule (RFC 5893)

    private enum BidiClass { L, R, AL, AN, EN, ES, CS, ET, ON, BN, NSM }

    private static bool IsRightToLeftLabel(string label) => ToCodePoints(label).Any(cp => GetBidiClass(cp) is BidiClass.R or BidiClass.AL or BidiClass.AN);

    private static bool SatisfiesBidiRule(string label) {
        BidiClass[] classes = ToCodePoints(label).Select(GetBidiClass).ToArray();
        BidiClass first = classes[0];
        int last = classes.Length - 1;
        while (last > 0 && classes[last] == BidiClass.NSM) last--;
        if (first is BidiClass.R or BidiClass.AL) {
            bool en = false, an = false;
            foreach (BidiClass bidi in classes) {
                if (bidi == BidiClass.L) return false;
                if (bidi == BidiClass.EN) en = true;
                if (bidi == BidiClass.AN) an = true;
            }
            return !(en && an) && classes[last] is BidiClass.R or BidiClass.AL or BidiClass.EN or BidiClass.AN;
        }
        if (first != BidiClass.L) return false;
        foreach (BidiClass bidi in classes) {
            if (bidi is BidiClass.R or BidiClass.AL or BidiClass.AN) return false;
        }
        return classes[last] is BidiClass.L or BidiClass.EN;
    }

    private static BidiClass GetBidiClass(int cp) {
        if (IsMark(cp) && CharUnicodeInfo.GetUnicodeCategory(char.ConvertFromUtf32(cp), 0) != UnicodeCategory.SpacingCombiningMark) return BidiClass.NSM;
        if (cp >= 0x0600 && cp <= 0x0605 || cp >= 0x0660 && cp <= 0x0669 || cp is 0x066B or 0x066C or 0x06DD or 0x08E2) return BidiClass.AN;
        if (cp >= '0' && cp <= '9' || cp >= 0x06F0 && cp <= 0x06F9 || cp is 0x00B2 or 0x00B3 or 0x00B9) return BidiClass.EN;
        if (cp >= 0x0590 && cp <= 0x05FF || cp >= 0x07C0 && cp <= 0x085F || cp >= 0xFB1D && cp <= 0xFB4F || cp >= 0x10800 && cp <= 0x10FFF) return BidiClass.R;
        if (cp >= 0x0600 && cp <= 0x07BF || cp >= 0x0860 && cp <= 0x08FF || cp >= 0xFB50 && cp <= 0xFDFF || cp >= 0xFE70 && cp <= 0xFEFF) return BidiClass.AL;
        if (cp == '-' || cp == '+') return BidiClass.ES;
        if (cp is 0x200C or 0x200D) return BidiClass.BN;
        return BidiClass.L;
    }

    #endregion

    private static int[] ToCodePoints(string text) {
        var codePoints = new List<int>(text.Length);
        for (int i = 0; i < text.Length; i++) {
            if (char.IsHighSurrogate(text[i]) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1])) {
                codePoints.Add(char.ConvertToUtf32(text[i], text[i + 1]));
                i++;
            } else {
                codePoints.Add(text[i]);
            }
        }
        return codePoints.ToArray();
    }

    private static bool IsAsciiLetterOrDigit(char c) => c >= '0' && c <= '9' || c >= 'a' && c <= 'z' || c >= 'A' && c <= 'Z';

    /// <summary>Punycode (RFC 3492) for IDNA labels.</summary>
    private static class Punycode {
        private const int Base = 36, TMin = 1, TMax = 26, Skew = 38, Damp = 700, InitialBias = 72, InitialN = 128;

        public static string? Decode(string input) {
            var output = new List<int>();
            int basic = input.LastIndexOf('-');
            for (int j = 0; j < Math.Max(basic, 0); j++) {
                if (input[j] >= 0x80) return null;
                output.Add(input[j]);
            }
            int n = InitialN, i = 0, bias = InitialBias;
            for (int index = basic > 0 ? basic + 1 : 0; index < input.Length;) {
                int oldi = i, w = 1;
                for (int k = Base; ; k += Base) {
                    if (index >= input.Length) return null;
                    int digit = DecodeDigit(input[index++]);
                    if (digit < 0 || digit > (int.MaxValue - i) / w) return null;
                    i += digit * w;
                    int t = k <= bias ? TMin : k >= bias + TMax ? TMax : k - bias;
                    if (digit < t) break;
                    if (w > int.MaxValue / (Base - t)) return null;
                    w *= Base - t;
                }
                bias = Adapt(i - oldi, output.Count + 1, oldi == 0);
                if (i / (output.Count + 1) > int.MaxValue - n) return null;
                n += i / (output.Count + 1);
                i %= output.Count + 1;
                if (n > 0x10FFFF || n >= 0xD800 && n <= 0xDFFF) return null;
                output.Insert(i++, n);
            }
            var text = new StringBuilder(output.Count);
            foreach (int cp in output) text.Append(char.ConvertFromUtf32(cp));
            return text.ToString();
        }

        public static string? Encode(string input) {
            int[] codePoints = ToCodePoints(input);
            var output = new StringBuilder();
            foreach (int cp in codePoints) {
                if (cp < 0x80) output.Append((char)cp);
            }
            int handled = output.Length, basic = handled;
            if (basic > 0) output.Append('-');
            int n = InitialN, delta = 0, bias = InitialBias;
            while (handled < codePoints.Length) {
                int m = codePoints.Where(cp => cp >= n).Min();
                if (m - n > (int.MaxValue - delta) / (handled + 1)) return null;
                delta += (m - n) * (handled + 1);
                n = m;
                foreach (int cp in codePoints) {
                    if (cp < n && ++delta == 0) return null;
                    if (cp != n) continue;
                    int q = delta;
                    for (int k = Base; ; k += Base) {
                        int t = k <= bias ? TMin : k >= bias + TMax ? TMax : k - bias;
                        if (q < t) break;
                        output.Append(EncodeDigit(t + (q - t) % (Base - t)));
                        q = (q - t) / (Base - t);
                    }
                    output.Append(EncodeDigit(q));
                    bias = Adapt(delta, handled + 1, handled == basic);
                    delta = 0;
                    handled++;
                }
                delta++;
                n++;
            }
            return output.ToString();
        }

        private static int Adapt(int delta, int points, bool first) {
            delta = first ? delta / Damp : delta / 2;
            delta += delta / points;
            int k = 0;
            while (delta > (Base - TMin) * TMax / 2) {
                delta /= Base - TMin;
                k += Base;
            }
            return k + (Base - TMin + 1) * delta / (delta + Skew);
        }

        private static int DecodeDigit(char c) => c >= '0' && c <= '9' ? c - '0' + 26 : c >= 'a' && c <= 'z' ? c - 'a' : c >= 'A' && c <= 'Z' ? c - 'A' : -1;

        private static char EncodeDigit(int d) => (char)(d < 26 ? 'a' + d : '0' + d - 26);
    }
}
