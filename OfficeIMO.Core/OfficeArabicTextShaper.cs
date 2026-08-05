using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Drawing;

/// <summary>
/// Applies dependency-free contextual presentation forms for the core Arabic alphabet.
/// </summary>
/// <remarks>
/// The shaper preserves one UTF-16 character per base letter and intentionally does not
/// synthesize optional ligatures. This keeps layout and logical-text mapping deterministic
/// while improving joining for renderers that do not execute OpenType shaping tables.
/// </remarks>
public static class OfficeArabicTextShaper {
    private static readonly IReadOnlyDictionary<char, ArabicForms> Forms = CreateForms();
    private static readonly IReadOnlyDictionary<char, char> LogicalForms = CreateLogicalForms();

    /// <summary>Shapes supported Arabic letters into their contextual presentation forms.</summary>
    public static string Shape(string? value) {
        if (string.IsNullOrEmpty(value)) return value ?? string.Empty;

        string logical = ToLogicalText(value);
        var result = new StringBuilder(logical.Length);
        for (int index = 0; index < logical.Length; index++) {
            char current = logical[index];
            if (!Forms.TryGetValue(current, out ArabicForms forms)) {
                result.Append(current);
                continue;
            }

            bool joinsPrevious = forms.CanJoinPrevious && TryFindPrevious(logical, index, out ArabicForms previous) && previous.CanJoinNext;
            bool joinsNext = forms.CanJoinNext && TryFindNext(logical, index, out ArabicForms next) && next.CanJoinPrevious;
            result.Append(forms.Resolve(joinsPrevious, joinsNext));
        }

        return result.ToString();
    }

    /// <summary>
    /// Returns <see langword="true"/> when every joining-script letter in the value is covered
    /// by the bounded core-Arabic shaper. Marks, controls, punctuation, and digits are neutral.
    /// </summary>
    public static bool CanShapeAllJoiningCharacters(string? value) {
        if (string.IsNullOrEmpty(value)) return true;
        for (int index = 0; index < value!.Length; index++) {
            char character = value[index];
            int scalar = character;
            UnicodeCategory category = char.GetUnicodeCategory(character);
            if (char.IsHighSurrogate(character) && index + 1 < value.Length && char.IsLowSurrogate(value[index + 1])) {
                scalar = char.ConvertToUtf32(character, value[index + 1]);
                category = CharUnicodeInfo.GetUnicodeCategory(value, index);
                index++;
            }
            if (!OfficeTextElements.IsRightToLeftScalar(scalar) || !IsLetter(category)) continue;
            if (scalar <= char.MaxValue && (Forms.ContainsKey((char)scalar) || LogicalForms.ContainsKey((char)scalar))) continue;
            return false;
        }
        return true;
    }

    /// <summary>Maps presentation forms produced by this shaper back to core Arabic letters.</summary>
    public static string ToLogicalText(string? value) {
        if (string.IsNullOrEmpty(value)) return value ?? string.Empty;
        var result = new StringBuilder(value!.Length);
        foreach (char character in value) {
            result.Append(LogicalForms.TryGetValue(character, out char logical) ? logical : character);
        }
        return result.ToString();
    }

    private static bool TryFindPrevious(string value, int index, out ArabicForms forms) {
        for (int current = index - 1; current >= 0; current--) {
            char character = value[current];
            if (character == '\u200C') break;
            if (IsTransparent(character)) continue;
            if (character == '\u0640') {
                forms = ArabicForms.JoiningCarrier;
                return true;
            }
            if (Forms.TryGetValue(character, out forms)) return true;
            break;
        }
        forms = default;
        return false;
    }

    private static bool TryFindNext(string value, int index, out ArabicForms forms) {
        for (int current = index + 1; current < value.Length; current++) {
            char character = value[current];
            if (character == '\u200C') break;
            if (IsTransparent(character)) continue;
            if (character == '\u0640') {
                forms = ArabicForms.JoiningCarrier;
                return true;
            }
            if (Forms.TryGetValue(character, out forms)) return true;
            break;
        }
        forms = default;
        return false;
    }

    private static bool IsTransparent(char value) =>
        value == '\u200D' || CharUnicodeInfo.GetUnicodeCategory(value) is UnicodeCategory.NonSpacingMark or UnicodeCategory.EnclosingMark or UnicodeCategory.Format;

    private static bool IsLetter(UnicodeCategory category) => category is
        UnicodeCategory.UppercaseLetter or UnicodeCategory.LowercaseLetter or
        UnicodeCategory.TitlecaseLetter or UnicodeCategory.ModifierLetter or UnicodeCategory.OtherLetter;

    private static IReadOnlyDictionary<char, ArabicForms> CreateForms() => new Dictionary<char, ArabicForms> {
        ['\u0621'] = new('\uFE80'),
        ['\u0622'] = new('\uFE81', '\uFE82'),
        ['\u0623'] = new('\uFE83', '\uFE84'),
        ['\u0624'] = new('\uFE85', '\uFE86'),
        ['\u0625'] = new('\uFE87', '\uFE88'),
        ['\u0626'] = new('\uFE89', '\uFE8A', '\uFE8B', '\uFE8C'),
        ['\u0627'] = new('\uFE8D', '\uFE8E'),
        ['\u0628'] = new('\uFE8F', '\uFE90', '\uFE91', '\uFE92'),
        ['\u0629'] = new('\uFE93', '\uFE94'),
        ['\u062A'] = new('\uFE95', '\uFE96', '\uFE97', '\uFE98'),
        ['\u062B'] = new('\uFE99', '\uFE9A', '\uFE9B', '\uFE9C'),
        ['\u062C'] = new('\uFE9D', '\uFE9E', '\uFE9F', '\uFEA0'),
        ['\u062D'] = new('\uFEA1', '\uFEA2', '\uFEA3', '\uFEA4'),
        ['\u062E'] = new('\uFEA5', '\uFEA6', '\uFEA7', '\uFEA8'),
        ['\u062F'] = new('\uFEA9', '\uFEAA'),
        ['\u0630'] = new('\uFEAB', '\uFEAC'),
        ['\u0631'] = new('\uFEAD', '\uFEAE'),
        ['\u0632'] = new('\uFEAF', '\uFEB0'),
        ['\u0633'] = new('\uFEB1', '\uFEB2', '\uFEB3', '\uFEB4'),
        ['\u0634'] = new('\uFEB5', '\uFEB6', '\uFEB7', '\uFEB8'),
        ['\u0635'] = new('\uFEB9', '\uFEBA', '\uFEBB', '\uFEBC'),
        ['\u0636'] = new('\uFEBD', '\uFEBE', '\uFEBF', '\uFEC0'),
        ['\u0637'] = new('\uFEC1', '\uFEC2', '\uFEC3', '\uFEC4'),
        ['\u0638'] = new('\uFEC5', '\uFEC6', '\uFEC7', '\uFEC8'),
        ['\u0639'] = new('\uFEC9', '\uFECA', '\uFECB', '\uFECC'),
        ['\u063A'] = new('\uFECD', '\uFECE', '\uFECF', '\uFED0'),
        ['\u0641'] = new('\uFED1', '\uFED2', '\uFED3', '\uFED4'),
        ['\u0642'] = new('\uFED5', '\uFED6', '\uFED7', '\uFED8'),
        ['\u0643'] = new('\uFED9', '\uFEDA', '\uFEDB', '\uFEDC'),
        ['\u0644'] = new('\uFEDD', '\uFEDE', '\uFEDF', '\uFEE0'),
        ['\u0645'] = new('\uFEE1', '\uFEE2', '\uFEE3', '\uFEE4'),
        ['\u0646'] = new('\uFEE5', '\uFEE6', '\uFEE7', '\uFEE8'),
        ['\u0647'] = new('\uFEE9', '\uFEEA', '\uFEEB', '\uFEEC'),
        ['\u0648'] = new('\uFEED', '\uFEEE'),
        ['\u0649'] = new('\uFEEF', '\uFEF0'),
        ['\u064A'] = new('\uFEF1', '\uFEF2', '\uFEF3', '\uFEF4')
    };

    private static IReadOnlyDictionary<char, char> CreateLogicalForms() {
        var result = new Dictionary<char, char>();
        foreach (KeyValuePair<char, ArabicForms> entry in Forms) {
            entry.Value.AddLogicalMappings(result, entry.Key);
        }
        return result;
    }

    private readonly struct ArabicForms {
        internal static ArabicForms JoiningCarrier { get; } = new('\0', '\u0640', '\u0640', '\u0640');

        internal ArabicForms(char isolated, char final = '\0', char initial = '\0', char medial = '\0') {
            Isolated = isolated;
            Final = final;
            Initial = initial;
            Medial = medial;
        }

        internal char Isolated { get; }
        internal char Final { get; }
        internal char Initial { get; }
        internal char Medial { get; }
        internal bool CanJoinPrevious => Final != '\0' || Medial != '\0';
        internal bool CanJoinNext => Initial != '\0' || Medial != '\0';

        internal char Resolve(bool joinsPrevious, bool joinsNext) {
            if (joinsPrevious && joinsNext && Medial != '\0') return Medial;
            if (joinsPrevious && Final != '\0') return Final;
            if (joinsNext && Initial != '\0') return Initial;
            return Isolated;
        }

        internal void AddLogicalMappings(IDictionary<char, char> mappings, char logical) {
            foreach (char form in new[] { Isolated, Final, Initial, Medial }) {
                if (form != '\0') mappings[form] = logical;
            }
        }
    }
}
