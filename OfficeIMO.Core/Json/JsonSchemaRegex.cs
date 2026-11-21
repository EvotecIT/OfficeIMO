using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Json;

/// <summary>
/// Translates ECMA-262 regular expressions, the dialect JSON Schema patterns use, to .NET: <c>\d</c>, <c>\w</c>,
/// <c>\s</c> and <c>\b</c> keep their ECMA-262 character sets, <c>$</c> matches only at the end, <c>.</c> and literal
/// astral characters match whole code points, captures keep their ECMA-262 numbers, <c>(?m:...)</c> and
/// <c>(?s:...)</c> groups change how <c>^</c>, <c>$</c> and <c>.</c> match, <c>\p{Letter}</c>-style property names map to .NET categories, and syntax
/// ECMA-262 does not have (<c>\a</c>, <c>\z</c>, <c>(?#...)</c>, global <c>(?i)</c> flags) is rejected.
/// </summary>
internal static class JsonSchemaRegex {
    private const string AstralInClass = "Characters outside the Basic Multilingual Plane are not supported in character classes.";
    private const string DigitRanges = "0-9";
    private const string NotDigitRanges = "\\u0000-\\u002F\\u003A-\\uFFFF";
    private const string WordRanges = "0-9A-Za-z_";
    private const string NotWordRanges = "\\u0000-\\u002F\\u003A-\\u0040\\u005B-\\u005E\\u0060\\u007B-\\uFFFF";
    private const string SpaceRanges = "\\u0009-\\u000D\\u0020\\u00A0\\u1680\\u2000-\\u200A\\u2028\\u2029\\u202F\\u205F\\u3000\\uFEFF";
    private const string NotSpaceRanges = "\\u0000-\\u0008\\u000E-\\u001F\\u0021-\\u009F\\u00A1-\\u167F\\u1681-\\u1FFF\\u200B-\\u2027\\u202A-\\u202E\\u2030-\\u205E\\u2060-\\u2FFF\\u3001-\\uFEFE\\uFF00-\\uFFFF";
    private const string Pair = "[\\uD800-\\uDBFF][\\uDC00-\\uDFFF]";
    // One UTF-16 unit that is not half of a surrogate pair, so backtracking cannot split a pair between two dots.
    private const string LoneUnit = "(?!" + Pair + ")(?!(?<=[\\uD800-\\uDBFF])[\\uDC00-\\uDFFF])";
    private const string LineTerminators = "\\n\\r\\u2028\\u2029";
    private const string AnyCodePoint = "(?:" + Pair + "|" + LoneUnit + "[^" + LineTerminators + "])";
    private const string AnyCodePointDotAll = "(?:" + Pair + "|" + LoneUnit + "[\\s\\S])";
    private const string LineStart = "(?<![^" + LineTerminators + "])";
    private const string LineEnd = "(?![^" + LineTerminators + "])";
    private const string WordBoundary = "(?:(?<=[" + WordRanges + "])(?![" + WordRanges + "])|(?<![" + WordRanges + "])(?=[" + WordRanges + "]))";
    private const string NotWordBoundary = "(?:(?<=[" + WordRanges + "])(?=[" + WordRanges + "])|(?<![" + WordRanges + "])(?![" + WordRanges + "]))";

    private static readonly Dictionary<string, string> Properties = new Dictionary<string, string>(StringComparer.Ordinal) {
        ["Letter"] = "L", ["Uppercase_Letter"] = "Lu", ["Lowercase_Letter"] = "Ll", ["Titlecase_Letter"] = "Lt",
        ["Modifier_Letter"] = "Lm", ["Other_Letter"] = "Lo", ["Mark"] = "M", ["Combining_Mark"] = "M",
        ["Nonspacing_Mark"] = "Mn", ["Spacing_Mark"] = "Mc", ["Enclosing_Mark"] = "Me", ["Number"] = "N",
        ["Decimal_Number"] = "Nd", ["digit"] = "Nd", ["Letter_Number"] = "Nl", ["Other_Number"] = "No",
        ["Punctuation"] = "P", ["punct"] = "P", ["Connector_Punctuation"] = "Pc", ["Dash_Punctuation"] = "Pd",
        ["Open_Punctuation"] = "Ps", ["Close_Punctuation"] = "Pe", ["Initial_Punctuation"] = "Pi",
        ["Final_Punctuation"] = "Pf", ["Other_Punctuation"] = "Po", ["Symbol"] = "S", ["Math_Symbol"] = "Sm",
        ["Currency_Symbol"] = "Sc", ["Modifier_Symbol"] = "Sk", ["Other_Symbol"] = "So", ["Separator"] = "Z",
        ["Space_Separator"] = "Zs", ["Line_Separator"] = "Zl", ["Paragraph_Separator"] = "Zp", ["Other"] = "C",
        ["Control"] = "Cc", ["cntrl"] = "Cc", ["Format"] = "Cf", ["Surrogate"] = "Cs", ["Private_Use"] = "Co",
        ["Unassigned"] = "Cn"
    };

    /// <summary>Compiles an ECMA-262 pattern; throws <see cref="ArgumentException"/> when it is not valid ECMA-262.</summary>
    public static Regex Create(string pattern, TimeSpan timeout) => new Regex(Translate(pattern), RegexOptions.CultureInvariant, timeout);

    public static string Translate(string pattern) {
        var state = new State(pattern);
        StringBuilder text = state.Text;
        bool inClass = false;
        int i = 0;
        while (i < pattern.Length) {
            char c = pattern[i];
            if (c == '\\') {
                i = TranslateEscape(pattern, i, inClass, state);
                continue;
            }
            if (inClass) {
                if (char.IsHighSurrogate(c) && char.IsLowSurrogate(At(pattern, i + 1))) {
                    // A .NET class matches one UTF-16 unit, so an astral member would match half a character.
                    throw new ArgumentException(AstralInClass);
                }
                if (c == ']') inClass = false;
                text.Append(c == '[' ? "\\[" : c.ToString());
                i++;
                continue;
            }
            switch (c) {
                case '[':
                    if (At(pattern, i + 1) == ']') {
                        text.Append("(?!)");
                        i += 2;
                        continue;
                    }
                    if (At(pattern, i + 1) == '^' && At(pattern, i + 2) == ']') {
                        text.Append("[\\s\\S]");
                        i += 3;
                        continue;
                    }
                    inClass = true;
                    text.Append('[');
                    i++;
                    if (At(pattern, i) == '^') {
                        text.Append('^');
                        i++;
                    }
                    continue;
                case '(':
                    i = TranslateGroup(pattern, i, state);
                    continue;
                case ')':
                    state.Close();
                    text.Append(')');
                    break;
                case '^':
                    text.Append(state.Multiline ? LineStart : "\\A");
                    break;
                case '$':
                    text.Append(state.Multiline ? LineEnd : "\\z");
                    break;
                case '.':
                    text.Append(state.DotAll ? AnyCodePointDotAll : AnyCodePoint);
                    break;
                default:
                    if (char.IsHighSurrogate(c) && char.IsLowSurrogate(At(pattern, i + 1))) {
                        text.Append("(?:").Append(c).Append(pattern[i + 1]).Append(')');
                        i += 2;
                        continue;
                    }
                    text.Append(c);
                    break;
            }
            i++;
        }
        if (inClass) throw new ArgumentException("The character class is not closed.");
        return text.ToString();
    }

    private static int TranslateGroup(string pattern, int i, State state) {
        StringBuilder text = state.Text;
        if (At(pattern, i + 1) != '?') {
            // Every capture gets its ECMA-262 number explicitly: .NET would number named groups after unnamed ones.
            state.Open(state.Multiline, state.DotAll);
            text.Append("(?<").Append((++state.Captures).ToString(CultureInfo.InvariantCulture)).Append('>');
            return i + 1;
        }
        char kind = At(pattern, i + 2);
        if (kind is ':' or '=' or '!') {
            state.Open(state.Multiline, state.DotAll);
            text.Append(pattern, i, 3);
            return i + 3;
        }
        if (kind == '<') {
            char next = At(pattern, i + 3);
            state.Open(state.Multiline, state.DotAll);
            if (next is '=' or '!') {
                text.Append(pattern, i, 4);
                return i + 4;
            }
            int close = pattern.IndexOf('>', i + 3);
            if (close < 0) throw new ArgumentException("The group name is not closed.");
            RequireGroupName(pattern.Substring(i + 3, close - i - 3));
            text.Append("(?<").Append((++state.Captures).ToString(CultureInfo.InvariantCulture)).Append('>');
            return close + 1;
        }
        // Modifier groups such as (?i:...) scope flags to the group; global (?i) flags are not ECMA-262. .NET has no
        // equivalent for the ECMA-262 meaning of m and s, so they change how ^, $ and . translate inside the group.
        int j = i + 2;
        bool removing = false, multiline = state.Multiline, dotAll = state.DotAll;
        string ignoreCase = string.Empty;
        while (j < pattern.Length && (pattern[j] is 'i' or 'm' or 's' or '-')) {
            switch (pattern[j]) {
                case '-': removing = true; break;
                case 'm': multiline = !removing; break;
                case 's': dotAll = !removing; break;
                default: ignoreCase = removing ? "-i" : "i"; break;
            }
            j++;
        }
        if (j > i + 2 && At(pattern, j) == ':') {
            state.Open(multiline, dotAll);
            text.Append("(?").Append(ignoreCase).Append(':');
            return j + 1;
        }
        throw new ArgumentException("'(?" + kind + "' is not ECMA-262 group syntax.");
    }

    private static int TranslateEscape(string pattern, int i, bool inClass, State state) {
        StringBuilder text = state.Text;
        if (i + 1 >= pattern.Length) throw new ArgumentException("The pattern ends with a backslash.");
        char n = pattern[i + 1];
        switch (n) {
            case 'd': AppendSet(text, DigitRanges, inClass); return i + 2;
            case 'D': AppendSet(text, NotDigitRanges, inClass); return i + 2;
            case 'w': AppendSet(text, WordRanges, inClass); return i + 2;
            case 'W': AppendSet(text, NotWordRanges, inClass); return i + 2;
            case 's': AppendSet(text, SpaceRanges, inClass); return i + 2;
            case 'S': AppendSet(text, NotSpaceRanges, inClass); return i + 2;
            case 'p':
            case 'P': {
                int close = At(pattern, i + 2) == '{' ? pattern.IndexOf('}', i + 3) : -1;
                if (close < 0) throw new ArgumentException("\\" + n + " needs a property name in braces.");
                string name = pattern.Substring(i + 3, close - i - 3);
                text.Append('\\').Append(n).Append('{').Append(MapProperty(name)).Append('}');
                return close + 1;
            }
            case 'u': {
                if (At(pattern, i + 2) == '{') {
                    int close = pattern.IndexOf('}', i + 3);
                    if (close < 0 || !int.TryParse(pattern.Substring(i + 3, close - i - 3), NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out int codePoint) || codePoint > 0x10FFFF) {
                        throw new ArgumentException("Invalid \\u{...} escape.");
                    }
                    AppendCodePoint(text, codePoint, inClass);
                    return close + 1;
                }
                if (!TryHex(pattern, i + 2, 4, out int unit)) throw new ArgumentException("Invalid \\u escape.");
                // A surrogate pair written as two escapes is one code point.
                if (char.IsHighSurrogate((char)unit) && At(pattern, i + 6) == '\\' && At(pattern, i + 7) == 'u' && TryHex(pattern, i + 8, 4, out int low) && char.IsLowSurrogate((char)low)) {
                    AppendCodePoint(text, char.ConvertToUtf32((char)unit, (char)low), inClass);
                    return i + 12;
                }
                text.Append(pattern, i, 6);
                return i + 6;
            }
            case 'x':
                if (!TryHex(pattern, i + 2, 2, out _)) throw new ArgumentException("Invalid \\x escape.");
                text.Append(pattern, i, 4);
                return i + 4;
            case 'c':
                if (!IsAsciiLetter(At(pattern, i + 2))) throw new ArgumentException("\\c needs a control letter.");
                text.Append(pattern, i, 3);
                return i + 3;
            case 'b':
                // Word boundaries use the same ASCII word set as \w; .NET's \b would use Unicode letters.
                text.Append(inClass ? "\\u0008" : WordBoundary);
                return i + 2;
            case 'B':
                if (inClass) throw new ArgumentException("\\B is not valid in a character class.");
                text.Append(NotWordBoundary);
                return i + 2;
            case 'k': {
                if (At(pattern, i + 2) != '<') throw new ArgumentException("\\k needs a group name.");
                int close = pattern.IndexOf('>', i + 3);
                if (close < 0) throw new ArgumentException("\\k needs a group name.");
                string name = RequireGroupName(pattern.Substring(i + 3, close - i - 3));
                if (!state.Names.TryGetValue(name, out List<int>? numbers)) throw new ArgumentException("\\k<" + name + "> names no group.");
                AppendBackreference(text, numbers);
                return close + 1;
            }
            case 't':
            case 'n':
            case 'v':
            case 'f':
            case 'r':
            case '0':
                text.Append('\\').Append(n);
                return i + 2;
            default:
                if (n >= '1' && n <= '9') {
                    int end = i + 2;
                    while (end < pattern.Length && pattern[end] >= '0' && pattern[end] <= '9') end++;
                    if (!int.TryParse(pattern.Substring(i + 1, end - i - 1), NumberStyles.None, CultureInfo.InvariantCulture, out int group) || group > state.TotalCaptures) {
                        throw new ArgumentException("\\" + pattern.Substring(i + 1, end - i - 1) + " refers to no group.");
                    }
                    AppendBackreference(text, new List<int> { group });
                    return end;
                }
                if (IsAsciiLetter(n)) throw new ArgumentException("\\" + n + " is not an ECMA-262 escape.");
                text.Append('\\').Append(n);
                return i + 2;
        }
    }

    /// <summary>
    /// Checks a group name against the ECMA-262 identifier grammar: a letter, <c>$</c> or <c>_</c>, then letters,
    /// digits, marks, connector punctuation, <c>$</c>, ZWNJ or ZWJ.
    /// </summary>
    private static string RequireGroupName(string name) {
        if (name.Length == 0) throw new ArgumentException("A capture group needs a name.");
        for (int i = 0; i < name.Length; i++) {
            UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(name, i);
            bool start = category is UnicodeCategory.UppercaseLetter or UnicodeCategory.LowercaseLetter or UnicodeCategory.TitlecaseLetter
                or UnicodeCategory.ModifierLetter or UnicodeCategory.OtherLetter or UnicodeCategory.LetterNumber || name[i] is '$' or '_';
            bool part = start || category is UnicodeCategory.NonSpacingMark or UnicodeCategory.SpacingCombiningMark
                or UnicodeCategory.DecimalDigitNumber or UnicodeCategory.ConnectorPunctuation || name[i] is '\u200C' or '\u200D';
            if (!(i == 0 ? start : part)) throw new ArgumentException("'" + name + "' is not a valid capture group name.");
            if (char.IsHighSurrogate(name[i]) && i + 1 < name.Length && char.IsLowSurrogate(name[i + 1])) i++;
        }
        return name;
    }

    /// <summary>
    /// Appends a backreference to whichever of the groups captured. A reference to a group that captured nothing matches
    /// the empty string in ECMA-262, while .NET fails, so each reference is conditional on its group.
    /// </summary>
    private static void AppendBackreference(StringBuilder text, List<int> groups) {
        foreach (int group in groups) {
            string number = group.ToString(CultureInfo.InvariantCulture);
            text.Append("(?(").Append(number).Append(")\\k<").Append(number).Append(">|");
        }
        for (int i = 0; i < groups.Count; i++) text.Append(')');
    }

    /// <summary>Translation state: the output, the m and s flags of each open group, and capture numbers by name.</summary>
    private sealed class State {
        private readonly Stack<bool> _scopes = new Stack<bool>();

        public State(string pattern) {
            Text = new StringBuilder(pattern.Length + 16);
            Names = NumberNamedGroups(pattern, out int total);
            TotalCaptures = total;
        }

        public StringBuilder Text { get; }
        public Dictionary<string, List<int>> Names { get; }
        public int TotalCaptures { get; }
        public int Captures { get; set; }
        public bool Multiline { get; private set; }
        public bool DotAll { get; private set; }

        public void Open(bool multiline, bool dotAll) {
            _scopes.Push(Multiline);
            _scopes.Push(DotAll);
            Multiline = multiline;
            DotAll = dotAll;
        }

        public void Close() {
            if (_scopes.Count == 0) return;
            DotAll = _scopes.Pop();
            Multiline = _scopes.Pop();
        }

        /// <summary>Gives each group name its ECMA-262 capture number: groups count left to right, named or not.</summary>
        private static Dictionary<string, List<int>> NumberNamedGroups(string pattern, out int captures) {
            var names = new Dictionary<string, List<int>>(StringComparer.Ordinal);
            captures = 0;
            bool inClass = false;
            for (int i = 0; i < pattern.Length; i++) {
                char c = pattern[i];
                if (c == '\\') {
                    i++;
                } else if (inClass) {
                    if (c == ']') inClass = false;
                } else if (c == '[') {
                    inClass = true;
                } else if (c == '(') {
                    if (At(pattern, i + 1) != '?') {
                        captures++;
                    } else if (At(pattern, i + 2) == '<' && At(pattern, i + 3) is not ('=' or '!')) {
                        captures++;
                        int close = pattern.IndexOf('>', i + 3);
                        if (close < 0) continue;
                        // A name can repeat in other alternatives; each group keeps its own number.
                        string name = pattern.Substring(i + 3, close - i - 3);
                        if (!names.TryGetValue(name, out List<int>? numbers)) names.Add(name, numbers = new List<int>());
                        numbers.Add(captures);
                    }
                }
            }
            return names;
        }
    }

    private static string MapProperty(string name) {
        int equals = name.IndexOf('=');
        if (equals >= 0) {
            string property = name.Substring(0, equals), value = name.Substring(equals + 1);
            if (property is "General_Category" or "gc") return Properties.TryGetValue(value, out string? category) ? category : value;
            // .NET has Unicode blocks rather than scripts; a script maps to the block of the same name.
            if (property is "Script" or "sc" or "Script_Extensions" or "scx") return "Is" + value;
            throw new ArgumentException("The Unicode property '" + property + "' is not supported.");
        }
        return Properties.TryGetValue(name, out string? mapped) ? mapped : name;
    }

    private static void AppendSet(StringBuilder text, string ranges, bool inClass) {
        if (inClass) text.Append(ranges);
        else text.Append('[').Append(ranges).Append(']');
    }

    private static void AppendCodePoint(StringBuilder text, int codePoint, bool inClass) {
        if (codePoint <= 0xFFFF) {
            text.Append("\\u").Append(codePoint.ToString("X4", CultureInfo.InvariantCulture));
            return;
        }
        string pair = char.ConvertFromUtf32(codePoint);
        if (inClass) throw new ArgumentException(AstralInClass);
        text.Append("(?:\\u").Append(((int)pair[0]).ToString("X4", CultureInfo.InvariantCulture))
            .Append("\\u").Append(((int)pair[1]).ToString("X4", CultureInfo.InvariantCulture)).Append(')');
    }

    private static bool TryHex(string pattern, int start, int length, out int value) {
        value = 0;
        if (start + length > pattern.Length) return false;
        return int.TryParse(pattern.Substring(start, length), NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out value);
    }

    private static char At(string pattern, int index) => index < pattern.Length ? pattern[index] : '\0';

    private static bool IsAsciiLetter(char c) => (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z');
}
