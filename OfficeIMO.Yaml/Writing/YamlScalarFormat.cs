namespace OfficeIMO.Yaml;

/// <summary>
/// Decides which scalar styles can represent a value exactly, and formats typed values in spellings that every
/// supported schema reads the same way. The emitter and the model's style inference share these rules.
/// </summary>
internal static class YamlScalarFormat {
#if NET8_0_OR_GREATER
    private static readonly SearchValues<char> s_plainEverywhere = SearchValues.Create(
        " !\"$%&'()*+-./0123456789;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ\\^_`abcdefghijklmnopqrstuvwxyz|~");
#endif

    /// <summary>The longest scalar text written as an implicit key; YAML limits implicit keys to 1024 characters.</summary>
    public const int MaxImplicitKeyLength = 1000;

    /// <summary>
    /// Returns whether the value can be written as a plain scalar on one line in the given context. Plain scalars never
    /// span lines here, because the writer does not fold.
    /// </summary>
    public static bool CanBePlain(ReadOnlySpan<char> value, bool flow, bool key = false) {
        if (value.Length == 0) return false;
        char first = value[0];
        char last = value[value.Length - 1];
        // A final ':' would read as a value indicator, except in a key, where the key's own ':' follows it.
        if (first == ' ' || first == '\t' || last == ' ' || last == '\t' || (last == ':' && !key)) return false;
        if (YamlChars.IsIndicator(first)) {
            // '-', '?' and ':' start a plain scalar only before a "safe" character.
            if (first is not ('-' or '?' or ':')) return false;
            // A one-character key is followed by its own ':', which keeps the indicator part of the scalar.
            if (value.Length == 1) return key;
            char next = value[1];
            if (next == ' ' || next == '\t' || (flow && YamlChars.IsFlowIndicator(next))) return false;
        }
        // Document markers at the start of a line would end the document.
        if (value.Length >= 3 && ((first == '-' && value[1] == '-' && value[2] == '-') || (first == '.' && value[1] == '.' && value[2] == '.'))
            && (value.Length == 3 || value[3] == ' ' || value[3] == '\t')) return false;
#if NET8_0_OR_GREATER
        // Printable ASCII without ':', '#' and the flow indicators is plain in any context, so only the other
        // characters (a ':' in an image reference, a '#' in a fragment) go through the rules below.
        int i = value.IndexOfAnyExcept(s_plainEverywhere);
        while (i >= 0) {
            if (!IsPlainCharacter(value, i, flow)) return false;
            int next = value.Slice(i + 1).IndexOfAnyExcept(s_plainEverywhere);
            if (next < 0) break;
            i += next + 1;
        }
#else
        for (int i = 0; i < value.Length; i++) {
            if (!IsPlainCharacter(value, i, flow)) return false;
        }
#endif
        return true;
    }

    /// <summary>Whether the character at <paramref name="i"/> (not the first) keeps a single-line plain scalar plain.</summary>
    private static bool IsPlainCharacter(ReadOnlySpan<char> value, int i, bool flow) {
        char c = value[i];
        if (!IsPrintableInline(c)) return false;
        if (c == ':') {
            return !(i + 1 < value.Length && (value[i + 1] == ' ' || value[i + 1] == '\t' || (flow && YamlChars.IsFlowIndicator(value[i + 1]))));
        }
        if (c == '#') return !(value[i - 1] == ' ' || value[i - 1] == '\t');
        return !(flow && YamlChars.IsFlowIndicator(c));
    }

    /// <summary>
    /// Returns whether the value can be a plain scalar spread over lines, each line break written as an empty line:
    /// every line is non-empty plain text without surrounding blanks and continuation lines cannot start a comment.
    /// </summary>
    public static bool CanBePlainMultiLine(ReadOnlySpan<char> value, bool flow) {
        if (value.Length == 0 || value[0] == '\n' || value[value.Length - 1] == '\n' || value.IndexOf('\r') >= 0) return false;
        bool first = true;
        while (true) {
            int end = value.IndexOf('\n');
            ReadOnlySpan<char> line = end < 0 ? value : value.Slice(0, end);
            if (line.Length > 0) {
                if (first) {
                    if (!CanBePlain(line, flow)) return false;
                } else if (line[0] == '#' || !CanBePlain(line, flow)) {
                    return false;
                }
                first = false;
            }
            if (end < 0) return true;
            value = value.Slice(end + 1);
        }
    }

    /// <summary>
    /// Returns whether a plain spelling of the text would be read as anything other than a string by the YAML 1.2 core
    /// schema or by the YAML 1.1 readers in common use (<c>yes</c>, <c>on</c>, <c>~</c>, <c>1.5</c>, <c>012</c>,
    /// <c>1:20</c>, <c>2001-12-14</c>, <c>&lt;&lt;</c>...). Versions such as <c>1.29.2</c> stay plain: those readers take
    /// them as strings.
    /// </summary>
    public static bool IsSchemaAmbiguous(ReadOnlySpan<char> value) {
        if (StartsAsText(value)) return false;
        switch (StartsAsNumber(value)) {
            case NumberShape.Digits:
                return true;
            case NumberShape.Text:
                return false;
            default:
                return IsSchemaAmbiguousResolved(value);
        }
    }

    /// <summary>The definition the shortcuts in <see cref="IsSchemaAmbiguous"/> must agree with: full resolution under both schemas.</summary>
    internal static bool IsSchemaAmbiguousResolved(ReadOnlySpan<char> value) =>
        YamlSchemaResolver.Resolve(value, YamlSchema.Core) != YamlScalarKind.String
        || YamlSchemaResolver.ReadsAsNonString11(value);

    private enum NumberShape : byte { Other, Digits, Text }

#if NET8_0_OR_GREATER
    private static readonly SearchValues<char> s_numberCharacters = SearchValues.Create(NumberCharacters);
#endif

    /// <summary>
    /// Every character a number, timestamp, infinity or NaN of either schema can contain once it starts with a digit, a
    /// sign or a dot: digits, signs, '.', '_', ',', ':', the hexadecimal digits, the base prefixes, the timestamp
    /// separators and the letters of <c>.inf</c> and <c>.nan</c>.
    /// </summary>
    private const string NumberCharacters = "0123456789+-._,: \tabcdefABCDEFxXoOtTZinIN";

    /// <summary>
    /// Classifies text that starts with a digit, a sign or a dot without the full resolution: decimal digits only are
    /// a core integer, and a character no number or timestamp contains (the <c>m</c> of <c>100m</c>, the <c>Mi</c> of
    /// <c>128Mi</c>) makes it a string in every schema. Anything else needs the resolver.
    /// </summary>
    private static NumberShape StartsAsNumber(ReadOnlySpan<char> value) {
        if (value.Length == 0) return NumberShape.Other;
        char first = value[0];
        if (!IsDigit(first) && first is not ('-' or '+' or '.')) return NumberShape.Other;
#if NET8_0_OR_GREATER
        if (value.IndexOfAnyExceptInRange('0', '9') < 0) return NumberShape.Digits;
        return value.IndexOfAnyExcept(s_numberCharacters) >= 0 ? NumberShape.Text : NumberShape.Other;
#else
        bool digits = true;
        for (int i = 0; i < value.Length; i++) {
            char c = value[i];
            if (IsDigit(c)) continue;
            digits = false;
            if (NumberCharacters.IndexOf(c) < 0) return NumberShape.Text;
        }
        return digits ? NumberShape.Digits : NumberShape.Other;
#endif
    }

    /// <summary>
    /// Whether the text is a string in every schema without the full resolution. Numbers, timestamps, merge keys and
    /// <c>~</c> start with a digit, a dot, a symbol, or a sign followed by a digit, a dot or an underscore; the null and boolean words
    /// are the 1.1 spellings of <c>null</c>, <c>true</c>, <c>false</c>, <c>yes</c>, <c>no</c>, <c>on</c>,
    /// <c>off</c>, <c>y</c> and <c>n</c>. Most names and values in real files pass, so only text that starts like a
    /// number is resolved.
    /// </summary>
    private static bool StartsAsText(ReadOnlySpan<char> value) {
        if (value.Length == 0) return false;
        char first = value[0];
        // go-yaml v2 drops underscores before parsing 1.1 numbers, so '+_1.' is a float there.
        if (first is '-' or '+') return value.Length > 1 && !IsDigit(value[1]) && value[1] is not ('.' or '_');
        char c = (char)(first | 0x20);
        if (c < 'a' || c > 'z') return false;
        return value.Length > 5 || c is not ('n' or 'y' or 'o' or 't' or 'f') || !IsNullOrBooleanWord(value);
    }

    private static bool IsDigit(char c) => c >= '0' && c <= '9';

    /// <summary>The null and boolean words of the YAML 1.2 core and 1.1 schemas, in their three casings.</summary>
    private static bool IsNullOrBooleanWord(ReadOnlySpan<char> value) => value.Length switch {
        1 => value[0] is 'y' or 'Y' or 'n' or 'N',
        2 => IsWord(value, "no") || IsWord(value, "on"),
        3 => IsWord(value, "yes") || IsWord(value, "off"),
        4 => IsWord(value, "null") || IsWord(value, "true"),
        5 => IsWord(value, "false"),
        _ => false
    };

    /// <summary>Whether the text is the lower-case word, the word capitalized, or the word in capitals.</summary>
    private static bool IsWord(ReadOnlySpan<char> value, string lower) {
        if (char.ToLowerInvariant(value[0]) != lower[0]) return false;
        bool upper = value.Length > 1 && value[1] >= 'A' && value[1] <= 'Z';
        if (upper && value[0] != char.ToUpperInvariant(lower[0])) return false;
        for (int i = 1; i < value.Length; i++) {
            char expected = upper ? char.ToUpperInvariant(lower[i]) : lower[i];
            if (value[i] != expected) return false;
        }
        return true;
    }

    /// <summary>Returns whether the value fits single quotes on one line: printable text without line breaks.</summary>
    public static bool CanBeSingleQuoted(ReadOnlySpan<char> value) {
        for (int i = 0; i < value.Length; i++) {
            if (!IsPrintableInline(value[i])) return false;
        }
        return true;
    }

    /// <summary>Returns whether the value can be a literal or folded block scalar: printable text whose only breaks are <c>\n</c>.</summary>
    public static bool CanBeBlock(ReadOnlySpan<char> value) {
        for (int i = 0; i < value.Length; i++) {
            char c = value[i];
            if (c != '\n' && !IsPrintableInline(c)) return false;
        }
        return true;
    }

    /// <summary>
    /// Picks the style for a string: the requested style when it can represent the value in this context, otherwise the
    /// most conservative one that can. <see cref="YamlStyle.None"/> asks for the default: plain when no core or YAML 1.1
    /// reader could take the text for anything but a string, a literal block for multi-line text in block context,
    /// otherwise <paramref name="quote"/> quotes, or double quotes when the text needs escapes.
    /// </summary>
    public static YamlStyle Choose(string value, YamlStyle requested, bool flow, bool key, YamlStyle quote) {
        ReadOnlySpan<char> text = value.AsSpan();
        switch (requested) {
            case YamlStyle.Plain:
                if (text.Length == 0 || CanBePlain(text, flow, key)) return YamlStyle.Plain;
                if (!key && CanBePlainMultiLine(text, flow)) return YamlStyle.Plain;
                break;
            case YamlStyle.SingleQuoted:
                return CanBeSingleQuoted(text) ? YamlStyle.SingleQuoted : YamlStyle.DoubleQuoted;
            case YamlStyle.DoubleQuoted:
                return YamlStyle.DoubleQuoted;
            case YamlStyle.Literal:
            case YamlStyle.Folded:
                if (!flow && !key && CanBeBlock(text)) return requested;
                break;
            default:
                if (CanBePlain(text, flow, key) && !IsSchemaAmbiguous(text)) return YamlStyle.Plain;
                if (!flow && !key && text.IndexOf('\n') >= 0 && CanBeBlock(text)) return YamlStyle.Literal;
                break;
        }
        if (quote == YamlStyle.SingleQuoted && CanBeSingleQuoted(text)) return YamlStyle.SingleQuoted;
        return YamlStyle.DoubleQuoted;
    }

    /// <summary>Returns whether the value contains a line break.</summary>
    public static bool IsMultiLine(ReadOnlySpan<char> value) => value.IndexOf('\n') >= 0 || value.IndexOf('\r') >= 0;

    /// <summary>
    /// Printable characters that every reader treats as content on a line: tab, ASCII text and the non-ASCII printable
    /// ranges, excluding the byte-order mark and the characters YAML 1.1 treats as line breaks (U+0085, U+2028, U+2029).
    /// Surrogate pairs are printable; a lone surrogate is not.
    /// </summary>
    public static bool IsPrintableInline(char c) {
        if (c < 0x80) return c == '\t' || (c >= 0x20 && c != 0x7F);
        if (c < 0xA0) return false;
        if (c >= 0xD800 && c <= 0xDFFF) return true;
        return c != '\u2028' && c != '\u2029' && c != '\uFEFF' && c <= '\uFFFD';
    }

    /// <summary>Formats a 64-bit integer into a buffer of at least 20 characters.</summary>
    public static int Format(long value, Span<char> destination) {
#if NET8_0_OR_GREATER
        value.TryFormat(destination, out int written, default, CultureInfo.InvariantCulture);
        return written;
#else
        ulong magnitude = value < 0 ? (ulong)(-(value + 1)) + 1 : (ulong)value;
        int length = 0;
        do {
            destination[length++] = (char)('0' + (int)(magnitude % 10));
            magnitude /= 10;
        } while (magnitude != 0);
        if (value < 0) destination[length++] = '-';
        destination.Slice(0, length).Reverse();
        return length;
#endif
    }

    /// <summary>
    /// Formats a double so the YAML 1.2 core schema and the YAML 1.1 float type both read it back exactly: a decimal
    /// point is always present and exponents carry a sign (<c>1.0</c>, <c>1.5e-05</c>), and the special values are
    /// <c>.nan</c>, <c>.inf</c> and <c>-.inf</c>.
    /// </summary>
    public static string Format(double value) {
        if (double.IsNaN(value)) return ".nan";
        if (double.IsPositiveInfinity(value)) return ".inf";
        if (double.IsNegativeInfinity(value)) return "-.inf";
        return Shape(value.ToString("R", CultureInfo.InvariantCulture));
    }

    /// <summary>Formats a single-precision float with its own shortest round-trip digits, in the same shape as <see cref="Format(double)"/>.</summary>
    public static string Format(float value) {
        if (float.IsNaN(value)) return ".nan";
        if (float.IsPositiveInfinity(value)) return ".inf";
        if (float.IsNegativeInfinity(value)) return "-.inf";
        return Shape(value.ToString("R", CultureInfo.InvariantCulture));
    }

    private static string Shape(string text) {
        int exponent = text.IndexOf('E');
        string mantissa = exponent < 0 ? text : text.Substring(0, exponent);
        if (mantissa.IndexOf('.') < 0) mantissa += ".0";
        if (exponent < 0) return mantissa;
        string power = text.Substring(exponent + 1);
        if (power[0] != '-' && power[0] != '+') power = "+" + power;
        return mantissa + "e" + power;
    }

#if NET8_0_OR_GREATER
    /// <summary>Formats a double as <see cref="Format(double)"/> does, into a buffer of at least 40 characters, without allocating.</summary>
    public static int Format(double value, Span<char> destination) {
        if (!double.IsFinite(value)) {
            string special = Format(value);
            special.AsSpan().CopyTo(destination);
            return special.Length;
        }
        // On .NET Core 3.0 and later the default format is the shortest round-trip one, the same digits as "R",
        // through a faster path.
        value.TryFormat(destination, out int length, default, CultureInfo.InvariantCulture);
        return ShapeInPlace(destination, length);
    }

    /// <summary>Formats a float as <see cref="Format(float)"/> does, into a buffer of at least 40 characters, without allocating.</summary>
    public static int Format(float value, Span<char> destination) {
        if (!float.IsFinite(value)) {
            string special = Format(value);
            special.AsSpan().CopyTo(destination);
            return special.Length;
        }
        value.TryFormat(destination, out int length, default, CultureInfo.InvariantCulture);
        return ShapeInPlace(destination, length);
    }

    /// <summary>Applies the shape of <see cref="Shape"/> to formatted digits in place.</summary>
    private static int ShapeInPlace(Span<char> destination, int length) {
        int exponent = destination.Slice(0, length).IndexOf('E');
        int mantissaEnd = exponent < 0 ? length : exponent;
        Span<char> power = stackalloc char[8];
        int powerLength = 0;
        if (exponent >= 0) {
            ReadOnlySpan<char> digits = destination.Slice(exponent + 1, length - exponent - 1);
            if (digits[0] != '-' && digits[0] != '+') power[powerLength++] = '+';
            digits.CopyTo(power.Slice(powerLength));
            powerLength += digits.Length;
        }
        int position = mantissaEnd;
        if (destination.Slice(0, mantissaEnd).IndexOf('.') < 0) {
            destination[position++] = '.';
            destination[position++] = '0';
        }
        if (powerLength > 0) {
            destination[position++] = 'e';
            power.Slice(0, powerLength).CopyTo(destination.Slice(position));
            position += powerLength;
        }
        return position;
    }
#endif

    /// <summary>The spelling of a boolean in every schema.</summary>
    public static string Format(bool value) => value ? "true" : "false";

    /// <summary>
    /// Returns the shorthand or verbatim form of a resolved tag: <c>tag:yaml.org,2002:str</c> becomes <c>!!str</c>,
    /// local tags stay <c>!name</c>, and anything else is written verbatim as <c>!&lt;uri&gt;</c>. Characters a tag
    /// cannot hold are percent-encoded as UTF-8.
    /// </summary>
    public static string FormatTag(string tag, bool secondaryHandleRedefined) {
        if (tag.Length == 0) throw new ArgumentException("A tag cannot be empty.", nameof(tag));
        if (tag == "!") return "!";
        if (!secondaryHandleRedefined && tag.StartsWith(YamlSchemaResolver.TagPrefix, StringComparison.Ordinal) && tag.Length > YamlSchemaResolver.TagPrefix.Length) {
            return "!!" + Escape(tag, YamlSchemaResolver.TagPrefix.Length, shorthand: true);
        }
        if (tag[0] == '!' && tag.Length > 1 && tag[1] != '<') return "!" + Escape(tag, 1, shorthand: true);
        return "!<" + Escape(tag, 0, shorthand: false) + ">";
    }

    private static string Escape(string tag, int start, bool shorthand) {
        bool clean = true;
        for (int i = start; i < tag.Length && clean; i++) clean = shorthand ? YamlChars.IsTagChar(tag[i]) : YamlChars.IsUri(tag[i]) || tag[i] == '!';
        if (clean) return start == 0 ? tag : tag.Substring(start);
        var text = new StringBuilder(tag.Length + 8);
        byte[] utf8 = Encoding.UTF8.GetBytes(tag.Substring(start));
        int position = 0;
        while (position < utf8.Length) {
            byte b = utf8[position++];
            char c = (char)b;
            bool allowed = b < 0x80 && (shorthand ? YamlChars.IsTagChar(c) : YamlChars.IsUri(c) || c == '!');
            if (allowed) {
                text.Append(c);
            } else {
                text.Append('%').Append("0123456789ABCDEF"[b >> 4]).Append("0123456789ABCDEF"[b & 15]);
            }
        }
        return text.ToString();
    }

    /// <summary>Returns whether a name is a valid anchor: one or more printable characters other than spaces and flow indicators.</summary>
    public static bool IsValidAnchor(string name) {
        if (string.IsNullOrEmpty(name)) return false;
        foreach (char c in name) {
            if (c == ' ' || c == '\t' || YamlChars.IsFlowIndicator(c) || !IsPrintableInline(c)) return false;
        }
        return true;
    }
}
