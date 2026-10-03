namespace OfficeIMO.Rtf;

internal static class RtfTextEncoding {
    // RTF fallbacks are bytes for readers that cannot decode \u. Large widths
    // have no practical interoperability value and amplify one model character
    // into attacker-controlled amounts of output.
    internal const int MaxUnicodeFallbackCharacterCount = 8;

    public static string EncodeText(string text) {
        return EncodeText(text, unicodeFallbackCharacterCount: 1);
    }

    public static string EncodeText(string text, int unicodeFallbackCharacterCount, bool useNamedCharacters = true) {
        if (unicodeFallbackCharacterCount < 0 || unicodeFallbackCharacterCount > MaxUnicodeFallbackCharacterCount)
            throw new ArgumentOutOfRangeException(nameof(unicodeFallbackCharacterCount),
                $"Unicode fallback character count must be between 0 and {MaxUnicodeFallbackCharacterCount}.");
        if (string.IsNullOrEmpty(text)) return string.Empty;

        var builder = new StringBuilder(text.Length);
        foreach (char ch in text) {
            if (!useNamedCharacters && ch > 0x7F) {
                AppendUnicodeCharacter(builder, ch, unicodeFallbackCharacterCount);
                continue;
            }
            switch (ch) {
                case '\\':
                    builder.Append(@"\\");
                    break;
                case '{':
                    builder.Append(@"\{");
                    break;
                case '}':
                    builder.Append(@"\}");
                    break;
                case '\t':
                    builder.Append(@"\tab ");
                    break;
                case '\n':
                    builder.Append(@"\line ");
                    break;
                case '\r':
                    break;
                case '\f':
                    builder.Append(@"\page ");
                    break;
                case '\v':
                    builder.Append(@"\column ");
                    break;
                case '\u00A0':
                    builder.Append(@"\~");
                    break;
                case '\u2011':
                    builder.Append(@"\_");
                    break;
                case '\u00AD':
                    builder.Append(@"\-");
                    break;
                case '\u2014':
                    AppendNamedCharacter(builder, "emdash");
                    break;
                case '\u2013':
                    AppendNamedCharacter(builder, "endash");
                    break;
                case '\u2003':
                    AppendNamedCharacter(builder, "emspace");
                    break;
                case '\u2002':
                    AppendNamedCharacter(builder, "enspace");
                    break;
                case '\u2005':
                    AppendNamedCharacter(builder, "qmspace");
                    break;
                case '\u2022':
                    AppendNamedCharacter(builder, "bullet");
                    break;
                case '\u2018':
                    AppendNamedCharacter(builder, "lquote");
                    break;
                case '\u2019':
                    AppendNamedCharacter(builder, "rquote");
                    break;
                case '\u201C':
                    AppendNamedCharacter(builder, "ldblquote");
                    break;
                case '\u201D':
                    AppendNamedCharacter(builder, "rdblquote");
                    break;
                case '\u200E':
                    AppendNamedCharacter(builder, "ltrmark");
                    break;
                case '\u200F':
                    AppendNamedCharacter(builder, "rtlmark");
                    break;
                case '\u200D':
                    AppendNamedCharacter(builder, "zwj");
                    break;
                case '\u200C':
                    AppendNamedCharacter(builder, "zwnj");
                    break;
                default:
                    if (ch <= 0x7F) {
                        builder.Append(ch);
                    } else {
                        AppendUnicodeCharacter(builder, ch, unicodeFallbackCharacterCount);
                    }
                    break;
            }
        }

        return builder.ToString();
    }

    private static void AppendNamedCharacter(StringBuilder builder, string controlWord) {
        builder.Append('\\');
        builder.Append(controlWord);
        builder.Append(' ');
    }

    private static void AppendUnicodeCharacter(StringBuilder builder, char character, int fallbackCount) {
        int value = character > short.MaxValue ? character - 65536 : character;
        builder.Append(@"\u");
        builder.Append(value.ToString(CultureInfo.InvariantCulture));
        if (fallbackCount == 0) builder.Append(' ');
        else for (int index = 0; index < fallbackCount; index++) builder.Append('?');
    }
}
