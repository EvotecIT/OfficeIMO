namespace OfficeIMO.Rtf;

internal static class RtfTextEncoding {
    public static string EncodeText(string text) {
        return EncodeText(text, unicodeFallbackCharacterCount: 1);
    }

    public static string EncodeText(string text, int unicodeFallbackCharacterCount) {
        if (unicodeFallbackCharacterCount < 0) throw new ArgumentOutOfRangeException(nameof(unicodeFallbackCharacterCount), "Unicode fallback character count cannot be negative.");
        if (string.IsNullOrEmpty(text)) return string.Empty;

        var builder = new StringBuilder(text.Length);
        foreach (char ch in text) {
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
                        int value = ch;
                        if (value > short.MaxValue) {
                            value -= 65536;
                        }

                        builder.Append(@"\u");
                        builder.Append(value.ToString(CultureInfo.InvariantCulture));
                        if (unicodeFallbackCharacterCount == 0) {
                            builder.Append(' ');
                        } else {
                            for (int i = 0; i < unicodeFallbackCharacterCount; i++) {
                                builder.Append('?');
                            }
                        }
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
}
