namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static partial class DirectDataSetWorkbookWriter {
#if NET8_0_OR_GREATER
            private static readonly System.Buffers.SearchValues<char> XmlAttributeSpecialCharacters =
                System.Buffers.SearchValues.Create("&<>\"'");

            // XML text escapes and the controls removed by the existing sanitizer.
            // Tab, line feed and carriage return remain valid text characters.
            private static readonly System.Buffers.SearchValues<char> XmlTextSpecialCharacters =
                System.Buffers.SearchValues.Create("\u0000\u0001\u0002\u0003\u0004\u0005\u0006\u0007\u0008\u000B\u000C\u000E\u000F\u0010\u0011\u0012\u0013\u0014\u0015\u0016\u0017\u0018\u0019\u001A\u001B\u001C\u001D\u001E\u001F&<>");
#endif
            private static void AppendEscaped(StringBuilder builder, string value) {
                int escapeIndex = IndexOfXmlEscape(value);
                if (escapeIndex < 0) {
                    builder.Append(value);
                    return;
                }

                int start = 0;
                while (escapeIndex >= 0) {
                    if (escapeIndex > start) {
                        builder.Append(value, start, escapeIndex - start);
                    }

                    AppendEscapedCharacter(builder, value[escapeIndex]);
                    start = escapeIndex + 1;
                    escapeIndex = IndexOfXmlEscape(value, start);
                }

                if (start < value.Length) {
                    builder.Append(value, start, value.Length - start);
                }
            }

            private static void WriteEscaped(TextWriter writer, string value) {
                int escapeIndex = IndexOfXmlEscape(value);
                if (escapeIndex < 0) {
                    writer.Write(value);
                    return;
                }

                int start = 0;
                while (escapeIndex >= 0) {
                    if (escapeIndex > start) {
                        WriteSlice(writer, value, start, escapeIndex - start);
                    }

                    WriteEscapedCharacter(writer, value[escapeIndex]);
                    start = escapeIndex + 1;
                    escapeIndex = IndexOfXmlEscape(value, start);
                }

                if (start < value.Length) {
                    WriteSlice(writer, value, start, value.Length - start);
                }
            }

            private static void WriteSanitizedEscaped(TextWriter writer, string value) {
                int specialIndex = IndexOfXmlTextSpecial(value, 0);
                if (specialIndex < 0) {
                    writer.Write(value);
                    return;
                }

                int start = 0;
                while (specialIndex >= 0) {
                    if (specialIndex > start) {
                        WriteSlice(writer, value, start, specialIndex - start);
                    }

                    char current = value[specialIndex];
                    if (!IsInvalidXmlControl(current)) {
                        WriteEscapedTextCharacter(writer, current);
                    }

                    start = specialIndex + 1;
                    specialIndex = IndexOfXmlTextSpecial(value, start);
                    if (specialIndex >= 0 && specialIndex - start < 16) {
                        WriteSanitizedEscapedScalar(writer, value, start);
                        return;
                    }
                }

                if (start < value.Length) {
                    WriteSlice(writer, value, start, value.Length - start);
                }
            }

            // Dense markup benefits from the original single character loop.
            // Keep its segment writes and sanitization when escapes cluster.
            private static void WriteSanitizedEscapedScalar(TextWriter writer, string value, int start) {
                for (int index = start; index < value.Length; index++) {
                    char current = value[index];
                    if (!IsInvalidXmlControl(current) && !IsXmlTextEscape(current)) continue;
                    if (index > start) WriteSlice(writer, value, start, index - start);
                    if (!IsInvalidXmlControl(current)) WriteEscapedTextCharacter(writer, current);
                    start = index + 1;
                }

                if (start < value.Length) WriteSlice(writer, value, start, value.Length - start);
            }

            private static void WriteSlice(TextWriter writer, string value, int startIndex, int length) {
#if NET6_0_OR_GREATER
                writer.Write(value.AsSpan(startIndex, length));
#else
                writer.Write(value.Substring(startIndex, length));
#endif
            }

            private static int IndexOfXmlEscape(string value, int startIndex = 0) {
#if NET8_0_OR_GREATER
                // After an escape, inspect nearby characters before starting
                // another vector search. Markup often puts escapes close together.
                if (startIndex > 0) {
                    int prefixEnd = startIndex + Math.Min(16, value.Length - startIndex);
                    for (; startIndex < prefixEnd; startIndex++) {
                        if (IsXmlEscape(value[startIndex])) return startIndex;
                    }
                }

                int relativeIndex = value.AsSpan(startIndex).IndexOfAny(XmlAttributeSpecialCharacters);
                return relativeIndex < 0 ? -1 : startIndex + relativeIndex;
#else
                for (int i = startIndex; i < value.Length; i++) {
                    if (IsXmlEscape(value[i])) {
                        return i;
                    }
                }

                return -1;
#endif
            }

            private static int IndexOfXmlTextSpecial(string value, int startIndex) {
#if NET8_0_OR_GREATER
                int relativeIndex = value.AsSpan(startIndex).IndexOfAny(XmlTextSpecialCharacters);
                return relativeIndex < 0 ? -1 : startIndex + relativeIndex;
#else
                for (int i = startIndex; i < value.Length; i++) {
                    if (IsInvalidXmlControl(value[i]) || IsXmlTextEscape(value[i])) {
                        return i;
                    }
                }

                return -1;
#endif
            }

            private static bool IsInvalidXmlControl(char value)
                => value < 0x20 && value != '\t' && value != '\n' && value != '\r';

            private static bool IsXmlEscape(char value)
                => value is '&' or '<' or '>' or '"' or '\'';

            private static bool IsXmlTextEscape(char value)
                => value is '&' or '<' or '>';

            private static bool NeedsPreserveSpace(string value) {
                return value.Length > 0 && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[value.Length - 1]));
            }

            private static void AppendEscapedCharacter(StringBuilder builder, char value) {
                switch (value) {
                    case '&':
                        builder.Append("&amp;");
                        break;
                    case '<':
                        builder.Append("&lt;");
                        break;
                    case '>':
                        builder.Append("&gt;");
                        break;
                    case '"':
                        builder.Append("&quot;");
                        break;
                    case '\'':
                        builder.Append("&apos;");
                        break;
                }
            }

            private static void WriteEscapedCharacter(TextWriter writer, char value) {
                switch (value) {
                    case '&':
                        writer.Write("&amp;");
                        break;
                    case '<':
                        writer.Write("&lt;");
                        break;
                    case '>':
                        writer.Write("&gt;");
                        break;
                    case '"':
                        writer.Write("&quot;");
                        break;
                    case '\'':
                        writer.Write("&apos;");
                        break;
                }
            }

            private static void WriteEscapedTextCharacter(TextWriter writer, char value) {
                switch (value) {
                    case '&':
                        writer.Write("&amp;");
                        break;
                    case '<':
                        writer.Write("&lt;");
                        break;
                    case '>':
                        writer.Write("&gt;");
                        break;
                }
            }
        }
    }
}
