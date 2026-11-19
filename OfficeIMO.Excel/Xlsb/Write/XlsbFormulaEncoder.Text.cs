namespace OfficeIMO.Excel.Xlsb.Write {
    internal static partial class XlsbFormulaEncoder {
        private const char MaskedCharacter = 'x';

        /// <summary>
        /// Returns <paramref name="text"/> with the contents of string literals, quoted sheet names and bracketed
        /// structured-reference or workbook parts replaced by a letter, keeping every position. Scanners that look for
        /// operators, parentheses, commas, braces or function names read the mask and cut the original text.
        /// </summary>
        /// <remarks>
        /// A string literal doubles its quotes (<c>"a""b"</c>) and a quoted sheet name doubles its apostrophes
        /// (<c>'Bob''s'</c>). Inside brackets an apostrophe escapes the next character, as in <c>Sales['[Note']]</c>,
        /// and brackets nest (<c>Sales[[#Totals],[Amount]]</c>).
        /// </remarks>
        private static string MaskQuotedText(string text) {
            char[] mask = text.ToCharArray();
            int bracketDepth = 0;
            for (int index = 0; index < text.Length; index++) {
                char character = text[index];
                if (bracketDepth > 0) {
                    if (character == '\'' && index + 1 < text.Length) {
                        mask[index] = MaskedCharacter;
                        mask[++index] = MaskedCharacter;
                    } else if (character == '[') {
                        bracketDepth++;
                    } else if (character == ']') {
                        bracketDepth--;
                    } else {
                        mask[index] = MaskedCharacter;
                    }
                    continue;
                }
                if (character == '[') {
                    bracketDepth = 1;
                    continue;
                }
                if (character != '"' && character != '\'') continue;
                // Mask up to the closing delimiter; a doubled delimiter stays inside.
                int position = index + 1;
                while (position < text.Length) {
                    if (text[position] == character) {
                        if (position + 1 < text.Length && text[position + 1] == character) {
                            mask[position] = mask[position + 1] = MaskedCharacter;
                            position += 2;
                            continue;
                        }
                        break;
                    }
                    mask[position++] = MaskedCharacter;
                }
                index = position;
            }
            return new string(mask);
        }

        private static bool IsIdentifierCharacter(char character) =>
            char.IsLetterOrDigit(character) || character is '_' or '.' or '\\';
    }
}
