namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static class LegacyDocField {
        internal const char Begin = '\u0013';
        internal const char Separator = '\u0014';
        internal const char End = '\u0015';

        internal static bool TryReadExternalHyperlink(
            IReadOnlyList<LegacyDocTextCharacter> characters,
            int startIndex,
            out string? uri,
            out int resultStartIndex,
            out int resultEndIndex,
            out int fieldEndIndex) {
            uri = null;
            resultStartIndex = -1;
            resultEndIndex = -1;
            fieldEndIndex = -1;

            if (startIndex < 0
                || startIndex >= characters.Count
                || characters[startIndex].Character != Begin) {
                return false;
            }

            int separatorIndex = -1;
            for (int index = startIndex + 1; index < characters.Count; index++) {
                char character = characters[index].Character;
                if (character == Separator) {
                    separatorIndex = index;
                    break;
                }

                if (character == Begin || character == End || IsBodyBoundary(character)) {
                    return false;
                }
            }

            if (separatorIndex < 0) {
                return false;
            }

            int endIndex = -1;
            for (int index = separatorIndex + 1; index < characters.Count; index++) {
                char character = characters[index].Character;
                if (character == End) {
                    endIndex = index;
                    break;
                }

                if (character == Begin || character == Separator || IsBodyBoundary(character) || char.IsControl(character)) {
                    return false;
                }
            }

            if (endIndex <= separatorIndex + 1) {
                return false;
            }

            string instruction = new string(characters
                .Skip(startIndex + 1)
                .Take(separatorIndex - startIndex - 1)
                .Select(character => character.Character)
                .ToArray());

            if (!TryReadExternalHyperlinkInstruction(instruction, out uri)) {
                return false;
            }

            resultStartIndex = separatorIndex + 1;
            resultEndIndex = endIndex;
            fieldEndIndex = endIndex;
            return true;
        }

        private static bool TryReadExternalHyperlinkInstruction(string instruction, out string? uri) {
            uri = null;
            string trimmed = instruction.Trim();
            if (!trimmed.StartsWith("HYPERLINK", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            int position = "HYPERLINK".Length;
            if (position < trimmed.Length && !char.IsWhiteSpace(trimmed[position])) {
                return false;
            }

            int quoteStart = trimmed.IndexOf('"', position);
            int quoteEnd = quoteStart >= 0 ? trimmed.IndexOf('"', quoteStart + 1) : -1;
            if (quoteStart < 0 || quoteEnd <= quoteStart + 1) {
                return false;
            }

            string target = trimmed.Substring(quoteStart + 1, quoteEnd - quoteStart - 1);
            if (!Uri.TryCreate(target, UriKind.Absolute, out Uri? parsed) || string.IsNullOrEmpty(parsed.Scheme)) {
                return false;
            }

            uri = parsed.ToString();
            return true;
        }

        private static bool IsBodyBoundary(char character) {
            return character == '\r' || character == '\n' || character == '\a';
        }
    }
}
