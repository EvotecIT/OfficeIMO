namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static class LegacyDocField {
        internal const char Begin = '\u0013';
        internal const char Separator = '\u0014';
        internal const char End = '\u0015';

        internal static bool TryReadHyperlink(
            IReadOnlyList<LegacyDocTextCharacter> characters,
            int startIndex,
            out LegacyDocHyperlinkTarget target,
            out int resultStartIndex,
            out int resultEndIndex,
            out int fieldEndIndex) {
            target = default;
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

                if (character == Begin || character == Separator || IsBodyBoundary(character) || !IsSupportedResultCharacter(character)) {
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

            if (!TryReadHyperlinkInstruction(instruction, out target)) {
                return false;
            }

            resultStartIndex = separatorIndex + 1;
            resultEndIndex = endIndex;
            fieldEndIndex = endIndex;
            return true;
        }

        internal static bool TryReadPageNumber(
            IReadOnlyList<LegacyDocTextCharacter> characters,
            int startIndex,
            out int resultStartIndex,
            out int resultEndIndex,
            out int fieldEndIndex) {
            resultStartIndex = -1;
            resultEndIndex = -1;
            fieldEndIndex = -1;

            if (!TryReadField(
                characters,
                startIndex,
                out string instruction,
                out resultStartIndex,
                out resultEndIndex,
                out fieldEndIndex)) {
                return false;
            }

            return IsPageNumberInstruction(instruction);
        }

        internal static bool TryReadNumberOfPages(
            IReadOnlyList<LegacyDocTextCharacter> characters,
            int startIndex,
            out int resultStartIndex,
            out int resultEndIndex,
            out int fieldEndIndex) {
            resultStartIndex = -1;
            resultEndIndex = -1;
            fieldEndIndex = -1;

            if (!TryReadField(
                characters,
                startIndex,
                out string instruction,
                out resultStartIndex,
                out resultEndIndex,
                out fieldEndIndex)) {
                return false;
            }

            return IsNumberOfPagesInstruction(instruction);
        }

        internal static bool TryReadDate(
            IReadOnlyList<LegacyDocTextCharacter> characters,
            int startIndex,
            out string instruction,
            out int resultStartIndex,
            out int resultEndIndex,
            out int fieldEndIndex) {
            instruction = string.Empty;
            resultStartIndex = -1;
            resultEndIndex = -1;
            fieldEndIndex = -1;

            if (!TryReadField(
                characters,
                startIndex,
                out instruction,
                out resultStartIndex,
                out resultEndIndex,
                out fieldEndIndex)) {
                return false;
            }

            return IsDateInstruction(instruction);
        }

        private static bool TryReadHyperlinkInstruction(string instruction, out LegacyDocHyperlinkTarget target) {
            target = default;
            string trimmed = instruction.Trim();
            if (!trimmed.StartsWith("HYPERLINK", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            int position = "HYPERLINK".Length;
            if (position < trimmed.Length && !char.IsWhiteSpace(trimmed[position])) {
                return false;
            }

            int anchorSwitch = IndexOfHyperlinkAnchorSwitch(trimmed, position);
            if (anchorSwitch >= 0) {
                int anchorPosition = anchorSwitch + 2;
                if (!TryReadQuotedValue(trimmed, anchorPosition, out string? anchor)) {
                    return false;
                }

                target = LegacyDocHyperlinkTarget.ForAnchor(anchor);
                return true;
            }

            if (!TryReadQuotedValue(trimmed, position, out string? uriText)) {
                return false;
            }

            if (!Uri.TryCreate(uriText, UriKind.Absolute, out Uri? parsed) || string.IsNullOrEmpty(parsed.Scheme)) {
                return false;
            }

            target = LegacyDocHyperlinkTarget.ForUri(parsed.ToString());
            return true;
        }

        private static bool TryReadField(
            IReadOnlyList<LegacyDocTextCharacter> characters,
            int startIndex,
            out string instruction,
            out int resultStartIndex,
            out int resultEndIndex,
            out int fieldEndIndex) {
            instruction = string.Empty;
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

                if (character == Begin || character == Separator || IsBodyBoundary(character) || !IsSupportedResultCharacter(character)) {
                    return false;
                }
            }

            if (endIndex <= separatorIndex + 1) {
                return false;
            }

            instruction = new string(characters
                .Skip(startIndex + 1)
                .Take(separatorIndex - startIndex - 1)
                .Select(character => character.Character)
                .ToArray());

            resultStartIndex = separatorIndex + 1;
            resultEndIndex = endIndex;
            fieldEndIndex = endIndex;
            return true;
        }

        private static bool IsPageNumberInstruction(string instruction) {
            string trimmed = instruction.Trim();
            if (!trimmed.StartsWith("PAGE", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            return trimmed.Length == "PAGE".Length
                || char.IsWhiteSpace(trimmed["PAGE".Length]);
        }

        private static bool IsNumberOfPagesInstruction(string instruction) {
            string trimmed = instruction.Trim();
            if (!trimmed.StartsWith("NUMPAGES", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            return trimmed.Length == "NUMPAGES".Length
                || char.IsWhiteSpace(trimmed["NUMPAGES".Length]);
        }

        private static bool IsDateInstruction(string instruction) {
            string trimmed = instruction.Trim();
            if (!trimmed.StartsWith("DATE", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            return trimmed.Length == "DATE".Length
                || char.IsWhiteSpace(trimmed["DATE".Length]);
        }

        private static int IndexOfHyperlinkAnchorSwitch(string instruction, int startIndex) {
            bool inQuotedText = false;
            bool escaped = false;
            for (int index = startIndex; index < instruction.Length - 1; index++) {
                char current = instruction[index];
                if (escaped) {
                    escaped = false;
                    continue;
                }

                if (current == '\\' && inQuotedText) {
                    escaped = true;
                    continue;
                }

                if (current == '"') {
                    inQuotedText = !inQuotedText;
                    continue;
                }

                if (inQuotedText) {
                    continue;
                }

                if (instruction[index] != '\\' || char.ToUpperInvariant(instruction[index + 1]) != 'L') {
                    continue;
                }

                bool hasTokenStart = index == 0 || char.IsWhiteSpace(instruction[index - 1]);
                bool hasTokenEnd = index + 2 >= instruction.Length || char.IsWhiteSpace(instruction[index + 2]);
                if (hasTokenStart && hasTokenEnd) {
                    return index;
                }
            }

            return -1;
        }

        private static bool TryReadQuotedValue(string text, int startIndex, out string value) {
            value = string.Empty;
            int quoteStart = text.IndexOf('"', startIndex);
            if (quoteStart < 0) {
                return false;
            }

            var builder = new System.Text.StringBuilder();
            bool escaped = false;
            for (int index = quoteStart + 1; index < text.Length; index++) {
                char character = text[index];
                if (escaped) {
                    builder.Append(character);
                    escaped = false;
                    continue;
                }

                if (character == '\\') {
                    escaped = true;
                    continue;
                }

                if (character == '"') {
                    value = builder.ToString();
                    return !string.IsNullOrWhiteSpace(value);
                }

                builder.Append(character);
            }

            return false;
        }

        private static bool IsBodyBoundary(char character) {
            return character == '\r' || character == '\n' || character == '\a';
        }

        private static bool IsSupportedResultCharacter(char character) {
            return !char.IsControl(character)
                || LegacyDocSpecialCharacters.IsSupportedInlineControl(character);
        }
    }
}
