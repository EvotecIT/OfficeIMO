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
            int nestedDepth = 0;
            for (int index = separatorIndex + 1; index < characters.Count; index++) {
                char character = characters[index].Character;
                if (character == Begin) {
                    nestedDepth++;
                    continue;
                }
                if (character == End) {
                    if (nestedDepth > 0) {
                        nestedDepth--;
                        continue;
                    }
                    endIndex = index;
                    break;
                }
                if (character == Separator && nestedDepth > 0) {
                    continue;
                }

                if ((character == Separator && nestedDepth == 0) || IsBodyBoundary(character) || !IsSupportedResultCharacter(character)) {
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

        internal static IEnumerable<int> EnumerateVisibleResultIndexes(
            IReadOnlyList<LegacyDocTextCharacter> characters,
            int resultStartIndex,
            int resultEndIndex) {
            int index = resultStartIndex;
            while (index < resultEndIndex) {
                if (characters[index].Character == Begin &&
                    TryFindNestedFieldResult(characters, index, resultEndIndex, out int nestedResultStart, out int nestedResultEnd, out int nestedFieldEnd)) {
                    foreach (int nestedIndex in EnumerateVisibleResultIndexes(characters, nestedResultStart, nestedResultEnd)) {
                        yield return nestedIndex;
                    }
                    index = nestedFieldEnd + 1;
                    continue;
                }

                char character = characters[index].Character;
                if (character != Begin && character != Separator && character != End) {
                    yield return index;
                }
                index++;
            }
        }

        private static bool TryFindNestedFieldResult(
            IReadOnlyList<LegacyDocTextCharacter> characters,
            int fieldStartIndex,
            int rangeEndIndex,
            out int resultStartIndex,
            out int resultEndIndex,
            out int fieldEndIndex) {
            resultStartIndex = -1;
            resultEndIndex = -1;
            fieldEndIndex = -1;
            int separatorIndex = -1;
            int nestedDepth = 0;

            for (int index = fieldStartIndex + 1; index < rangeEndIndex; index++) {
                char character = characters[index].Character;
                if (character == Begin) {
                    nestedDepth++;
                    continue;
                }
                if (character == End) {
                    if (nestedDepth > 0) {
                        nestedDepth--;
                        continue;
                    }
                    if (separatorIndex < 0) return false;
                    resultStartIndex = separatorIndex + 1;
                    resultEndIndex = index;
                    fieldEndIndex = index;
                    return true;
                }
                if (character == Separator && nestedDepth == 0 && separatorIndex < 0) {
                    separatorIndex = index;
                }
            }

            return false;
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

        internal static bool TryReadDateTimeField(
            IReadOnlyList<LegacyDocTextCharacter> characters,
            int startIndex,
            out LegacyDocFieldKind fieldKind,
            out string instruction,
            out int resultStartIndex,
            out int resultEndIndex,
            out int fieldEndIndex) {
            fieldKind = LegacyDocFieldKind.None;
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

            return TryReadDateTimeFieldKind(instruction, out fieldKind);
        }

        internal static bool TryReadDocumentPropertyField(
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

            return IsDocumentPropertyInstruction(instruction);
        }

        internal static bool TryReadEquationField(
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

            return IsInstruction(instruction.Trim(), "EQ");
        }

        internal static bool TryReadDisplayField(
            IReadOnlyList<LegacyDocTextCharacter> characters,
            int startIndex,
            out int resultStartIndex,
            out int resultEndIndex,
            out int fieldEndIndex) {
            resultStartIndex = -1;
            resultEndIndex = -1;
            fieldEndIndex = -1;

            if (startIndex < 0
                || startIndex >= characters.Count
                || characters[startIndex].Character != Begin) {
                return false;
            }

            int separatorIndex = -1;
            int endIndex = -1;
            for (int index = startIndex + 1; index < characters.Count; index++) {
                char character = characters[index].Character;
                if (character == Separator) {
                    separatorIndex = index;
                    break;
                }

                if (character == End) {
                    endIndex = index;
                    break;
                }

                if (character == Begin || IsBodyBoundary(character)) {
                    return false;
                }
            }

            if (endIndex >= 0) {
                resultStartIndex = endIndex;
                resultEndIndex = endIndex;
                fieldEndIndex = endIndex;
                return true;
            }

            if (separatorIndex < 0) {
                return false;
            }

            for (int index = separatorIndex + 1; index < characters.Count; index++) {
                char character = characters[index].Character;
                if (character == End) {
                    resultStartIndex = separatorIndex + 1;
                    resultEndIndex = index;
                    fieldEndIndex = index;
                    return true;
                }

                if (character == Begin || character == Separator || IsBodyBoundary(character) || !IsSupportedResultCharacter(character)) {
                    return false;
                }
            }

            return false;
        }

        internal static bool IsDocumentPropertyInstruction(string instruction) {
            string trimmed = instruction.Trim();
            if (IsInstruction(trimmed, "DOCPROPERTY")) {
                return HasDocPropertyNameToken(trimmed, "DOCPROPERTY".Length);
            }

            return IsInstruction(trimmed, "AUTHOR")
                || IsInstruction(trimmed, "TITLE")
                || IsInstruction(trimmed, "SUBJECT")
                || IsInstruction(trimmed, "KEYWORDS")
                || IsInstruction(trimmed, "COMMENTS")
                || IsInstruction(trimmed, "LASTSAVEDBY")
                || IsInstruction(trimmed, "REVNUM");
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

        private static bool TryReadDateTimeFieldKind(string instruction, out LegacyDocFieldKind fieldKind) {
            fieldKind = LegacyDocFieldKind.None;
            string trimmed = instruction.Trim();
            if (IsInstruction(trimmed, "DATE")) {
                fieldKind = LegacyDocFieldKind.Date;
                return true;
            }

            if (IsInstruction(trimmed, "TIME")) {
                fieldKind = LegacyDocFieldKind.Time;
                return true;
            }

            if (IsInstruction(trimmed, "CREATEDATE")) {
                fieldKind = LegacyDocFieldKind.CreateDate;
                return true;
            }

            if (IsInstruction(trimmed, "SAVEDATE")) {
                fieldKind = LegacyDocFieldKind.SaveDate;
                return true;
            }

            if (IsInstruction(trimmed, "PRINTDATE")) {
                fieldKind = LegacyDocFieldKind.PrintDate;
                return true;
            }

            return false;
        }

        private static bool IsInstruction(string trimmedInstruction, string fieldName) {
            return trimmedInstruction.StartsWith(fieldName, StringComparison.OrdinalIgnoreCase)
                && (trimmedInstruction.Length == fieldName.Length || char.IsWhiteSpace(trimmedInstruction[fieldName.Length]));
        }

        private static bool HasDocPropertyNameToken(string instruction, int startIndex) {
            int index = startIndex;
            while (index < instruction.Length && char.IsWhiteSpace(instruction[index])) {
                index++;
            }

            if (index >= instruction.Length || instruction[index] == '\\') {
                return false;
            }

            if (instruction[index] == '"') {
                return TryReadQuotedValue(instruction, index, out string? value)
                    && !string.IsNullOrWhiteSpace(value);
            }

            return true;
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
