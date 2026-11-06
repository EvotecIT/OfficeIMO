namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static partial class LegacyDocField {
        /// <summary>Reads inert filename and style references while retaining their native instruction and cached text.</summary>
        internal static bool TryReadTextReferenceField(
            IReadOnlyList<LegacyDocTextCharacter> characters, int startIndex, out string instruction,
            out int resultStartIndex, out int resultEndIndex, out int fieldEndIndex) =>
            TryReadField(characters, startIndex, out instruction,
                out resultStartIndex, out resultEndIndex, out fieldEndIndex)
            && IsTextReferenceInstruction(instruction);

        /// <summary>Finds fields whose instructions would be reduced to cached text by the shared projection.</summary>
        internal static bool HasUnprojectedFields(IReadOnlyList<LegacyDocTextCharacter> characters) {
            for (int index = 0; index < characters.Count; index++) {
                if (characters[index].Character != Begin) continue;
                if (!TryReadField(characters, index, out string instruction, out int resultStart, out int resultEnd, out int end,
                    allowEmptyResult: true)) return true;
                string trimmed = instruction.Trim();
                if (resultStart == resultEnd && !IsInstruction(trimmed, "SECTIONPAGES")) return true;
                bool supported = IsPageNumberInstruction(instruction) || IsNumberOfPagesInstruction(instruction)
                    || IsInstruction(trimmed, "SECTIONPAGES") || TryReadDateTimeFieldKind(instruction, out _)
                    || IsDocumentPropertyInstruction(instruction) || IsTextReferenceInstruction(instruction)
                    || IsInstruction(trimmed, "EQ") || TryReadHyperlinkInstruction(instruction, out _);
                if (!supported) return true;
                index = end;
            }
            return false;
        }

        /// <summary>Reports style references whose source name has no editable paragraph-style target.</summary>
        internal static bool HasUnresolvedStyleReferences(IReadOnlyList<LegacyDocTextCharacter> characters, LegacyDocStyleSheet styles) {
            HashSet<string>? names = null;
            for (int index = 0; index < characters.Count; index++) {
                if (characters[index].Character != Begin) continue;
                if (!TryReadField(characters, index, out string instruction, out _, out _, out int end, allowEmptyResult: true)) continue;
                index = end;
                string trimmed = instruction.Trim();
                if (!IsInstruction(trimmed, "STYLEREF")) continue;
                int nameStart = "STYLEREF".Length;
                while (nameStart < trimmed.Length && char.IsWhiteSpace(trimmed[nameStart])) nameStart++;
                if (nameStart == trimmed.Length) return true;
                string name;
                if (trimmed[nameStart] == '"') {
                    if (!TryReadQuotedValue(trimmed, nameStart, out name)) return true;
                } else {
                    int nameEnd = nameStart;
                    while (nameEnd < trimmed.Length && !char.IsWhiteSpace(trimmed[nameEnd])) nameEnd++;
                    name = trimmed.Substring(nameStart, nameEnd - nameStart);
                }
                if (names == null) {
                    names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var style in styles.ParagraphStyles) {
                        string? builtIn = style.BuiltInStyle?.ToString();
                        // Built-in styles can use the current DOCX template's name. Treat a
                        // localized source name as unresolved until its alias is qualified.
                        if (builtIn == null || string.Equals(style.Name.Replace(" ", string.Empty),
                            builtIn, StringComparison.OrdinalIgnoreCase)) names.Add(style.Name);
                        if (builtIn != null) names.Add(builtIn);
                        if (builtIn?.StartsWith("Heading", StringComparison.Ordinal) == true) {
                            names.Add(builtIn.Substring("Heading".Length));
                        }
                    }
                }
                if (!names.Contains(name)) return true;
            }
            return false;
        }

        internal static bool IsTextReferenceInstruction(string instruction) {
            string trimmed = instruction.Trim();
            return IsInstruction(trimmed, "FILENAME")
                || (IsInstruction(trimmed, "STYLEREF") && HasDocPropertyNameToken(trimmed, "STYLEREF".Length));
        }
    }
}
