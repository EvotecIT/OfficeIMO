namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static partial class LegacyDocField {
        /// <summary>Reports omitted empty caches without widening the field import grammar.</summary>
        internal static ISet<int> ReportUnprojectedEmptyResults(
            IReadOnlyList<LegacyDocTextCharacter> characters, LegacyDocFib fib,
            Action<LegacyDocUnsupportedFeature> report, bool readableComments, bool readableBodyTextBox) {
            var reportedEnds = new HashSet<int>();
            for (int index = 0; index < characters.Count; index++) {
                if (characters[index].Character != Begin
                    || !TryReadField(characters, index, out string instruction, out int resultStart,
                        out int resultEnd, out int fieldEnd, allowEmptyResult: true)) continue;

                int beginPosition = characters[index].CharacterPosition;
                int endPosition = characters[fieldEnd].CharacterPosition;
                index = fieldEnd;
                if (resultStart != resultEnd) continue;
                string trimmed = instruction.Trim();
                if (IsInstruction(trimmed, "SECTIONPAGES")) continue;
                if (!IsPageNumberInstruction(instruction) && !IsNumberOfPagesInstruction(instruction)
                    && !TryReadDateTimeFieldKind(instruction, out _) && !IsDocumentPropertyInstruction(instruction)
                    && !IsTextReferenceInstruction(instruction) && !IsInstruction(trimmed, "EQ")
                    && !TryReadHyperlinkInstruction(instruction, out _)) continue;

                string? story = FindSupportedStory(beginPosition, endPosition, fib, readableComments, readableBodyTextBox);
                if (story == null) continue;
                string resultState = resultEnd > 0 && characters[resultEnd - 1].Character == Separator
                    ? "has a separator and an empty cached result"
                    : "has no result separator or cached result";
                reportedEnds.Add(endPosition);
                report(new LegacyDocUnsupportedFeature(LegacyDocUnsupportedFeatureKind.Formatting,
                    "doc-field-empty-result-unprojected",
                    $"The {story} field '{trimmed}' beginning at global character position {beginPosition} and ending at {endPosition} {resultState}. Its source instruction and empty result are not retained as an editable field by the current projection.",
                    "WordDocument", "Field:" + story + ":EmptyResult"));
            }
            return reportedEnds;
        }

        private static string? FindSupportedStory(int begin, int end, LegacyDocFib fib, bool readableComments, bool readableBodyTextBox) {
            long footnoteStart = fib.CcpText;
            long headerStart = footnoteStart + fib.CcpFtn;
            long commentStart = headerStart + fib.CcpHdd;
            long endnoteStart = commentStart + fib.CcpAtn;
            long textBoxStart = endnoteStart + fib.CcpEdn;
            if (begin >= 0 && end < footnoteStart) return "Main";
            if (begin >= footnoteStart && end < headerStart) return "Footnote";
            if (begin >= headerStart && end < commentStart) return "Header";
            if (readableComments && begin >= commentStart && end < endnoteStart) return "Comment";
            if (begin >= endnoteStart && end < textBoxStart) return "Endnote";
            if (readableBodyTextBox && begin >= textBoxStart && end < textBoxStart + fib.CcpTxbx) return "TextBox";
            return null;
        }
    }
}
