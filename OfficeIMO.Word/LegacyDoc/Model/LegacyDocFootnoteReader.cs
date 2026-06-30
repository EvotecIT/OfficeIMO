namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static class LegacyDocFootnoteReader {
        internal const char FootnoteReferenceCharacter = '\u0002';

        internal static bool HasReadableFootnoteTables(byte[] tableStream, LegacyDocFib fib) {
            return fib.CcpFtn == 0
                || TryReadFootnoteTables(tableStream, fib, out _, out _, out _);
        }

        internal static IReadOnlyList<LegacyDocFootnote> Read(byte[] tableStream, LegacyDocTextContent textContent, LegacyDocFib fib, out string? warning) {
            warning = null;
            if (fib.CcpFtn == 0) {
                return Array.Empty<LegacyDocFootnote>();
            }

            if (!TryReadFootnoteTables(tableStream, fib, out int[] referencePositions, out int[] textPositions, out warning)) {
                return Array.Empty<LegacyDocFootnote>();
            }

            int footnoteBaseCharacterPosition = fib.CcpText;
            int footnoteCount = referencePositions.Length;
            var footnotes = new List<LegacyDocFootnote>(footnoteCount);
            for (int index = 0; index < footnoteCount; index++) {
                int startCharacter = textPositions[index];
                int endCharacter = textPositions[index + 1];
                if (endCharacter <= startCharacter) {
                    continue;
                }

                string rawText = ExtractStoryText(
                    textContent.AllCharacters,
                    footnoteBaseCharacterPosition + startCharacter,
                    footnoteBaseCharacterPosition + endCharacter);
                IReadOnlyList<string> paragraphs = SplitStoryParagraphs(rawText);
                if (paragraphs.Count == 0) {
                    continue;
                }

                footnotes.Add(new LegacyDocFootnote(referencePositions[index], paragraphs));
            }

            return footnotes;
        }

        private static bool TryReadFootnoteTables(byte[] tableStream, LegacyDocFib fib, out int[] referencePositions, out int[] textPositions, out string? warning) {
            referencePositions = Array.Empty<int>();
            textPositions = Array.Empty<int>();
            warning = null;

            if (fib.CcpFtn == 0) {
                return true;
            }

            if (!TryReadFootnoteReferencePositions(tableStream, fib, out referencePositions, out warning)) {
                return false;
            }

            if (!TryReadFootnoteTextPositions(tableStream, fib, out textPositions, out warning)) {
                return false;
            }

            if (referencePositions.Length == 0 || textPositions.Length < referencePositions.Length + 1) {
                warning = "The footnote reference and text PLCs do not contain matching simple footnote ranges.";
                referencePositions = Array.Empty<int>();
                textPositions = Array.Empty<int>();
                return false;
            }

            textPositions = textPositions.Take(referencePositions.Length + 1).ToArray();
            int previousTextPosition = -1;
            for (int index = 0; index < textPositions.Length; index++) {
                int position = textPositions[index];
                if (position < previousTextPosition || position < 0 || position > fib.CcpFtn) {
                    warning = "The footnote text PLC contains a non-monotonic or out-of-range character position.";
                    referencePositions = Array.Empty<int>();
                    textPositions = Array.Empty<int>();
                    return false;
                }

                previousTextPosition = position;
            }

            return true;
        }

        private static bool TryReadFootnoteReferencePositions(byte[] tableStream, LegacyDocFib fib, out int[] positions, out string? warning) {
            positions = Array.Empty<int>();
            warning = null;
            if (fib.LcbPlcffndRef == 0) {
                warning = "The FIB reports footnote story text without a footnote reference PLC.";
                return false;
            }

            if (fib.FcPlcffndRef < 0
                || fib.LcbPlcffndRef < 4
                || fib.FcPlcffndRef + fib.LcbPlcffndRef > tableStream.Length
                || (fib.LcbPlcffndRef - 4) % 6 != 0) {
                warning = "The FIB points outside the selected table stream for the footnote reference PLC.";
                return false;
            }

            int noteCount = (fib.LcbPlcffndRef - 4) / 6;
            var cps = new int[noteCount + 1];
            for (int index = 0; index < cps.Length; index++) {
                cps[index] = LegacyDocFib.ReadInt32(tableStream, fib.FcPlcffndRef + (index * 4));
            }

            positions = cps.Take(noteCount).ToArray();
            return true;
        }

        private static bool TryReadFootnoteTextPositions(byte[] tableStream, LegacyDocFib fib, out int[] positions, out string? warning) {
            positions = Array.Empty<int>();
            warning = null;
            if (fib.LcbPlcffndTxt == 0) {
                warning = "The FIB reports footnote story text without a footnote text PLC.";
                return false;
            }

            if (fib.FcPlcffndTxt < 0
                || fib.LcbPlcffndTxt < 8
                || fib.FcPlcffndTxt + fib.LcbPlcffndTxt > tableStream.Length
                || fib.LcbPlcffndTxt % 4 != 0) {
                warning = "The FIB points outside the selected table stream for the footnote text PLC.";
                return false;
            }

            int positionCount = fib.LcbPlcffndTxt / 4;
            positions = new int[positionCount];
            for (int index = 0; index < positionCount; index++) {
                positions[index] = LegacyDocFib.ReadInt32(tableStream, fib.FcPlcffndTxt + (index * 4));
            }

            return true;
        }

        private static string ExtractStoryText(IReadOnlyList<LegacyDocTextCharacter> characters, int startCharacter, int endCharacter) {
            if (endCharacter <= startCharacter) {
                return string.Empty;
            }

            var builder = new System.Text.StringBuilder(endCharacter - startCharacter);
            foreach (LegacyDocTextCharacter character in characters) {
                if (character.CharacterPosition < startCharacter) {
                    continue;
                }

                if (character.CharacterPosition >= endCharacter) {
                    break;
                }

                builder.Append(character.Character == '\a' ? '\r' : character.Character);
            }

            return builder.ToString();
        }

        private static IReadOnlyList<string> SplitStoryParagraphs(string rawText) {
            if (string.IsNullOrEmpty(rawText)) {
                return Array.Empty<string>();
            }

            string text = rawText;
            if (text.EndsWith("\r", StringComparison.Ordinal)) {
                text = text.Substring(0, text.Length - 1);
            }

            if (text.Length == 0) {
                return Array.Empty<string>();
            }

            return text
                .Split(new[] { '\r' }, StringSplitOptions.None)
                .Select((paragraph, index) => index == 0 ? StripLeadingFootnoteReferenceMark(paragraph) : paragraph)
                .Where(paragraph => paragraph.Length > 0)
                .ToArray();
        }

        private static string StripLeadingFootnoteReferenceMark(string paragraph) {
            if (string.IsNullOrEmpty(paragraph) || paragraph[0] != FootnoteReferenceCharacter) {
                return paragraph;
            }

            return paragraph.Length > 1 && paragraph[1] == ' '
                ? paragraph.Substring(2)
                : paragraph.Substring(1);
        }
    }
}
