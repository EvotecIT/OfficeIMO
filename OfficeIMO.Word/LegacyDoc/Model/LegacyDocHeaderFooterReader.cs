using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static class LegacyDocHeaderFooterReader {
        private const int SeparatorStoryCount = 6;
        private const int StoriesPerSection = 6;

        internal static IReadOnlyList<LegacyDocHeaderFooterStory> Read(
            byte[] tableStream,
            LegacyDocTextContent textContent,
            LegacyDocFib fib,
            IReadOnlyList<LegacyDocCharacterFormatRange> formattingRanges,
            out string? warning) {
            warning = null;
            if (fib.CcpHdd == 0 || fib.LcbPlcfHdd == 0) {
                return Array.Empty<LegacyDocHeaderFooterStory>();
            }

            if (fib.FcPlcfHdd < 0 || fib.LcbPlcfHdd < 8 || fib.FcPlcfHdd + fib.LcbPlcfHdd > tableStream.Length) {
                warning = "The FIB points outside the selected table stream for the header/footer story PLC.";
                return Array.Empty<LegacyDocHeaderFooterStory>();
            }

            int cpCount = fib.LcbPlcfHdd / 4;
            if (fib.LcbPlcfHdd % 4 != 0 || cpCount < SeparatorStoryCount + 2) {
                warning = "The header/footer story PLC has an invalid length.";
                return Array.Empty<LegacyDocHeaderFooterStory>();
            }

            int[] storyPositions = new int[cpCount];
            for (int index = 0; index < cpCount; index++) {
                storyPositions[index] = LegacyDocFib.ReadInt32(tableStream, fib.FcPlcfHdd + (index * 4));
            }

            var stories = new List<LegacyDocHeaderFooterStory>();
            int storyCount = cpCount - 2;
            int headerBaseCharacterPosition = fib.CcpText + fib.CcpFtn;
            for (int storyIndex = SeparatorStoryCount; storyIndex < storyCount; storyIndex++) {
                int startCharacter = storyPositions[storyIndex];
                int endCharacter = storyPositions[storyIndex + 1];
                if (endCharacter < startCharacter) {
                    warning = "The header/footer story PLC contains a non-monotonic character range.";
                    return Array.Empty<LegacyDocHeaderFooterStory>();
                }

                if (startCharacter < 0 || endCharacter > fib.CcpHdd) {
                    warning = "The header/footer story PLC contains a character range outside the header/footer story.";
                    return Array.Empty<LegacyDocHeaderFooterStory>();
                }

                int relativeStoryIndex = storyIndex - SeparatorStoryCount;
                int sectionIndex = relativeStoryIndex / StoriesPerSection;
                int sectionStorySlot = relativeStoryIndex % StoriesPerSection;
                if (!TryMapStorySlot(sectionStorySlot, out bool isHeader, out HeaderFooterValues type)) {
                    continue;
                }

                IReadOnlyList<LegacyDocHeaderFooterParagraph> paragraphs = BuildStoryParagraphs(
                    textContent.AllCharacters,
                    headerBaseCharacterPosition + startCharacter,
                    headerBaseCharacterPosition + endCharacter,
                    formattingRanges);
                if (paragraphs.Count == 0) {
                    continue;
                }

                stories.Add(new LegacyDocHeaderFooterStory(sectionIndex, isHeader, type, paragraphs));
            }

            return stories;
        }

        private static bool TryMapStorySlot(int storySlot, out bool isHeader, out HeaderFooterValues type) {
            isHeader = false;
            type = HeaderFooterValues.Default;
            switch (storySlot) {
                case 0:
                    isHeader = true;
                    type = HeaderFooterValues.Even;
                    return true;
                case 1:
                    isHeader = true;
                    type = HeaderFooterValues.Default;
                    return true;
                case 2:
                    type = HeaderFooterValues.Even;
                    return true;
                case 3:
                    type = HeaderFooterValues.Default;
                    return true;
                case 4:
                    isHeader = true;
                    type = HeaderFooterValues.First;
                    return true;
                case 5:
                    type = HeaderFooterValues.First;
                    return true;
                default:
                    return false;
            }
        }

        private static IReadOnlyList<LegacyDocHeaderFooterParagraph> BuildStoryParagraphs(
            IReadOnlyList<LegacyDocTextCharacter> characters,
            int startCharacter,
            int endCharacter,
            IReadOnlyList<LegacyDocCharacterFormatRange> formattingRanges) {
            if (endCharacter <= startCharacter) {
                return Array.Empty<LegacyDocHeaderFooterParagraph>();
            }

            var paragraphs = new List<LegacyDocHeaderFooterParagraph>();
            var currentRuns = new List<LegacyDocTextRun>();
            var runText = new System.Text.StringBuilder(endCharacter - startCharacter);
            var runCharacterPositions = new List<int>();
            LegacyDocCharacterFormat currentFormat = LegacyDocCharacterFormat.Default;
            bool hasCurrentRun = false;

            foreach (LegacyDocTextCharacter character in characters) {
                if (character.CharacterPosition < startCharacter) {
                    continue;
                }

                if (character.CharacterPosition >= endCharacter) {
                    break;
                }

                char normalized = character.Character == '\a' ? '\r' : character.Character;
                if (normalized == '\r') {
                    AddCurrentParagraph();
                    continue;
                }

                if (char.IsControl(normalized) && normalized != '\t') {
                    continue;
                }

                AppendRunCharacter(
                    normalized,
                    GetFormatForFileOffset(formattingRanges, character.FileOffset),
                    character.CharacterPosition);
            }

            AddCurrentParagraph();
            return paragraphs;

            void AppendRunCharacter(char value, LegacyDocCharacterFormat format, int characterPosition) {
                if (!hasCurrentRun || !format.Equals(currentFormat)) {
                    FlushRun();
                    currentFormat = format;
                    hasCurrentRun = true;
                }

                runText.Append(value);
                runCharacterPositions.Add(characterPosition);
            }

            void AddCurrentParagraph() {
                FlushRun();
                if (currentRuns.Count > 0) {
                    paragraphs.Add(new LegacyDocHeaderFooterParagraph(currentRuns.ToArray()));
                    currentRuns.Clear();
                }

                hasCurrentRun = false;
            }

            void FlushRun() {
                if (runText.Length == 0) {
                    return;
                }

                currentRuns.Add(new LegacyDocTextRun(
                    runText.ToString(),
                    currentFormat.Bold,
                    currentFormat.Italic,
                    currentFormat.Strike,
                    currentFormat.DoubleStrike,
                    currentFormat.Outline,
                    currentFormat.Shadow,
                    currentFormat.Emboss,
                    currentFormat.Imprint,
                    currentFormat.Hidden,
                    currentFormat.Caps,
                    currentFormat.VerticalPosition,
                    currentFormat.Underline,
                    currentFormat.Highlight,
                    currentFormat.FontSizeHalfPoints,
                    currentFormat.ColorHex,
                    currentFormat.FontFamily,
                    runCharacterPositions));
                runText.Clear();
                runCharacterPositions.Clear();
            }
        }

        private static LegacyDocCharacterFormat GetFormatForFileOffset(IReadOnlyList<LegacyDocCharacterFormatRange> ranges, int fileOffset) {
            for (int i = 0; i < ranges.Count; i++) {
                if (ranges[i].Contains(fileOffset)) {
                    return ranges[i].Format;
                }
            }

            return LegacyDocCharacterFormat.Default;
        }
    }
}
