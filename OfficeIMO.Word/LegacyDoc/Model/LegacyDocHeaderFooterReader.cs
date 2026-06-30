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
            IReadOnlyList<LegacyDocParagraphFormatRange> paragraphFormattingRanges,
            LegacyDocBookmarkProjectionTracker bookmarkProjection,
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
                    formattingRanges,
                    paragraphFormattingRanges,
                    bookmarkProjection);
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
            IReadOnlyList<LegacyDocCharacterFormatRange> formattingRanges,
            IReadOnlyList<LegacyDocParagraphFormatRange> paragraphFormattingRanges,
            LegacyDocBookmarkProjectionTracker bookmarkProjection) {
            if (endCharacter <= startCharacter) {
                return Array.Empty<LegacyDocHeaderFooterParagraph>();
            }

            var paragraphs = new List<LegacyDocHeaderFooterParagraph>();
            var currentRuns = new List<LegacyDocTextRun>();
            var runText = new System.Text.StringBuilder(endCharacter - startCharacter);
            var runCharacterPositions = new List<int>();
            LegacyDocCharacterFormat currentFormat = LegacyDocCharacterFormat.Default;
            LegacyDocHyperlinkTarget currentHyperlinkTarget = default;
            bool hasCurrentRun = false;
            int currentParagraphStartCharacter = startCharacter;

            LegacyDocTextCharacter[] storyCharacters = characters
                .Where(character => character.CharacterPosition >= startCharacter && character.CharacterPosition < endCharacter)
                .ToArray();

            for (int index = 0; index < storyCharacters.Length; index++) {
                LegacyDocTextCharacter character = storyCharacters[index];

                if (LegacyDocField.TryReadHyperlink(
                    storyCharacters,
                    index,
                    out LegacyDocHyperlinkTarget hyperlinkTarget,
                    out int resultStartIndex,
                    out int resultEndIndex,
                    out int fieldEndIndex)) {
                    for (int resultIndex = resultStartIndex; resultIndex < resultEndIndex; resultIndex++) {
                        LegacyDocTextCharacter resultCharacter = storyCharacters[resultIndex];
                        AppendRunCharacter(
                            resultCharacter.Character,
                            GetFormatForFileOffset(formattingRanges, resultCharacter.FileOffset),
                            resultCharacter.CharacterPosition,
                            hyperlinkTarget);
                    }

                    index = fieldEndIndex;
                    continue;
                }

                char normalized = character.Character == '\a' ? '\r' : character.Character;
                if (normalized == '\r') {
                    AddCurrentParagraph(GetParagraphFormatForFileOffset(paragraphFormattingRanges, character.FileOffset), character.CharacterPosition);
                    currentParagraphStartCharacter = character.CharacterPosition + 1;
                    continue;
                }

                if (char.IsControl(normalized)
                    && normalized != '\t'
                    && normalized != '\v'
                    && normalized != '\f') {
                    continue;
                }

                AppendRunCharacter(
                    normalized,
                    GetFormatForFileOffset(formattingRanges, character.FileOffset),
                    character.CharacterPosition,
                    default);
            }

            AddCurrentParagraph(LegacyDocParagraphFormat.Default, endCharacter);
            return paragraphs;

            void AppendRunCharacter(char value, LegacyDocCharacterFormat format, int characterPosition, LegacyDocHyperlinkTarget hyperlinkTarget) {
                if (!hasCurrentRun
                    || !format.Equals(currentFormat)
                    || hyperlinkTarget != currentHyperlinkTarget) {
                    FlushRun();
                    currentFormat = format;
                    currentHyperlinkTarget = hyperlinkTarget;
                    hasCurrentRun = true;
                }

                runText.Append(value);
                runCharacterPositions.Add(characterPosition);
            }

            void AddCurrentParagraph(LegacyDocParagraphFormat format, int paragraphEndCharacter) {
                FlushRun();
                if (currentRuns.Count > 0) {
                    paragraphs.Add(new LegacyDocHeaderFooterParagraph(
                        currentRuns.ToArray(),
                        format,
                        currentParagraphStartCharacter,
                        paragraphEndCharacter,
                        bookmarkProjection.ExtractProjectedParagraphBookmarks(currentParagraphStartCharacter, paragraphEndCharacter)));
                    currentRuns.Clear();
                }

                hasCurrentRun = false;
                currentHyperlinkTarget = default;
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
                    currentFormat.NoProof,
                    currentFormat.Caps,
                    currentFormat.VerticalPosition,
                    currentFormat.Underline,
                    currentFormat.Highlight,
                    currentFormat.FontSizeHalfPoints,
                    currentFormat.ColorHex,
                    currentFormat.FontFamily,
                    runCharacterPositions,
                    currentHyperlinkTarget.Uri,
                    currentHyperlinkTarget.Anchor));
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

        private static LegacyDocParagraphFormat GetParagraphFormatForFileOffset(IReadOnlyList<LegacyDocParagraphFormatRange> ranges, int fileOffset) {
            for (int i = 0; i < ranges.Count; i++) {
                if (ranges[i].Contains(fileOffset)) {
                    return ranges[i].Format;
                }
            }

            return LegacyDocParagraphFormat.Default;
        }
    }
}
