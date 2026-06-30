using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static class LegacyDocHeaderFooterReader {
        private const int SeparatorStoryCount = 6;
        private const int StoriesPerSection = 6;

        internal static IReadOnlyList<LegacyDocHeaderFooterStory> Read(byte[] tableStream, LegacyDocTextContent textContent, LegacyDocFib fib, out string? warning) {
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

                string rawText = ExtractStoryText(
                    textContent.AllCharacters,
                    headerBaseCharacterPosition + startCharacter,
                    headerBaseCharacterPosition + endCharacter);
                IReadOnlyList<string> paragraphs = SplitStoryParagraphs(rawText);
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
                .Where(paragraph => paragraph.Length > 0)
                .ToArray();
        }
    }
}
