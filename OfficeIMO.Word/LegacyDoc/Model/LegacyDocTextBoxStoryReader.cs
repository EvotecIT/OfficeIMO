namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static class LegacyDocTextBoxStoryReader {
        internal static IReadOnlyList<LegacyDocTextBoxStory> Read(
            LegacyDocTextContent textContent,
            LegacyDocFib fib,
            IReadOnlyList<LegacyDocCharacterFormatRange> formattingRanges,
            LegacyDocBookmarkProjectionTracker bookmarkProjection) {
            var stories = new List<LegacyDocTextBoxStory>(2);
            int bodyTextBoxBase = fib.CcpText + fib.CcpFtn + fib.CcpHdd + fib.CcpAtn + fib.CcpEdn;
            AddStoryIfPresent(stories, textContent.AllCharacters, bodyTextBoxBase, fib.CcpTxbx, isHeaderFooterTextBox: false, formattingRanges, bookmarkProjection);
            AddStoryIfPresent(stories, textContent.AllCharacters, bodyTextBoxBase + fib.CcpTxbx, fib.CcpHdrTxbx, isHeaderFooterTextBox: true, formattingRanges, bookmarkProjection);
            return stories;
        }

        private static void AddStoryIfPresent(
            List<LegacyDocTextBoxStory> stories,
            IReadOnlyList<LegacyDocTextCharacter> characters,
            int startCharacter,
            int characterCount,
            bool isHeaderFooterTextBox,
            IReadOnlyList<LegacyDocCharacterFormatRange> formattingRanges,
            LegacyDocBookmarkProjectionTracker bookmarkProjection) {
            if (characterCount <= 0) {
                return;
            }

            int endCharacter = startCharacter + characterCount;
            string text = ReadStoryText(characters, startCharacter, endCharacter);
            if (string.IsNullOrWhiteSpace(text)) {
                return;
            }

            IReadOnlyList<LegacyDocTextRun> runs = CreateStoryRuns(characters, startCharacter, endCharacter, formattingRanges);
            IReadOnlyList<LegacyDocBookmark> bookmarks = bookmarkProjection.ExtractProjectedParagraphBookmarks(startCharacter, endCharacter);
            stories.Add(new LegacyDocTextBoxStory(isHeaderFooterTextBox, text, startCharacter, endCharacter, runs, bookmarks));
        }

        private static string ReadStoryText(IReadOnlyList<LegacyDocTextCharacter> characters, int startCharacter, int endCharacter) {
            var builder = new System.Text.StringBuilder(Math.Max(0, endCharacter - startCharacter));
            foreach (LegacyDocTextCharacter character in characters) {
                if (character.CharacterPosition < startCharacter || character.CharacterPosition >= endCharacter) {
                    continue;
                }

                switch (character.Character) {
                    case '\r':
                    case '\v':
                    case '\f':
                        builder.Append(Environment.NewLine);
                        break;
                    case '\0':
                    case '\u0007':
                        break;
                    default:
                        builder.Append(character.Character);
                        break;
                }
            }

            return builder.ToString().Trim();
        }

        private static IReadOnlyList<LegacyDocTextRun> CreateStoryRuns(
            IReadOnlyList<LegacyDocTextCharacter> characters,
            int startCharacter,
            int endCharacter,
            IReadOnlyList<LegacyDocCharacterFormatRange> formattingRanges) {
            var runs = new List<LegacyDocTextRun>();
            var text = new System.Text.StringBuilder(Math.Max(0, endCharacter - startCharacter));
            var positions = new List<int>();
            LegacyDocCharacterFormat currentFormat = LegacyDocCharacterFormat.Default;
            bool hasCurrentFormat = false;
            foreach (LegacyDocTextCharacter character in characters) {
                if (character.CharacterPosition < startCharacter || character.CharacterPosition >= endCharacter) {
                    continue;
                }

                char normalized;
                switch (character.Character) {
                    case '\0':
                    case '\u0007':
                        continue;
                    case '\r':
                    case '\n':
                    case '\v':
                    case '\f':
                        normalized = LegacyDocSpecialCharacters.TextWrappingBreak;
                        break;
                    default:
                        if (char.IsControl(character.Character) && character.Character != '\t') {
                            continue;
                        }

                        normalized = character.Character;
                        break;
                }

                LegacyDocCharacterFormat format = GetFormatForFileOffset(formattingRanges, character.FileOffset);
                if (hasCurrentFormat && !currentFormat.Equals(format)) {
                    AddRun(runs, text, positions, currentFormat);
                }

                currentFormat = format;
                hasCurrentFormat = true;
                text.Append(normalized);
                positions.Add(character.CharacterPosition);
            }

            if (!hasCurrentFormat) {
                return Array.Empty<LegacyDocTextRun>();
            }

            AddRun(runs, text, positions, currentFormat);
            return runs.Count == 0
                ? Array.Empty<LegacyDocTextRun>()
                : runs.ToArray();
        }

        private static void AddRun(
            List<LegacyDocTextRun> runs,
            System.Text.StringBuilder text,
            List<int> positions,
            LegacyDocCharacterFormat format) {
            if (text.Length == 0) {
                return;
            }

            runs.Add(new LegacyDocTextRun(
                text.ToString(),
                format.Bold,
                format.Italic,
                format.Strike,
                format.DoubleStrike,
                format.Outline,
                format.Shadow,
                format.Emboss,
                format.Imprint,
                format.Hidden,
                format.NoProof,
                format.Caps,
                format.VerticalPosition,
                format.Underline,
                format.Highlight,
                format.FontSizeHalfPoints,
                format.ColorHex,
                format.FontFamily,
                positions,
                specified: format.Specified,
                characterSpacingTwips: format.CharacterSpacingTwips,
                language: format.Language,
                eastAsiaLanguage: format.EastAsiaLanguage,
                revision: format.Revision));
            text.Clear();
            positions.Clear();
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
