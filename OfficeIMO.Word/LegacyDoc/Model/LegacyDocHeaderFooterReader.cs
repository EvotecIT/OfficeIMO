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
            var pendingBookmarks = new List<LegacyDocBookmark>();
            LegacyDocCharacterFormat currentFormat = LegacyDocCharacterFormat.Default;
            LegacyDocHyperlinkTarget currentHyperlinkTarget = default;
            bool hasCurrentRun = false;
            int currentParagraphStartCharacter = startCharacter;
            int pendingBookmarkStartCharacter = int.MaxValue;
            int pendingBookmarkEndCharacter = int.MinValue;

            LegacyDocTextCharacter[] storyCharacters = characters
                .Where(character => character.CharacterPosition >= startCharacter && character.CharacterPosition < endCharacter)
                .ToArray();
            bool preserveEmptyParagraph = false;

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

                if (LegacyDocField.TryReadPageNumber(
                    storyCharacters,
                    index,
                    out int pageNumberResultStartIndex,
                    out int pageNumberResultEndIndex,
                    out int pageNumberFieldEndIndex)) {
                    AppendPageNumberResult(pageNumberResultStartIndex, pageNumberResultEndIndex);
                    index = pageNumberFieldEndIndex;
                    continue;
                }

                if (LegacyDocField.TryReadNumberOfPages(
                    storyCharacters,
                    index,
                    out int numberOfPagesResultStartIndex,
                    out int numberOfPagesResultEndIndex,
                    out int numberOfPagesFieldEndIndex)) {
                    AppendFieldResult(LegacyDocFieldKind.NumPages, fieldInstruction: null, numberOfPagesResultStartIndex, numberOfPagesResultEndIndex);
                    index = numberOfPagesFieldEndIndex;
                    continue;
                }

                if (LegacyDocField.TryReadDateTimeField(
                    storyCharacters,
                    index,
                    out LegacyDocFieldKind dateTimeFieldKind,
                    out string dateInstruction,
                    out int dateResultStartIndex,
                    out int dateResultEndIndex,
                    out int dateFieldEndIndex)) {
                    AppendFieldResult(dateTimeFieldKind, dateInstruction, dateResultStartIndex, dateResultEndIndex);
                    index = dateFieldEndIndex;
                    continue;
                }

                char normalized = character.Character == '\a' ? '\r' : character.Character;
                if (normalized == '\r') {
                    preserveEmptyParagraph = HasLaterHeaderFooterParagraphContent(storyCharacters, index + 1);
                    AddCurrentParagraph(GetParagraphFormatForFileOffset(paragraphFormattingRanges, character.FileOffset), character.CharacterPosition, isFinalParagraph: false);
                    currentParagraphStartCharacter = character.CharacterPosition + 1;
                    preserveEmptyParagraph = false;
                    continue;
                }

                if (char.IsControl(normalized)
                    && !LegacyDocSpecialCharacters.IsSupportedInlineControl(normalized)) {
                    continue;
                }

                AppendRunCharacter(
                    normalized,
                    GetFormatForFileOffset(formattingRanges, character.FileOffset),
                    character.CharacterPosition,
                    default);
            }

            AddCurrentParagraph(LegacyDocParagraphFormat.Default, endCharacter, isFinalParagraph: true);
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

            void AppendPageNumberResult(int resultStartIndex, int resultEndIndex) {
                AppendFieldResult(LegacyDocFieldKind.Page, fieldInstruction: null, resultStartIndex, resultEndIndex);
            }

            void AppendFieldResult(LegacyDocFieldKind fieldKind, string? fieldInstruction, int resultStartIndex, int resultEndIndex) {
                FlushRun();
                LegacyDocCharacterFormat format = LegacyDocCharacterFormat.Default;
                var positions = new List<int>();
                var resultText = new System.Text.StringBuilder();
                for (int resultIndex = resultStartIndex; resultIndex < resultEndIndex; resultIndex++) {
                    LegacyDocTextCharacter resultCharacter = storyCharacters[resultIndex];
                    if (char.IsControl(resultCharacter.Character)
                        && !LegacyDocSpecialCharacters.IsSupportedInlineControl(resultCharacter.Character)) {
                        continue;
                    }

                    if (positions.Count == 0) {
                        format = GetFormatForFileOffset(formattingRanges, resultCharacter.FileOffset);
                    }

                    resultText.Append(resultCharacter.Character);
                    positions.Add(resultCharacter.CharacterPosition);
                }

                currentRuns.Add(LegacyDocTextRunFactory.CreateFieldRun(
                    fieldKind == LegacyDocFieldKind.Page ? string.Empty : resultText.ToString(),
                    fieldKind,
                    fieldInstruction,
                    format,
                    positions));
            }

            void AddCurrentParagraph(LegacyDocParagraphFormat format, int paragraphEndCharacter, bool isFinalParagraph) {
                FlushRun();
                if (isFinalParagraph
                    && currentRuns.Count == 0
                    && pendingBookmarks.Count == 0
                    && currentParagraphStartCharacter == paragraphEndCharacter) {
                    hasCurrentRun = false;
                    currentHyperlinkTarget = default;
                    return;
                }

                IReadOnlyList<LegacyDocBookmark> paragraphBookmarks = bookmarkProjection.ExtractProjectedParagraphBookmarks(currentParagraphStartCharacter, paragraphEndCharacter);
                if (currentRuns.Count > 0) {
                    bool hasPendingBookmarks = pendingBookmarks.Count > 0;
                    paragraphs.Add(new LegacyDocHeaderFooterParagraph(
                        currentRuns.ToArray(),
                        format,
                        hasPendingBookmarks ? Math.Min(currentParagraphStartCharacter, pendingBookmarkStartCharacter) : currentParagraphStartCharacter,
                        hasPendingBookmarks ? Math.Max(paragraphEndCharacter, pendingBookmarkEndCharacter) : paragraphEndCharacter,
                        MergePendingBookmarks(paragraphBookmarks)));
                    currentRuns.Clear();
                    ClearPendingBookmarks();
                } else {
                    AddPendingBookmarks(paragraphBookmarks);
                    if (preserveEmptyParagraph || paragraphBookmarks.Count > 0) {
                        paragraphs.Add(new LegacyDocHeaderFooterParagraph(
                            Array.Empty<LegacyDocTextRun>(),
                            format,
                            currentParagraphStartCharacter,
                            paragraphEndCharacter,
                            MergePendingBookmarks(paragraphBookmarks)));
                        ClearPendingBookmarks();
                    } else if (isFinalParagraph && pendingBookmarks.Count > 0) {
                        paragraphs.Add(new LegacyDocHeaderFooterParagraph(
                            Array.Empty<LegacyDocTextRun>(),
                            format,
                            pendingBookmarkStartCharacter,
                            pendingBookmarkEndCharacter,
                            pendingBookmarks.ToArray()));
                        ClearPendingBookmarks();
                    }
                }

                hasCurrentRun = false;
                currentHyperlinkTarget = default;
            }

            void AddPendingBookmarks(IReadOnlyList<LegacyDocBookmark> bookmarks) {
                foreach (LegacyDocBookmark bookmark in bookmarks) {
                    if (!pendingBookmarks.Contains(bookmark)) {
                        pendingBookmarks.Add(bookmark);
                    }

                    pendingBookmarkStartCharacter = Math.Min(pendingBookmarkStartCharacter, bookmark.StartCharacter);
                    pendingBookmarkEndCharacter = Math.Max(pendingBookmarkEndCharacter, bookmark.EndCharacter);
                }
            }

            IReadOnlyList<LegacyDocBookmark> MergePendingBookmarks(IReadOnlyList<LegacyDocBookmark> paragraphBookmarks) {
                if (pendingBookmarks.Count == 0) {
                    return paragraphBookmarks;
                }

                if (paragraphBookmarks.Count == 0) {
                    return pendingBookmarks.ToArray();
                }

                return pendingBookmarks
                    .Concat(paragraphBookmarks)
                    .Distinct()
                    .ToArray();
            }

            void ClearPendingBookmarks() {
                pendingBookmarks.Clear();
                pendingBookmarkStartCharacter = int.MaxValue;
                pendingBookmarkEndCharacter = int.MinValue;
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
                    currentHyperlinkTarget.Anchor,
                    specified: currentFormat.Specified,
                    characterSpacingTwips: currentFormat.CharacterSpacingTwips));
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

        private static bool HasLaterHeaderFooterParagraphContent(IReadOnlyList<LegacyDocTextCharacter> characters, int startIndex) {
            for (int index = startIndex; index < characters.Count; index++) {
                char normalized = characters[index].Character == '\a' ? '\r' : characters[index].Character;
                if (normalized == '\r') {
                    continue;
                }

                if (char.IsControl(normalized)
                    && !LegacyDocSpecialCharacters.IsSupportedInlineControl(normalized)) {
                    continue;
                }

                return true;
            }

            return false;
        }
    }
}
