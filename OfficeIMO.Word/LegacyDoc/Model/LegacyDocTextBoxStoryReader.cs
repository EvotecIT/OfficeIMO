namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static class LegacyDocTextBoxStoryReader {
        internal static IReadOnlyList<LegacyDocTextBoxStory> Read(
            LegacyDocTextContent textContent,
            LegacyDocFib fib,
            IReadOnlyList<LegacyDocCharacterFormatRange> formattingRanges,
            IReadOnlyList<LegacyDocParagraphFormatRange> paragraphFormattingRanges,
            LegacyDocBookmarkProjectionTracker bookmarkProjection,
            IReadOnlyDictionary<int, LegacyDocPicture> picturesByCharacterPosition) {
            var stories = new List<LegacyDocTextBoxStory>(2);
            int bodyBase = fib.CcpText + fib.CcpFtn + fib.CcpHdd + fib.CcpAtn + fib.CcpEdn;
            AddStory(bodyBase, fib.CcpTxbx, isHeaderFooter: false);
            AddStory(bodyBase + fib.CcpTxbx, fib.CcpHdrTxbx, isHeaderFooter: true);
            return stories;

            void AddStory(int start, int count, bool isHeaderFooter) {
                if (count <= 0) return;
                int end = start + count;
                // The shared story reader retains PAPX styles, paragraph boundaries, fields and run operands.
                IReadOnlyList<LegacyDocNoteParagraph> paragraphs = LegacyDocFootnoteReader.BuildStoryParagraphs(
                    textContent.AllCharacters, start, end, formattingRanges, paragraphFormattingRanges,
                    bookmarkProjection, picturesByCharacterPosition);
                string text = string.Join(Environment.NewLine, paragraphs.Select(paragraph => paragraph.Text)).Trim();
                if (string.IsNullOrWhiteSpace(text)) return;
                stories.Add(new LegacyDocTextBoxStory(isHeaderFooter, text, start, end,
                    paragraphs.SelectMany(paragraph => paragraph.Runs).ToArray(),
                    paragraphs.SelectMany(paragraph => paragraph.Bookmarks).Distinct().ToArray(), paragraphs));
            }
        }
    }
}
