namespace OfficeIMO.Word.LegacyDoc.Model {
    internal sealed class LegacyDocTextBoxStory {
        internal LegacyDocTextBoxStory(
            bool isHeaderFooterTextBox,
            string text,
            int startCharacter,
            int endCharacter,
            IReadOnlyList<LegacyDocTextRun>? runs = null,
            IReadOnlyList<LegacyDocBookmark>? bookmarks = null) {
            IsHeaderFooterTextBox = isHeaderFooterTextBox;
            Text = text ?? throw new ArgumentNullException(nameof(text));
            StartCharacter = startCharacter;
            EndCharacter = endCharacter;
            Runs = runs ?? Array.Empty<LegacyDocTextRun>();
            Bookmarks = bookmarks ?? Array.Empty<LegacyDocBookmark>();
        }

        internal bool IsHeaderFooterTextBox { get; }

        internal string Text { get; }

        internal int StartCharacter { get; }

        internal int EndCharacter { get; }

        internal IReadOnlyList<LegacyDocTextRun> Runs { get; }

        internal IReadOnlyList<LegacyDocBookmark> Bookmarks { get; }
    }
}
