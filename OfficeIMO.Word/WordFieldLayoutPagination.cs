namespace OfficeIMO.Word {
    /// <summary>Layout results consumed by Word fields without depending on a rendering package.</summary>
    internal sealed class WordLayoutPagination {
        private readonly Dictionary<string, (int Number, string? Label)> _bookmarkPages;

        internal WordLayoutPagination(Dictionary<string, (int Number, string? Label)> bookmarkPages, int pageCount) {
            _bookmarkPages = bookmarkPages;
            PageCount = pageCount;
        }

        internal int PageCount { get; }

        internal bool TryGetBookmarkPage(string bookmarkName, out int pageNumber, out string? label) {
            if (_bookmarkPages.TryGetValue(bookmarkName, out var page)) {
                pageNumber = page.Number;
                label = page.Label;
                return true;
            }
            pageNumber = 0;
            label = null;
            return false;
        }
    }
}
