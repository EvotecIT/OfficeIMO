namespace OfficeIMO.Word {
    public partial class WordDocument {
        // Identifier counters are seeded by one scan and advanced on each allocation, so adding items one at a time
        // stays linear. Paths that can bring in identifiers from outside (merging, templates, legacy import, field
        // and table-of-contents updates, saving) drop the counters, and the next allocation scans again.
        private int? _lastBookmarkId;
        internal WordCommentIdentityState? CommentIdentities { get; set; }

        /// <summary>Records a bookmark identifier written without <see cref="GetNextBookmarkId"/>.</summary>
        internal void ObserveBookmarkId(int id) {
            if (_lastBookmarkId is int last && id > last) _lastBookmarkId = id;
        }

        /// <summary>Forgets cached bookmark, comment and drawing identifiers; the next allocation rescans the package.</summary>
        internal void InvalidateIdentifierState() {
            _lastBookmarkId = null;
            CommentIdentities = null;
            if (_wordprocessingDocument != null) WordDrawingIdAllocator.Invalidate(_wordprocessingDocument);
        }
    }
}
