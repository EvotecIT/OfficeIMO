using System.Collections.Generic;

namespace OfficeIMO.Word.LegacyDoc.Model {
    internal sealed class LegacyDocBookmarkProjectionTracker {
        private readonly IReadOnlyList<LegacyDocBookmark> _bookmarks;
        private readonly HashSet<LegacyDocBookmark> _projectedStarts = new();
        private readonly HashSet<LegacyDocBookmark> _projectedEnds = new();

        internal LegacyDocBookmarkProjectionTracker(IReadOnlyList<LegacyDocBookmark> bookmarks) {
            _bookmarks = bookmarks;
        }

        internal IReadOnlyList<LegacyDocBookmark> ExtractProjectedParagraphBookmarks(int paragraphStartCharacter, int paragraphEndCharacter) {
            if (_bookmarks.Count == 0) {
                return Array.Empty<LegacyDocBookmark>();
            }

            var result = new List<LegacyDocBookmark>();
            foreach (LegacyDocBookmark bookmark in _bookmarks) {
                bool containsStart = bookmark.StartCharacter >= paragraphStartCharacter
                    && bookmark.StartCharacter <= paragraphEndCharacter;
                bool containsEnd = bookmark.EndCharacter >= paragraphStartCharacter
                    && bookmark.EndCharacter <= paragraphEndCharacter;
                if (!containsStart && !containsEnd) {
                    continue;
                }

                result.Add(bookmark);
                if (containsStart) {
                    _projectedStarts.Add(bookmark);
                }

                if (containsEnd) {
                    _projectedEnds.Add(bookmark);
                }
            }

            return result;
        }

        internal IReadOnlyList<LegacyDocBookmark> ExtractUnprojectedBlockBookmarks(int blockStartCharacter, int blockEndCharacter) {
            if (_bookmarks.Count == 0) {
                return Array.Empty<LegacyDocBookmark>();
            }

            var result = new List<LegacyDocBookmark>();
            foreach (LegacyDocBookmark bookmark in _bookmarks) {
                if (!IsBlockBoundaryBookmark(bookmark, blockStartCharacter, blockEndCharacter)
                    || _projectedStarts.Contains(bookmark)
                    || _projectedEnds.Contains(bookmark)) {
                    continue;
                }

                result.Add(bookmark);
                _projectedStarts.Add(bookmark);
                _projectedEnds.Add(bookmark);
            }

            return result;
        }

        private static bool IsBlockBoundaryBookmark(LegacyDocBookmark bookmark, int blockStartCharacter, int blockEndCharacter) {
            if (bookmark.StartCharacter == blockStartCharacter && bookmark.EndCharacter == blockEndCharacter) {
                return true;
            }

            return bookmark.IsZeroLength
                && (bookmark.StartCharacter == blockStartCharacter
                    || bookmark.StartCharacter == blockEndCharacter);
        }

        internal IEnumerable<LegacyDocBookmark> GetUnprojectedBookmarks() {
            foreach (LegacyDocBookmark bookmark in _bookmarks) {
                if (!_projectedStarts.Contains(bookmark) || !_projectedEnds.Contains(bookmark)) {
                    yield return bookmark;
                }
            }
        }
    }
}
