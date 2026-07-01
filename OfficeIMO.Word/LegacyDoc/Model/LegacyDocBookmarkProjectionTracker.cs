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
                bool containsStart = IsBookmarkStartAtBlockBoundary(bookmark, blockStartCharacter, blockEndCharacter)
                    && !_projectedStarts.Contains(bookmark);
                bool containsEnd = IsBookmarkEndAtBlockBoundary(bookmark, blockStartCharacter, blockEndCharacter)
                    && !_projectedEnds.Contains(bookmark);
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

        internal IReadOnlyList<LegacyDocBookmark> ExtractZeroLengthBoundaryBookmarks(int boundaryCharacter) {
            if (_bookmarks.Count == 0) {
                return Array.Empty<LegacyDocBookmark>();
            }

            var result = new List<LegacyDocBookmark>();
            foreach (LegacyDocBookmark bookmark in _bookmarks) {
                if (!bookmark.IsZeroLength
                    || bookmark.StartCharacter != boundaryCharacter
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

        private static bool IsBookmarkStartAtBlockBoundary(LegacyDocBookmark bookmark, int blockStartCharacter, int blockEndCharacter) =>
            bookmark.StartCharacter == blockStartCharacter
            || (bookmark.IsZeroLength && bookmark.StartCharacter == blockEndCharacter);

        private static bool IsBookmarkEndAtBlockBoundary(LegacyDocBookmark bookmark, int blockStartCharacter, int blockEndCharacter) =>
            bookmark.EndCharacter == blockEndCharacter
            || (bookmark.IsZeroLength && bookmark.EndCharacter == blockStartCharacter);

        internal IEnumerable<LegacyDocBookmark> GetUnprojectedBookmarks() {
            foreach (LegacyDocBookmark bookmark in _bookmarks) {
                if (!_projectedStarts.Contains(bookmark) || !_projectedEnds.Contains(bookmark)) {
                    yield return bookmark;
                }
            }
        }
    }
}
