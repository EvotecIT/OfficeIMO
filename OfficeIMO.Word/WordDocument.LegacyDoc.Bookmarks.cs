using OfficeIMO.Word.LegacyDoc.Model;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private sealed class LegacyDocBookmarkProjection {
            internal static LegacyDocBookmarkProjection Empty { get; } = new LegacyDocBookmarkProjection(Array.Empty<LegacyDocProjectedBookmark>());

            private readonly IReadOnlyList<LegacyDocProjectedBookmark> _bookmarks;
            private readonly Dictionary<int, List<LegacyDocProjectedBookmark>> _starts;
            private readonly Dictionary<int, List<LegacyDocProjectedBookmark>> _ends;
            private readonly HashSet<int> _emittedPositions = new();

            private LegacyDocBookmarkProjection(IReadOnlyList<LegacyDocProjectedBookmark> bookmarks) {
                _bookmarks = bookmarks;
                _starts = bookmarks
                    .GroupBy(bookmark => bookmark.StartCharacter)
                    .ToDictionary(group => group.Key, group => group.OrderByDescending(bookmark => bookmark.EndCharacter).ThenBy(bookmark => bookmark.Name, StringComparer.Ordinal).ToList());
                _ends = bookmarks
                    .GroupBy(bookmark => bookmark.EndCharacter)
                    .ToDictionary(group => group.Key, group => group.OrderByDescending(bookmark => bookmark.StartCharacter).ThenBy(bookmark => bookmark.Name, StringComparer.Ordinal).ToList());
            }

            internal static LegacyDocBookmarkProjection Create(WordParagraph paragraph, IReadOnlyList<LegacyDocBookmark> bookmarks) {
                if (bookmarks.Count == 0) {
                    return Empty;
                }

                int nextBookmarkId = paragraph._document.BookmarkId;
                LegacyDocProjectedBookmark[] projected = bookmarks
                    .OrderBy(bookmark => bookmark.StartCharacter)
                    .ThenByDescending(bookmark => bookmark.EndCharacter)
                    .ThenBy(bookmark => bookmark.Name, StringComparer.Ordinal)
                    .Select(bookmark => new LegacyDocProjectedBookmark(bookmark, (nextBookmarkId++).ToString()))
                    .ToArray();

                return new LegacyDocBookmarkProjection(projected);
            }

            internal bool HasMarkers(int? characterPosition) {
                if (characterPosition == null || _emittedPositions.Contains(characterPosition.Value)) {
                    return false;
                }

                return _starts.ContainsKey(characterPosition.Value) || _ends.ContainsKey(characterPosition.Value);
            }

            internal void EmitAt(OpenXmlCompositeElement target, int? characterPosition) {
                if (characterPosition == null || !HasMarkers(characterPosition)) {
                    return;
                }

                int position = characterPosition.Value;
                if (_ends.TryGetValue(position, out List<LegacyDocProjectedBookmark>? ends)) {
                    foreach (LegacyDocProjectedBookmark bookmark in ends.Where(bookmark => !bookmark.IsZeroLength)) {
                        target.Append(new BookmarkEnd { Id = bookmark.Id });
                    }
                }

                if (_starts.TryGetValue(position, out List<LegacyDocProjectedBookmark>? starts)) {
                    foreach (LegacyDocProjectedBookmark bookmark in starts) {
                        target.Append(new BookmarkStart { Id = bookmark.Id, Name = bookmark.Name });
                    }
                }

                if (ends != null) {
                    foreach (LegacyDocProjectedBookmark bookmark in ends.Where(bookmark => bookmark.IsZeroLength)) {
                        target.Append(new BookmarkEnd { Id = bookmark.Id });
                    }
                }

                _emittedPositions.Add(position);
            }

            internal void EmitRemaining(OpenXmlCompositeElement target) {
                foreach (int position in _bookmarks
                    .SelectMany(bookmark => new[] { bookmark.StartCharacter, bookmark.EndCharacter })
                    .Distinct()
                    .OrderBy(position => position)) {
                    EmitAt(target, position);
                }
            }
        }

        private sealed class LegacyDocProjectedBookmark {
            internal LegacyDocProjectedBookmark(LegacyDocBookmark bookmark, string id) {
                Name = bookmark.Name;
                StartCharacter = bookmark.StartCharacter;
                EndCharacter = bookmark.EndCharacter;
                Id = id;
            }

            internal string Name { get; }

            internal int StartCharacter { get; }

            internal int EndCharacter { get; }

            internal string Id { get; }

            internal bool IsZeroLength => StartCharacter == EndCharacter;
        }
    }
}
