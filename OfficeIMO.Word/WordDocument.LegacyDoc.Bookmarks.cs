using OfficeIMO.Word.LegacyDoc.Model;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private sealed class LegacyDocBookmarkProjection {
            internal static LegacyDocBookmarkProjection Empty { get; } = new LegacyDocBookmarkProjection(Array.Empty<LegacyDocProjectedBookmark>());

            private readonly Dictionary<int, List<LegacyDocProjectedBookmark>> _starts;
            private readonly Dictionary<int, List<LegacyDocProjectedBookmark>> _ends;
            private readonly HashSet<int> _emittedPositions = new();

            private LegacyDocBookmarkProjection(IReadOnlyList<LegacyDocProjectedBookmark> bookmarks) {
                _starts = bookmarks
                    .Where(bookmark => bookmark.ProjectStart)
                    .GroupBy(bookmark => bookmark.StartCharacter)
                    .ToDictionary(group => group.Key, group => group.OrderByDescending(bookmark => bookmark.EndCharacter).ThenBy(bookmark => bookmark.Name, StringComparer.Ordinal).ToList());
                _ends = bookmarks
                    .Where(bookmark => bookmark.ProjectEnd)
                    .GroupBy(bookmark => bookmark.EndCharacter)
                    .ToDictionary(group => group.Key, group => group.OrderByDescending(bookmark => bookmark.StartCharacter).ThenBy(bookmark => bookmark.Name, StringComparer.Ordinal).ToList());
            }

            internal static LegacyDocBookmarkProjection Create(IReadOnlyList<LegacyDocBookmark> bookmarks, int paragraphStartCharacter, int paragraphEndCharacter) {
                if (bookmarks.Count == 0) {
                    return Empty;
                }

                LegacyDocProjectedBookmark[] projected = bookmarks
                    .OrderBy(bookmark => bookmark.StartCharacter)
                    .ThenByDescending(bookmark => bookmark.EndCharacter)
                    .ThenBy(bookmark => bookmark.Name, StringComparer.Ordinal)
                    .Select(bookmark => new LegacyDocProjectedBookmark(
                        bookmark,
                        IsWithinParagraph(bookmark.StartCharacter, paragraphStartCharacter, paragraphEndCharacter),
                        IsWithinParagraph(bookmark.EndCharacter, paragraphStartCharacter, paragraphEndCharacter)))
                    .Where(bookmark => bookmark.ProjectStart || bookmark.ProjectEnd)
                    .ToArray();

                return new LegacyDocBookmarkProjection(projected);
            }

            private static bool IsWithinParagraph(int characterPosition, int paragraphStartCharacter, int paragraphEndCharacter) {
                return characterPosition >= paragraphStartCharacter && characterPosition <= paragraphEndCharacter;
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
                foreach (int position in _starts.Keys
                    .Concat(_ends.Keys)
                    .Distinct()
                    .OrderBy(position => position)) {
                    EmitAt(target, position);
                }
            }
        }

        private sealed class LegacyDocProjectedBookmark {
            internal LegacyDocProjectedBookmark(LegacyDocBookmark bookmark, bool projectStart, bool projectEnd) {
                Name = bookmark.Name;
                StartCharacter = bookmark.StartCharacter;
                EndCharacter = bookmark.EndCharacter;
                Id = bookmark.ProjectionId;
                ProjectStart = projectStart;
                ProjectEnd = projectEnd;
            }

            internal string Name { get; }

            internal int StartCharacter { get; }

            internal int EndCharacter { get; }

            internal string Id { get; }

            internal bool ProjectStart { get; }

            internal bool ProjectEnd { get; }

            internal bool IsZeroLength => StartCharacter == EndCharacter;
        }
    }
}
