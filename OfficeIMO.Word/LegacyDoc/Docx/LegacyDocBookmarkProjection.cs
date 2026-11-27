using System.Threading;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Word.LegacyDoc.Model;
using static OfficeIMO.Word.LegacyDoc.Docx.LegacyDocXml;

namespace OfficeIMO.Word.LegacyDoc.Docx {
    /// <summary>
    /// Places the bookmark starts and ends that fall inside one paragraph at their character positions, in the order
    /// Word expects when several meet at one position.
    /// </summary>
    internal sealed class LegacyDocBookmarkProjection {
        internal static LegacyDocBookmarkProjection Empty { get; } = new(Array.Empty<ProjectedBookmark>());

        private readonly Dictionary<int, List<ProjectedBookmark>> _starts;
        private readonly Dictionary<int, List<ProjectedBookmark>> _ends;
        private readonly HashSet<int> _emittedPositions = new();

        private LegacyDocBookmarkProjection(IReadOnlyList<ProjectedBookmark> bookmarks) {
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
            if (bookmarks.Count == 0) return Empty;
            ProjectedBookmark[] projected = bookmarks
                .OrderBy(bookmark => bookmark.StartCharacter)
                .ThenByDescending(bookmark => bookmark.EndCharacter)
                .ThenBy(bookmark => bookmark.Name, StringComparer.Ordinal)
                .Select(bookmark => new ProjectedBookmark(
                    bookmark,
                    IsWithinParagraph(bookmark.StartCharacter, paragraphStartCharacter, paragraphEndCharacter),
                    IsWithinParagraph(bookmark.EndCharacter, paragraphStartCharacter, paragraphEndCharacter)))
                .Where(bookmark => bookmark.ProjectStart || bookmark.ProjectEnd)
                .ToArray();
            return new LegacyDocBookmarkProjection(projected);
        }

        private static bool IsWithinParagraph(int characterPosition, int paragraphStartCharacter, int paragraphEndCharacter) =>
            characterPosition >= paragraphStartCharacter && characterPosition <= paragraphEndCharacter;

        internal bool HasMarkers(int? characterPosition) {
            if (characterPosition == null || _emittedPositions.Contains(characterPosition.Value)) return false;
            return _starts.ContainsKey(characterPosition.Value) || _ends.ContainsKey(characterPosition.Value);
        }

        internal void EmitAt(OfficeXmlTree tree, int target, int? characterPosition) {
            if (characterPosition == null || !HasMarkers(characterPosition)) return;
            int position = characterPosition.Value;
            if (_ends.TryGetValue(position, out List<ProjectedBookmark>? ends)) {
                foreach (ProjectedBookmark bookmark in ends.Where(bookmark => !bookmark.IsZeroLength)) {
                    Append(tree, target, CreateBookmarkEnd(tree, bookmark.Id));
                }
            }

            if (_starts.TryGetValue(position, out List<ProjectedBookmark>? starts)) {
                foreach (ProjectedBookmark bookmark in starts) {
                    Append(tree, target, CreateBookmarkStart(tree, bookmark.Id, bookmark.Name));
                }
            }

            if (ends != null) {
                foreach (ProjectedBookmark bookmark in ends.Where(bookmark => bookmark.IsZeroLength)) {
                    Append(tree, target, CreateBookmarkEnd(tree, bookmark.Id));
                }
            }

            _emittedPositions.Add(position);
        }

        internal void EmitRemaining(OfficeXmlTree tree, int target) {
            foreach (int position in _starts.Keys.Concat(_ends.Keys).Distinct().OrderBy(position => position)) {
                EmitAt(tree, target, position);
            }
        }

        internal static int CreateBookmarkStart(OfficeXmlTree tree, string id, string name) {
            int element = Element(tree, "bookmarkStart");
            Set(tree, element, "id", id);
            Set(tree, element, "name", name);
            return element;
        }

        internal static int CreateBookmarkEnd(OfficeXmlTree tree, string id) {
            int element = Element(tree, "bookmarkEnd");
            Set(tree, element, "id", id);
            return element;
        }

        private sealed class ProjectedBookmark {
            internal ProjectedBookmark(LegacyDocBookmark bookmark, bool projectStart, bool projectEnd) {
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

    /// <summary>The footnotes, endnotes and comments of a story, by the character position of their reference.</summary>
    internal sealed class LegacyDocNoteProjection {
        private readonly IReadOnlyDictionary<int, LegacyDocFootnote> _footnotesByReferencePosition;
        private readonly IReadOnlyDictionary<int, LegacyDocEndnote> _endnotesByReferencePosition;
        private readonly IReadOnlyDictionary<int, LegacyDocComment> _commentsByReferencePosition;
        private LegacyDocNoteProjection? _withoutReferences;

        private LegacyDocNoteProjection(
            IReadOnlyDictionary<int, LegacyDocFootnote> footnotesByReferencePosition,
            IReadOnlyDictionary<int, LegacyDocEndnote> endnotesByReferencePosition,
            IReadOnlyDictionary<int, LegacyDocComment> commentsByReferencePosition,
            LegacyDocStyleSheet styleSheet,
            CancellationToken cancellationToken = default) {
            _footnotesByReferencePosition = footnotesByReferencePosition;
            _endnotesByReferencePosition = endnotesByReferencePosition;
            _commentsByReferencePosition = commentsByReferencePosition;
            StyleSheet = styleSheet;
            CancellationToken = cancellationToken;
        }

        internal static LegacyDocNoteProjection Create(IReadOnlyList<LegacyDocFootnote> footnotes, IReadOnlyList<LegacyDocEndnote> endnotes, IReadOnlyList<LegacyDocComment> comments, LegacyDocStyleSheet styleSheet, CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            return new LegacyDocNoteProjection(
                footnotes.GroupBy(footnote => footnote.ReferenceCharacterPosition).ToDictionary(group => group.Key, group => group.First()),
                endnotes.GroupBy(endnote => endnote.ReferenceCharacterPosition).ToDictionary(group => group.Key, group => group.First()),
                comments.GroupBy(comment => comment.ReferenceCharacterPosition).ToDictionary(group => group.Key, group => group.First()),
                styleSheet,
                cancellationToken);
        }

        internal static LegacyDocNoteProjection Empty { get; } = new(
            new Dictionary<int, LegacyDocFootnote>(),
            new Dictionary<int, LegacyDocEndnote>(),
            new Dictionary<int, LegacyDocComment>(),
            LegacyDocStyleSheet.Empty);

        internal LegacyDocStyleSheet StyleSheet { get; }

        internal CancellationToken CancellationToken { get; }

        internal LegacyDocNoteProjection WithoutReferences => _withoutReferences ??= new LegacyDocNoteProjection(
            Empty._footnotesByReferencePosition,
            Empty._endnotesByReferencePosition,
            Empty._commentsByReferencePosition,
            Empty.StyleSheet,
            CancellationToken);

        internal bool TryGetFootnote(int referenceCharacterPosition, out LegacyDocFootnote? footnote) =>
            _footnotesByReferencePosition.TryGetValue(referenceCharacterPosition, out footnote);

        internal bool TryGetEndnote(int referenceCharacterPosition, out LegacyDocEndnote? endnote) =>
            _endnotesByReferencePosition.TryGetValue(referenceCharacterPosition, out endnote);

        internal bool TryGetComment(int referenceCharacterPosition, out LegacyDocComment? comment) =>
            _commentsByReferencePosition.TryGetValue(referenceCharacterPosition, out comment);
    }
}
