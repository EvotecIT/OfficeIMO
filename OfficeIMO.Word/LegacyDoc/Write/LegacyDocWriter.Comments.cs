using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private const int LegacyCommentInitialsMaximumLength = 9;
        private const int LegacyCommentAuthorRecordSize = 30;

        private static LegacyDocWritableComments ReadSupportedComments(
            LegacyDocSourcePackage source,
            LegacyDocWritablePictures pictures,
            IReadOnlyDictionary<string, ushort> styleIndexes) {
            string? commentsPart = source.CommentsPart;
            LegacyDocElement comments = source.Root(commentsPart);
            if (comments.IsNone) {
                return LegacyDocWritableComments.Empty;
            }

            var stories = new Dictionary<string, LegacyDocWritableCommentStory>(StringComparer.Ordinal);
            foreach (LegacyDocElement comment in comments.Elements("comment")) {
                string? id = comment.Attr("id");
                if (string.IsNullOrWhiteSpace(id)) {
                    throw new NotSupportedException("Native DOC saving supports comments only when every comment has an identifier.");
                }

                if (stories.ContainsKey(id!)) {
                    throw new NotSupportedException($"Native DOC saving cannot write duplicate comment id '{id}'.");
                }

                string? initialsValue = comment.Attr("initials");
                string initials = string.IsNullOrWhiteSpace(initialsValue)
                    ? "DOC"
                    : initialsValue!;
                if (initials.Length > LegacyCommentInitialsMaximumLength) {
                    throw new NotSupportedException(
                        $"Native DOC saving supports comment initials up to {LegacyCommentInitialsMaximumLength} characters. "
                        + $"Comment id '{id}' uses {initials.Length} characters.");
                }

                stories.Add(id!, new LegacyDocWritableCommentStory(
                    id!,
                    initials,
                    ReadSimpleCommentStory(comment, id!, source.Part(commentsPart!), pictures, styleIndexes)));
            }

            return stories.Count == 0
                ? LegacyDocWritableComments.Empty
                : new LegacyDocWritableComments(stories);
        }

        private static LegacyDocWritableNoteStory ReadSimpleCommentStory(
            LegacyDocElement comment,
            string id,
            LegacyDocPart relationshipOwner,
            LegacyDocWritablePictures pictures,
            IReadOnlyDictionary<string, ushort> styleIndexes) {
            var builder = new StringBuilder();
            var runs = new List<LegacyDocWritableRun>();
            var formattedParagraphs = new List<LegacyDocWritableParagraph>();
            var bookmarks = new LegacyDocWritableBookmarksBuilder();
            bool hasText = false;
            foreach (LegacyDocElement child in comment.Elements()) {
                if (!child.Is("p")) {
                    throw new NotSupportedException(
                        $"Native DOC saving supports comment id '{id}' only when it contains simple paragraphs. "
                        + $"Unsupported comment element: {child.LocalName}.");
                }

                LegacyDocElement paragraph = child;

                int paragraphStart = builder.Length;
                LegacyDocWritableFormatting paragraphMarkFormatting =
                    ReadSupportedParagraphMarkRunFormatting(paragraph.Child("pPr"));
                LegacyDocWritableParagraphFormatting paragraphFormatting;
                try {
                    paragraphFormatting = ReadSupportedParagraphFormatting(
                        paragraph.Child("pPr"),
                        styleIndexes);
                } catch (NotSupportedException exception) {
                    throw new NotSupportedException(
                        $"Native DOC saving supports comment id '{id}' only with supported paragraph formatting. "
                        + exception.Message,
                        exception);
                }

                AppendSimpleCommentParagraph(
                    paragraph,
                    id,
                    builder,
                    runs,
                    bookmarks,
                    relationshipOwner,
                    pictures);
                if (builder.Length > paragraphStart) {
                    hasText = true;
                }

                builder.Append('\r');
                AddParagraphMarkRunFormatting(runs, builder.Length - 1, paragraphMarkFormatting);
                if (paragraphFormatting.HasFormatting) {
                    formattedParagraphs.Add(new LegacyDocWritableParagraph(
                        paragraphStart,
                        builder.Length - paragraphStart,
                        paragraphFormatting));
                }
            }

            if (!hasText) {
                throw new NotSupportedException($"Native DOC saving cannot write empty comment id '{id}'.");
            }

            return new LegacyDocWritableNoteStory(
                builder.ToString(),
                runs,
                formattedParagraphs,
                bookmarks.Create());
        }

        private static void AppendSimpleCommentParagraph(
            LegacyDocElement paragraph,
            string id,
            StringBuilder builder,
            List<LegacyDocWritableRun> runs,
            LegacyDocWritableBookmarksBuilder bookmarks,
            LegacyDocPart relationshipOwner,
            LegacyDocWritablePictures pictures) {
            foreach (LegacyDocElement child in paragraph.Elements()) {
                switch (child.WName) {
                    case "pPr":
                        break;
                    case "r":
                        AppendSimpleCommentRun(builder, runs, child, id, relationshipOwner, pictures);
                        break;
                    case "ins":
                        AppendSimpleCommentRevision(builder, runs, child, LegacyDocRevisionKind.Inserted, id, relationshipOwner, pictures);
                        break;
                    case "del":
                        AppendSimpleCommentRevision(builder, runs, child, LegacyDocRevisionKind.Deleted, id, relationshipOwner, pictures);
                        break;
                    case "bookmarkStart":
                        bookmarks.AddStart(child, builder.Length);
                        break;
                    case "bookmarkEnd":
                        bookmarks.AddEnd(child, builder.Length);
                        break;
                    default:
                        if (IsIgnorableParagraphMarkup(child)) {
                            break;
                        }

                        throw new NotSupportedException(
                            $"Native DOC saving supports comment id '{id}' only with text runs and bookmarks. "
                            + $"Unsupported comment paragraph element: {child.LocalName}.");
                }
            }
        }

        private static void AppendSimpleCommentRevision(
            StringBuilder builder,
            List<LegacyDocWritableRun> runs,
            LegacyDocElement revisionElement,
            LegacyDocRevisionKind revisionKind,
            string id,
            LegacyDocPart relationshipOwner,
            LegacyDocWritablePictures pictures) {
            LegacyDocRevision revision = ReadSupportedRevision(revisionElement, revisionKind);
            foreach (LegacyDocElement child in revisionElement.Elements()) {
                if (child.Is("r")) {
                    AppendSimpleCommentRun(builder, runs, child, id, relationshipOwner, pictures, revision);
                    continue;
                }

                throw new NotSupportedException($"Native DOC saving supports tracked insertions and deletions in comment id '{id}' only when they contain text runs. Unsupported revision element: {child.LocalName}.");
            }
        }

        private static void AppendSimpleCommentRun(
            StringBuilder builder,
            List<LegacyDocWritableRun> runs,
            LegacyDocElement run,
            string id,
            LegacyDocPart relationshipOwner,
            LegacyDocWritablePictures pictures,
            LegacyDocRevision revision = default) {
            LegacyDocWritableFormatting formatting = ReadSupportedRunFormatting(run.Child("rPr")).WithRevision(revision);
            foreach (LegacyDocElement child in run.Elements()) {
                switch (child.WName) {
                    case "rPr":
                    case "lastRenderedPageBreak":
                        break;
                    case "t":
                        AppendFormattedText(builder, runs, child.Text, formatting);
                        break;
                    case "delText":
                        AppendFormattedText(builder, runs, child.Text, formatting);
                        break;
                    case "tab":
                        AppendFormattedText(builder, runs, "\t", formatting);
                        break;
                    case "cr":
                        AppendFormattedText(
                            builder,
                            runs,
                            LegacyDocSpecialCharacters.TextWrappingBreak.ToString(),
                            formatting);
                        break;
                    case "noBreakHyphen":
                        AppendFormattedText(
                            builder,
                            runs,
                            LegacyDocSpecialCharacters.NoBreakHyphen.ToString(),
                            formatting);
                        break;
                    case "softHyphen":
                        AppendFormattedText(
                            builder,
                            runs,
                            LegacyDocSpecialCharacters.SoftHyphen.ToString(),
                            formatting);
                        break;
                    case "br":
                        AppendSupportedBreak(builder, runs, child, formatting);
                        break;
                    case "drawing":
                        int picturePosition = builder.Length;
                        int pictureDataOffset = pictures.AddInlinePicture(child, relationshipOwner);
                        builder.Append('\u0001');
                        runs.Add(new LegacyDocWritableRun(
                            picturePosition,
                            1,
                            LegacyDocWritableFormatting.SpecialCharacter.WithRevision(revision),
                            pictureDataOffset));
                        break;
                    default:
                        throw new NotSupportedException(
                            $"Native DOC saving supports comment id '{id}' only with text, embedded inline pictures, tabs, hyphens, and simple breaks. "
                            + $"Unsupported comment run element: {child.LocalName}.");
                }
            }
        }

        private sealed class LegacyDocWritableComments {
            internal static readonly LegacyDocWritableComments Empty = new(
                new Dictionary<string, LegacyDocWritableCommentStory>(StringComparer.Ordinal));

            private readonly IReadOnlyDictionary<string, LegacyDocWritableCommentStory> _storiesById;
            private readonly List<LegacyDocWritableCommentReference> _references = new();

            internal LegacyDocWritableComments(
                IReadOnlyDictionary<string, LegacyDocWritableCommentStory> storiesById) {
                _storiesById = storiesById;
            }

            internal void BindBodyReferences(LegacyDocElement body, string bodyText) {
                const int maximumCommentReferences = 100_000;
                LegacyDocElement[] references = body.Descendants("commentReference").ToList().ToArray();
                if (references.Length > maximumCommentReferences) {
                    throw new NotSupportedException(
                        $"Native DOC saving supports at most {maximumCommentReferences} comment references.");
                }

                var markerPositions = new List<int>(references.Length);
                for (int index = 0; index < bodyText.Length; index++) {
                    if (bodyText[index] == LegacyDocCommentReader.CommentReferenceCharacter) {
                        markerPositions.Add(index);
                    }
                }
                if (references.Length != markerPositions.Count) {
                    throw new NotSupportedException(
                        "Native DOC saving could not match body comment-reference elements to encoded comment markers.");
                }

                IReadOnlyDictionary<string, int> rangeStartCounts = CountCommentMarkers(
                    body.Descendants("commentRangeStart").Select(marker => marker.Attr("id")));
                IReadOnlyDictionary<string, int> rangeEndCounts = CountCommentMarkers(
                    body.Descendants("commentRangeEnd").Select(marker => marker.Attr("id")));

                var referencedIds = new HashSet<string>(StringComparer.Ordinal);
                for (int index = 0; index < references.Length; index++) {
                    string? id = references[index].Attr("id");
                    if (string.IsNullOrWhiteSpace(id) || !_storiesById.TryGetValue(id!, out LegacyDocWritableCommentStory story)) {
                        throw new NotSupportedException(
                            $"Native DOC saving cannot write comment reference id '{id ?? "(missing)"}' because the comment body is missing.");
                    }

                    if (!referencedIds.Add(id!)) {
                        throw new NotSupportedException(
                            $"Native DOC saving currently supports one body reference per comment id. Duplicate reference id '{id}' was found.");
                    }

                    int rangeStartCount = rangeStartCounts.TryGetValue(id!, out int starts) ? starts : 0;
                    int rangeEndCount = rangeEndCounts.TryGetValue(id!, out int ends) ? ends : 0;
                    if (rangeStartCount != 1 || rangeEndCount != 1) {
                        throw new NotSupportedException(
                            $"Native DOC saving supports comment id '{id}' only with one matching range start and range end in the body.");
                    }

                    _references.Add(new LegacyDocWritableCommentReference(story, markerPositions[index]));
                }

                foreach (string id in _storiesById.Keys) {
                    if (!referencedIds.Contains(id)) {
                        throw new NotSupportedException($"Native DOC saving cannot write unreferenced comment id '{id}'.");
                    }
                }
            }

            private static IReadOnlyDictionary<string, int> CountCommentMarkers(
                IEnumerable<string?> markerIds) {
                var counts = new Dictionary<string, int>(StringComparer.Ordinal);
                foreach (string? markerId in markerIds) {
                    if (string.IsNullOrWhiteSpace(markerId)) continue;
                    counts[markerId!] = counts.TryGetValue(markerId!, out int count)
                        ? checked(count + 1)
                        : 1;
                }
                return counts;
            }

            internal LegacyDocWritableCommentStories CreateStories() {
                if (_references.Count == 0) {
                    return LegacyDocWritableCommentStories.Empty;
                }

                var text = new StringBuilder();
                var textPositions = new List<int>(_references.Count + 1);
                var runs = new List<LegacyDocWritableRun>();
                var paragraphs = new List<LegacyDocWritableParagraph>();
                var bookmarks = new LegacyDocWritableBookmarksBuilder();
                for (int index = 0; index < _references.Count; index++) {
                    LegacyDocWritableCommentReference reference = _references[index];
                    textPositions.Add(text.Length);
                    foreach (LegacyDocWritableRun run in reference.Story.Content.FormattedRuns) {
                        runs.Add(new LegacyDocWritableRun(
                            text.Length + run.StartCharacter,
                            run.Length,
                            run.Formatting,
                            run.PictureDataOffset, run.FieldLocked));
                    }

                    foreach (LegacyDocWritableParagraph paragraph in reference.Story.Content.FormattedParagraphs) {
                        paragraphs.Add(new LegacyDocWritableParagraph(
                            text.Length + paragraph.StartCharacter,
                            paragraph.Length,
                            paragraph.Formatting));
                    }

                    bookmarks.AddRange(reference.Story.Content.Bookmarks, text.Length);
                    text.Append(reference.Story.Content.Text);
                }

                textPositions.Add(text.Length);
                return new LegacyDocWritableCommentStories(
                    text.ToString(),
                    CreateCommentReferencePlc(_references),
                    CreateCommentTextPlc(textPositions),
                    runs,
                    paragraphs,
                    bookmarks.Create());
            }

            private static byte[] CreateCommentReferencePlc(
                IReadOnlyList<LegacyDocWritableCommentReference> references) {
                byte[] plc = new byte[((references.Count + 1) * 4)
                    + (references.Count * LegacyCommentAuthorRecordSize)];
                for (int index = 0; index < references.Count; index++) {
                    WriteInt32(plc, index * 4, references[index].CharacterPosition);
                }

                WriteInt32(plc, references.Count * 4, references[references.Count - 1].CharacterPosition + 1);
                int authorOffset = (references.Count + 1) * 4;
                for (int index = 0; index < references.Count; index++) {
                    WriteCommentAuthorRecord(
                        plc,
                        authorOffset + (index * LegacyCommentAuthorRecordSize),
                        references[index].Story.Initials);
                }

                return plc;
            }

            private static void WriteCommentAuthorRecord(byte[] bytes, int offset, string initials) {
                WriteUInt16(bytes, offset, checked((ushort)initials.Length));
                for (int index = 0; index < initials.Length; index++) {
                    WriteUInt16(bytes, offset + 2 + (index * 2), initials[index]);
                }

                WriteUInt16(bytes, offset + 20, 0);
                WriteUInt16(bytes, offset + 22, 0);
                WriteUInt16(bytes, offset + 24, 0);
                WriteInt32(bytes, offset + 26, -1);
            }

            private static byte[] CreateCommentTextPlc(IReadOnlyList<int> textPositions) {
                byte[] plc = new byte[textPositions.Count * 4];
                for (int index = 0; index < textPositions.Count; index++) {
                    WriteInt32(plc, index * 4, textPositions[index]);
                }

                return plc;
            }
        }

        private readonly struct LegacyDocWritableCommentStory {
            internal LegacyDocWritableCommentStory(
                string id,
                string initials,
                LegacyDocWritableNoteStory content) {
                Id = id;
                Initials = initials;
                Content = content;
            }

            internal string Id { get; }

            internal string Initials { get; }

            internal LegacyDocWritableNoteStory Content { get; }
        }

        private readonly struct LegacyDocWritableCommentReference {
            internal LegacyDocWritableCommentReference(
                LegacyDocWritableCommentStory story,
                int characterPosition) {
                Story = story;
                CharacterPosition = characterPosition;
            }

            internal LegacyDocWritableCommentStory Story { get; }

            internal int CharacterPosition { get; }
        }

        private readonly struct LegacyDocWritableCommentStories {
            internal static readonly LegacyDocWritableCommentStories Empty = new(
                string.Empty,
                Array.Empty<byte>(),
                Array.Empty<byte>(),
                Array.Empty<LegacyDocWritableRun>(),
                Array.Empty<LegacyDocWritableParagraph>(),
                LegacyDocWritableBookmarks.Empty);

            internal LegacyDocWritableCommentStories(
                string text,
                byte[] plcfandRef,
                byte[] plcfandTxt,
                IReadOnlyList<LegacyDocWritableRun> formattedRuns,
                IReadOnlyList<LegacyDocWritableParagraph> formattedParagraphs,
                LegacyDocWritableBookmarks bookmarks) {
                Text = text;
                PlcfandRef = plcfandRef;
                PlcfandTxt = plcfandTxt;
                FormattedRuns = formattedRuns;
                FormattedParagraphs = formattedParagraphs;
                Bookmarks = bookmarks;
            }

            internal string Text { get; }

            internal byte[] PlcfandRef { get; }

            internal byte[] PlcfandTxt { get; }

            internal IReadOnlyList<LegacyDocWritableRun> FormattedRuns { get; }

            internal IReadOnlyList<LegacyDocWritableParagraph> FormattedParagraphs { get; }

            internal LegacyDocWritableBookmarks Bookmarks { get; }
        }
    }
}
