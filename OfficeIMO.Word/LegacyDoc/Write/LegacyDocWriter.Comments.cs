using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private const int LegacyCommentInitialsMaximumLength = 9;
        private const int LegacyCommentAuthorRecordSize = 30;

        private static LegacyDocWritableComments ReadSupportedComments(
            MainDocumentPart mainPart,
            IReadOnlyDictionary<string, ushort> styleIndexes) {
            Comments? comments = mainPart.WordprocessingCommentsPart?.Comments;
            if (comments == null) {
                return LegacyDocWritableComments.Empty;
            }

            var stories = new Dictionary<string, LegacyDocWritableCommentStory>(StringComparer.Ordinal);
            foreach (Comment comment in comments.Elements<Comment>()) {
                string? id = comment.Id?.Value;
                if (string.IsNullOrWhiteSpace(id)) {
                    throw new NotSupportedException("Native DOC saving supports comments only when every comment has an identifier.");
                }

                if (stories.ContainsKey(id!)) {
                    throw new NotSupportedException($"Native DOC saving cannot write duplicate comment id '{id}'.");
                }

                string initials = string.IsNullOrWhiteSpace(comment.Initials?.Value)
                    ? "DOC"
                    : comment.Initials!.Value!;
                if (initials.Length > LegacyCommentInitialsMaximumLength) {
                    throw new NotSupportedException(
                        $"Native DOC saving supports comment initials up to {LegacyCommentInitialsMaximumLength} characters. "
                        + $"Comment id '{id}' uses {initials.Length} characters.");
                }

                stories.Add(id!, new LegacyDocWritableCommentStory(
                    id!,
                    initials,
                    ReadSimpleCommentStory(comment, id!, styleIndexes)));
            }

            return stories.Count == 0
                ? LegacyDocWritableComments.Empty
                : new LegacyDocWritableComments(stories);
        }

        private static LegacyDocWritableNoteStory ReadSimpleCommentStory(
            Comment comment,
            string id,
            IReadOnlyDictionary<string, ushort> styleIndexes) {
            var builder = new StringBuilder();
            var runs = new List<LegacyDocWritableRun>();
            var formattedParagraphs = new List<LegacyDocWritableParagraph>();
            var bookmarks = new LegacyDocWritableBookmarksBuilder();
            bool hasText = false;
            foreach (OpenXmlElement child in comment.ChildElements) {
                if (child is not Paragraph paragraph) {
                    throw new NotSupportedException(
                        $"Native DOC saving supports comment id '{id}' only when it contains simple paragraphs. "
                        + $"Unsupported comment element: {child.LocalName}.");
                }

                int paragraphStart = builder.Length;
                LegacyDocWritableFormatting paragraphMarkFormatting =
                    ReadSupportedParagraphMarkRunFormatting(paragraph.ParagraphProperties);
                LegacyDocWritableParagraphFormatting paragraphFormatting;
                try {
                    paragraphFormatting = ReadSupportedParagraphFormatting(
                        paragraph.ParagraphProperties,
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
                    bookmarks);
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
            Paragraph paragraph,
            string id,
            StringBuilder builder,
            List<LegacyDocWritableRun> runs,
            LegacyDocWritableBookmarksBuilder bookmarks) {
            foreach (OpenXmlElement child in paragraph.ChildElements) {
                switch (child) {
                    case ParagraphProperties:
                        break;
                    case Run run:
                        AppendSimpleCommentRun(builder, runs, run, id);
                        break;
                    case BookmarkStart bookmarkStart:
                        bookmarks.AddStart(bookmarkStart, builder.Length);
                        break;
                    case BookmarkEnd bookmarkEnd:
                        bookmarks.AddEnd(bookmarkEnd, builder.Length);
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

        private static void AppendSimpleCommentRun(
            StringBuilder builder,
            List<LegacyDocWritableRun> runs,
            Run run,
            string id) {
            LegacyDocWritableFormatting formatting = ReadSupportedRunFormatting(run.RunProperties);
            foreach (OpenXmlElement child in run.ChildElements) {
                switch (child) {
                    case RunProperties:
                    case LastRenderedPageBreak:
                        break;
                    case Text text:
                        AppendFormattedText(builder, runs, text.Text, formatting);
                        break;
                    case TabChar:
                        AppendFormattedText(builder, runs, "\t", formatting);
                        break;
                    case CarriageReturn:
                        AppendFormattedText(
                            builder,
                            runs,
                            LegacyDocSpecialCharacters.TextWrappingBreak.ToString(),
                            formatting);
                        break;
                    case NoBreakHyphen:
                        AppendFormattedText(
                            builder,
                            runs,
                            LegacyDocSpecialCharacters.NoBreakHyphen.ToString(),
                            formatting);
                        break;
                    case SoftHyphen:
                        AppendFormattedText(
                            builder,
                            runs,
                            LegacyDocSpecialCharacters.SoftHyphen.ToString(),
                            formatting);
                        break;
                    case Break breakNode:
                        AppendSupportedBreak(builder, runs, breakNode, formatting);
                        break;
                    default:
                        throw new NotSupportedException(
                            $"Native DOC saving supports comment id '{id}' only with text, tabs, hyphens, and simple breaks. "
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

            internal void BindBodyReferences(Body body, string bodyText) {
                CommentReference[] references = body.Descendants<CommentReference>().ToArray();
                int[] markerPositions = bodyText
                    .Select((character, index) => new { character, index })
                    .Where(item => item.character == LegacyDocCommentReader.CommentReferenceCharacter)
                    .Select(item => item.index)
                    .ToArray();
                if (references.Length != markerPositions.Length) {
                    throw new NotSupportedException(
                        "Native DOC saving could not match body comment-reference elements to encoded comment markers.");
                }

                var referencedIds = new HashSet<string>(StringComparer.Ordinal);
                for (int index = 0; index < references.Length; index++) {
                    string? id = references[index].Id?.Value;
                    if (string.IsNullOrWhiteSpace(id) || !_storiesById.TryGetValue(id!, out LegacyDocWritableCommentStory story)) {
                        throw new NotSupportedException(
                            $"Native DOC saving cannot write comment reference id '{id ?? "(missing)"}' because the comment body is missing.");
                    }

                    if (!referencedIds.Add(id!)) {
                        throw new NotSupportedException(
                            $"Native DOC saving currently supports one body reference per comment id. Duplicate reference id '{id}' was found.");
                    }

                    int rangeStartCount = body.Descendants<CommentRangeStart>()
                        .Count(marker => string.Equals(marker.Id?.Value, id, StringComparison.Ordinal));
                    int rangeEndCount = body.Descendants<CommentRangeEnd>()
                        .Count(marker => string.Equals(marker.Id?.Value, id, StringComparison.Ordinal));
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
                            run.Formatting));
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
