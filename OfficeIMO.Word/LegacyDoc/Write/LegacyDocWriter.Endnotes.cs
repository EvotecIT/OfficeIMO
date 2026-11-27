using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static LegacyDocWritableEndnotes ReadSupportedEndnotes(LegacyDocSourcePackage source, LegacyDocWritablePictures pictures, IReadOnlyDictionary<string, ushort> styleIndexes) {
            LegacyDocElement endnotes = source.Endnotes;
            if (endnotes.IsNone) {
                return LegacyDocWritableEndnotes.Empty;
            }

            IReadOnlyDictionary<string, ushort> noteStyleIndexes = CreateNoteParagraphStyleIndexes(styleIndexes, "EndnoteText");
            LegacyDocPart relationshipOwner = source.Part(source.EndnotesPart!);
            var stories = new Dictionary<long, LegacyDocWritableNoteStory>();
            foreach (LegacyDocElement endnote in endnotes.Elements("endnote")) {
                if (!IsUserEndnote(endnote)) {
                    continue;
                }

                long? id = LegacyDocValues.Int64(endnote.Attr("id"));
                if (id == null || id.Value <= 0) {
                    throw new NotSupportedException("Native DOC saving supports endnotes only when every user endnote has a positive identifier.");
                }

                if (stories.ContainsKey(id.Value)) {
                    throw new NotSupportedException($"Native DOC saving cannot write duplicate endnote id '{id.Value}'.");
                }

                stories.Add(id.Value, ReadSimpleEndnoteStory(endnote, id.Value, relationshipOwner, pictures, noteStyleIndexes));
            }

            return stories.Count == 0
                ? LegacyDocWritableEndnotes.Empty
                : new LegacyDocWritableEndnotes(stories);
        }

        private static LegacyDocWritableNoteStory ReadSimpleEndnoteStory(LegacyDocElement endnote, long id, LegacyDocPart relationshipOwner, LegacyDocWritablePictures pictures, IReadOnlyDictionary<string, ushort> noteStyleIndexes) {
            var builder = new StringBuilder();
            var runs = new List<LegacyDocWritableRun>();
            var formattedParagraphs = new List<LegacyDocWritableParagraph>();
            var bookmarks = new LegacyDocWritableBookmarksBuilder();
            builder.Append(LegacyDocFootnoteReader.FootnoteReferenceCharacter);
            if (!HasExplicitNoteReferenceSeparator(endnote)) builder.Append(' ');
            bool hasBodyText = false;
            bool isFirstParagraph = true;
            foreach (LegacyDocElement child in endnote.Elements()) {
                AppendSimpleEndnoteStoryChild(
                    child,
                    id,
                    relationshipOwner,
                    pictures,
                    noteStyleIndexes,
                    builder,
                    runs,
                    formattedParagraphs,
                    bookmarks,
                    ref hasBodyText,
                    ref isFirstParagraph,
                    "endnote");
            }

            if (!hasBodyText) {
                throw new NotSupportedException($"Native DOC saving cannot write empty endnote id '{id}'.");
            }

            builder.Append('\r');
            return new LegacyDocWritableNoteStory(builder.ToString(), runs, formattedParagraphs, bookmarks.Create());
        }

        private static void AppendSimpleEndnoteStoryChild(
            LegacyDocElement child,
            long id,
            LegacyDocPart relationshipOwner,
            LegacyDocWritablePictures pictures,
            IReadOnlyDictionary<string, ushort> noteStyleIndexes,
            StringBuilder builder,
            List<LegacyDocWritableRun> runs,
            List<LegacyDocWritableParagraph> formattedParagraphs,
            LegacyDocWritableBookmarksBuilder bookmarks,
            ref bool hasBodyText,
            ref bool isFirstParagraph,
            string containerDescription) {
            switch (child.WName) {
                case "p":
                    int paragraphStart = isFirstParagraph ? 0 : builder.Length;
                    LegacyDocWritableFormatting paragraphMarkFormatting = ReadSupportedParagraphMarkRunFormatting(child.Child("pPr"));
                    LegacyDocWritableParagraphFormatting paragraphFormatting = ReadSimpleEndnoteParagraph(child, id, runs, bookmarks, builder.Length, isFirstParagraph, relationshipOwner, pictures, noteStyleIndexes, out string paragraphText);
                    if (!string.IsNullOrEmpty(paragraphText)) {
                        hasBodyText = true;
                    }

                    builder.Append(paragraphText);
                    builder.Append('\r');
                    AddParagraphMarkRunFormatting(runs, builder.Length - 1, paragraphMarkFormatting);
                    if (paragraphFormatting.HasFormatting) {
                        formattedParagraphs.Add(new LegacyDocWritableParagraph(paragraphStart, builder.Length - paragraphStart, paragraphFormatting));
                    }

                    isFirstParagraph = false;
                    break;
                case "sdt":
                    AppendSimpleEndnoteContentControl(
                        child,
                        id,
                        relationshipOwner,
                        pictures,
                        noteStyleIndexes,
                        builder,
                        runs,
                        formattedParagraphs,
                        bookmarks,
                        ref hasBodyText,
                        ref isFirstParagraph);
                    break;
                case "bookmarkStart":
                    bookmarks.AddStart(child, builder.Length);
                    break;
                case "bookmarkEnd":
                    bookmarks.AddEnd(child, builder.Length);
                    break;
                default:
                    throw new NotSupportedException($"Native DOC saving supports simple endnote paragraphs, content controls, and bookmarks only. Unsupported {containerDescription} element: {child.LocalName}.");
            }
        }

        private static void AppendSimpleEndnoteContentControl(
            LegacyDocElement sdtBlock,
            long id,
            LegacyDocPart relationshipOwner,
            LegacyDocWritablePictures pictures,
            IReadOnlyDictionary<string, ushort> noteStyleIndexes,
            StringBuilder builder,
            List<LegacyDocWritableRun> runs,
            List<LegacyDocWritableParagraph> formattedParagraphs,
            LegacyDocWritableBookmarksBuilder bookmarks,
            ref bool hasBodyText,
            ref bool isFirstParagraph) {
            LegacyDocElement contentBlock = sdtBlock.Child("sdtContent");
            if (contentBlock.IsNone) {
                throw new NotSupportedException($"Native DOC saving supports endnote id '{id}' content controls only when they contain simple paragraphs and bookmarks.");
            }

            foreach (LegacyDocElement child in contentBlock.Elements()) {
                AppendSimpleEndnoteStoryChild(
                    child,
                    id,
                    relationshipOwner,
                    pictures,
                    noteStyleIndexes,
                    builder,
                    runs,
                    formattedParagraphs,
                    bookmarks,
                    ref hasBodyText,
                    ref isFirstParagraph,
                    "endnote content control");
            }
        }

        private static LegacyDocWritableParagraphFormatting ReadSimpleEndnoteParagraph(LegacyDocElement paragraph, long id, List<LegacyDocWritableRun> runs, LegacyDocWritableBookmarksBuilder bookmarks, int storyStart, bool isFirstParagraph, LegacyDocPart relationshipOwner, LegacyDocWritablePictures pictures, IReadOnlyDictionary<string, ushort> noteStyleIndexes, out string paragraphText) {
            var builder = new StringBuilder();
            LegacyDocWritableParagraphFormatting paragraphFormatting = ReadSupportedNoteParagraphFormatting(paragraph, id, "endnote", noteStyleIndexes);
            if (isFirstParagraph && paragraphFormatting.HasFormatting && paragraphFormatting.StyleIndex == null) {
                paragraphFormatting = paragraphFormatting.WithStyleIndex(NoteTextParagraphStyleIndex);
            }

            LegacyDocElement[] children = paragraph.Elements().ToArray();
            for (int index = 0; index < children.Length; index++) {
                LegacyDocElement child = children[index];
                if (child.Is(OfficeXmlNamespaces.Math, "oMath") || child.Is(OfficeXmlNamespaces.Math, "oMathPara")) {
                    AppendMathEquationNoteField(builder, runs, bookmarks, child, storyStart);
                    continue;
                }

                switch (child.WName) {
                    case "pPr":
                        break;
                    case "r":
                        if (IsComplexFieldBeginRun(child)) {
                            AppendSupportedNoteComplexPageNumberField(children, ref index, builder, runs, bookmarks, storyStart);
                        } else {
                            AppendSimpleEndnoteRun(builder, runs, child, id, storyStart, relationshipOwner, pictures);
                        }

                        break;
                    case "ins":
                        AppendSimpleEndnoteRevision(builder, runs, child, LegacyDocRevisionKind.Inserted, id, storyStart, relationshipOwner, pictures);
                        break;
                    case "del":
                        AppendSimpleEndnoteRevision(builder, runs, child, LegacyDocRevisionKind.Deleted, id, storyStart, relationshipOwner, pictures);
                        break;
                    case "hyperlink":
                        AppendSupportedNoteHyperlinkText(builder, runs, bookmarks, child, relationshipOwner, id, "endnote", storyStart);
                        break;
                    case "fldSimple":
                        AppendSupportedNoteFieldFromSimpleField(builder, runs, bookmarks, child, storyStart);
                        break;
                    case "sdt":
                        AppendSupportedEndnoteInlineContentControl(builder, runs, bookmarks, child, relationshipOwner, pictures, id, storyStart);
                        break;
                    case "bookmarkStart":
                        bookmarks.AddStart(child, storyStart + builder.Length);
                        break;
                    case "bookmarkEnd":
                        bookmarks.AddEnd(child, storyStart + builder.Length);
                        break;
                    default:
                        if (IsIgnorableParagraphMarkup(child)) {
                            break;
                        }

                        throw new NotSupportedException($"Native DOC saving supports simple endnote paragraphs only with text runs, {SupportedFieldNames} simple fields, bookmarks, inline content controls, and simple hyperlinks. Unsupported endnote paragraph element: {child.LocalName}.");
                }
            }

            paragraphText = builder.ToString();
            return paragraphFormatting;
        }

        private static void AppendSimpleEndnoteRevision(
            StringBuilder builder,
            List<LegacyDocWritableRun> runs,
            LegacyDocElement revisionElement,
            LegacyDocRevisionKind revisionKind,
            long id,
            int storyStart,
            LegacyDocPart relationshipOwner,
            LegacyDocWritablePictures pictures) {
            LegacyDocRevision revision = ReadSupportedRevision(revisionElement, revisionKind);
            foreach (LegacyDocElement child in revisionElement.Elements()) {
                if (child.Is("r")) {
                    AppendSimpleEndnoteRun(builder, runs, child, id, storyStart, relationshipOwner, pictures, revision);
                    continue;
                }

                throw new NotSupportedException($"Native DOC saving supports tracked insertions and deletions in endnote id '{id}' only when they contain text runs. Unsupported revision element: {child.LocalName}.");
            }
        }

        private static void AppendSimpleEndnoteRun(StringBuilder builder, List<LegacyDocWritableRun> runs, LegacyDocElement run, long id, int storyStart, LegacyDocPart relationshipOwner, LegacyDocWritablePictures pictures, LegacyDocRevision revision = default) {
            if (IsEndnoteReferenceMarkRun(run)) {
                return;
            }

            LegacyDocWritableFormatting formatting = ReadSupportedRunFormatting(run.Child("rPr")).WithRevision(revision);

            foreach (LegacyDocElement child in run.Elements()) {
                switch (child.WName) {
                    case "rPr":
                        break;
                    case "lastRenderedPageBreak":
                        break;
                    case "t":
                        AppendFormattedNoteText(builder, runs, child.Text, formatting, storyStart);
                        break;
                    case "delText":
                        AppendFormattedNoteText(builder, runs, child.Text, formatting, storyStart);
                        break;
                    case "tab":
                        AppendFormattedNoteText(builder, runs, "\t", formatting, storyStart);
                        break;
                    case "cr":
                        AppendFormattedNoteText(builder, runs, LegacyDocSpecialCharacters.TextWrappingBreak.ToString(), formatting, storyStart);
                        break;
                    case "noBreakHyphen":
                        AppendFormattedNoteText(builder, runs, LegacyDocSpecialCharacters.NoBreakHyphen.ToString(), formatting, storyStart);
                        break;
                    case "softHyphen":
                        AppendFormattedNoteText(builder, runs, LegacyDocSpecialCharacters.SoftHyphen.ToString(), formatting, storyStart);
                        break;
                    case "br":
                        AppendSimpleEndnoteBreak(builder, runs, child, id, formatting, storyStart);
                        break;
                    case "drawing":
                        int picturePosition = storyStart + builder.Length;
                        int pictureDataOffset = pictures.AddInlinePicture(child, relationshipOwner);
                        builder.Append('\u0001');
                        runs.Add(new LegacyDocWritableRun(
                            picturePosition,
                            1,
                            LegacyDocWritableFormatting.SpecialCharacter.WithRevision(revision),
                            pictureDataOffset));
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple endnote id '{id}' only with text, embedded inline pictures, tabs, carriage returns, soft/no-break hyphens, and text-wrapping/page/column breaks. Unsupported endnote run element: {child.LocalName}.");
                }
            }
        }

        private static bool IsEndnoteReferenceMarkRun(LegacyDocElement run) {
            bool hasReferenceMark = false;
            foreach (LegacyDocElement child in run.Elements()) {
                switch (child.WName) {
                    case "rPr":
                    case "lastRenderedPageBreak":
                        break;
                    case "endnoteRef":
                        hasReferenceMark = true;
                        break;
                    default:
                        return false;
                }
            }

            return hasReferenceMark;
        }

        private static void AppendSimpleEndnoteBreak(StringBuilder builder, List<LegacyDocWritableRun> runs, LegacyDocElement breakNode, long id, LegacyDocWritableFormatting formatting, int storyStart) {
            string? breakType = breakNode.Attr("type");
            if (breakType == null || breakType == "textWrapping") {
                AppendFormattedNoteText(builder, runs, LegacyDocSpecialCharacters.TextWrappingBreak.ToString(), formatting, storyStart);
                return;
            }

            if (breakType == "page") {
                AppendFormattedNoteText(builder, runs, LegacyDocSpecialCharacters.PageBreak.ToString(), formatting, storyStart);
                return;
            }

            if (breakType == "column") {
                AppendFormattedNoteText(builder, runs, LegacyDocSpecialCharacters.ColumnBreak.ToString(), formatting, storyStart);
                return;
            }

            throw new NotSupportedException($"Native DOC saving supports simple endnote id '{id}' only with text-wrapping, page, and column breaks.");
        }

        private sealed class LegacyDocWritableEndnotes {
            internal static readonly LegacyDocWritableEndnotes Empty = new LegacyDocWritableEndnotes(new Dictionary<long, LegacyDocWritableNoteStory>());

            private readonly Dictionary<long, LegacyDocWritableNoteStory> _storiesById;
            private readonly List<LegacyDocWritableEndnoteReference> _references = new List<LegacyDocWritableEndnoteReference>();
            private readonly HashSet<long> _referencedIds = new HashSet<long>();

            internal LegacyDocWritableEndnotes(Dictionary<long, LegacyDocWritableNoteStory> storiesById) {
                _storiesById = storiesById;
            }

            internal bool HasReferences => _references.Count > 0;

            internal void AddReference(long id, int characterPosition) {
                if (!_storiesById.ContainsKey(id)) {
                    throw new NotSupportedException($"Native DOC saving cannot write endnote reference id '{id}' because the endnote body is missing.");
                }

                if (!_referencedIds.Add(id)) {
                    throw new NotSupportedException($"Native DOC saving currently supports one reference per endnote id. Duplicate reference id '{id}' was found.");
                }

                _references.Add(new LegacyDocWritableEndnoteReference(id, characterPosition));
            }

            internal void ThrowIfUnreferencedEndnotesRemain() {
                foreach (long id in _storiesById.Keys.OrderBy(item => item)) {
                    if (!_referencedIds.Contains(id)) {
                        throw new NotSupportedException($"Native DOC saving cannot write unreferenced endnote id '{id}'.");
                    }
                }
            }

            internal LegacyDocWritableEndnoteStories CreateStories(
                int bodyCharacterCount,
                int footnoteCharacterCount,
                int headerFooterCharacterCount,
                int commentCharacterCount,
                int terminalCharacterPadding) {
                if (_references.Count == 0) {
                    return LegacyDocWritableEndnoteStories.Empty;
                }

                var text = new StringBuilder();
                var runs = new List<LegacyDocWritableRun>();
                var paragraphs = new List<LegacyDocWritableParagraph>();
                var bookmarks = new LegacyDocWritableBookmarksBuilder();
                var textPositions = new List<int>(_references.Count + 2);
                var markerPositions = new List<int>(_references.Count);
                for (int index = 0; index < _references.Count; index++) {
                    LegacyDocWritableEndnoteReference reference = _references[index];
                    LegacyDocWritableNoteStory story = _storiesById[reference.Id];
                    textPositions.Add(text.Length);
                    markerPositions.Add(text.Length);
                    foreach (LegacyDocWritableRun run in story.FormattedRuns) {
                        runs.Add(new LegacyDocWritableRun(text.Length + run.StartCharacter, run.Length, run.Formatting, run.PictureDataOffset, run.FieldLocked));
                    }

                    foreach (LegacyDocWritableParagraph paragraph in story.FormattedParagraphs) {
                        paragraphs.Add(new LegacyDocWritableParagraph(text.Length + paragraph.StartCharacter, paragraph.Length, paragraph.Formatting));
                    }

                    bookmarks.AddRange(story.Bookmarks, text.Length);
                    text.Append(story.Text);
                }

                textPositions.Add(text.Length - 1);
                textPositions.Add(text.Length + 2);
                return new LegacyDocWritableEndnoteStories(
                    text.ToString(),
                    CreateEndnoteReferencePlc(
                        _references,
                        bodyCharacterCount
                            + footnoteCharacterCount
                            + headerFooterCharacterCount
                            + commentCharacterCount
                            + text.Length
                            + terminalCharacterPadding),
                    CreateEndnoteTextPlc(textPositions),
                    markerPositions,
                    runs,
                    paragraphs,
                    bookmarks.Create());
            }

            private static byte[] CreateEndnoteReferencePlc(IReadOnlyList<LegacyDocWritableEndnoteReference> references, int terminalCharacterPosition) {
                byte[] plc = new byte[(references.Count + 1) * 4 + references.Count * 2];
                for (int index = 0; index < references.Count; index++) {
                    WriteInt32(plc, index * 4, references[index].CharacterPosition);
                }

                WriteInt32(plc, references.Count * 4, terminalCharacterPosition);
                for (int index = 0; index < references.Count; index++) {
                    plc[((references.Count + 1) * 4) + (index * 2)] = 1;
                }

                return plc;
            }

            private static byte[] CreateEndnoteTextPlc(IReadOnlyList<int> textPositions) {
                byte[] plc = new byte[textPositions.Count * 4];
                for (int index = 0; index < textPositions.Count; index++) {
                    WriteInt32(plc, index * 4, textPositions[index]);
                }

                return plc;
            }
        }

        private readonly struct LegacyDocWritableEndnoteReference {
            internal LegacyDocWritableEndnoteReference(long id, int characterPosition) {
                Id = id;
                CharacterPosition = characterPosition;
            }

            internal long Id { get; }

            internal int CharacterPosition { get; }
        }

        private readonly struct LegacyDocWritableEndnoteStories {
            internal static readonly LegacyDocWritableEndnoteStories Empty = new LegacyDocWritableEndnoteStories(string.Empty, Array.Empty<byte>(), Array.Empty<byte>(), Array.Empty<int>(), Array.Empty<LegacyDocWritableRun>(), Array.Empty<LegacyDocWritableParagraph>(), LegacyDocWritableBookmarks.Empty);

            internal LegacyDocWritableEndnoteStories(string text, byte[] plcfendRef, byte[] plcfendTxt, IReadOnlyList<int> markerPositions, IReadOnlyList<LegacyDocWritableRun> formattedRuns, IReadOnlyList<LegacyDocWritableParagraph> formattedParagraphs, LegacyDocWritableBookmarks bookmarks) {
                Text = text;
                PlcfendRef = plcfendRef;
                PlcfendTxt = plcfendTxt;
                MarkerPositions = markerPositions;
                FormattedRuns = formattedRuns;
                FormattedParagraphs = formattedParagraphs;
                Bookmarks = bookmarks;
            }

            internal string Text { get; }

            internal byte[] PlcfendRef { get; }

            internal byte[] PlcfendTxt { get; }

            internal IReadOnlyList<int> MarkerPositions { get; }

            internal IReadOnlyList<LegacyDocWritableRun> FormattedRuns { get; }

            internal IReadOnlyList<LegacyDocWritableParagraph> FormattedParagraphs { get; }

            internal LegacyDocWritableBookmarks Bookmarks { get; }
        }
    }
}
