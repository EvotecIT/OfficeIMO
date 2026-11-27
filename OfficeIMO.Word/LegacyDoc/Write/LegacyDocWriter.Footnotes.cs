using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static LegacyDocWritableFootnotes ReadSupportedFootnotes(LegacyDocSourcePackage source, LegacyDocWritablePictures pictures, IReadOnlyDictionary<string, ushort> styleIndexes) {
            LegacyDocElement footnotes = source.Footnotes;
            if (footnotes.IsNone) {
                return LegacyDocWritableFootnotes.Empty;
            }

            IReadOnlyDictionary<string, ushort> noteStyleIndexes = CreateNoteParagraphStyleIndexes(styleIndexes, "FootnoteText");
            LegacyDocPart relationshipOwner = source.Part(source.FootnotesPart!);
            var stories = new Dictionary<long, LegacyDocWritableNoteStory>();
            foreach (LegacyDocElement footnote in footnotes.Elements("footnote")) {
                if (!IsUserFootnote(footnote)) {
                    continue;
                }

                long? id = LegacyDocValues.Int64(footnote.Attr("id"));
                if (id == null || id.Value <= 0) {
                    throw new NotSupportedException("Native DOC saving supports footnotes only when every user footnote has a positive identifier.");
                }

                if (stories.ContainsKey(id.Value)) {
                    throw new NotSupportedException($"Native DOC saving cannot write duplicate footnote id '{id.Value}'.");
                }

                stories.Add(id.Value, ReadSimpleFootnoteStory(footnote, id.Value, relationshipOwner, pictures, noteStyleIndexes));
            }

            return stories.Count == 0
                ? LegacyDocWritableFootnotes.Empty
                : new LegacyDocWritableFootnotes(stories);
        }

        private static bool IsUserFootnote(LegacyDocElement footnote) {
            return footnote.Attr("type") is not { } type || type == "normal";
        }

        private static LegacyDocWritableNoteStory ReadSimpleFootnoteStory(LegacyDocElement footnote, long id, LegacyDocPart relationshipOwner, LegacyDocWritablePictures pictures, IReadOnlyDictionary<string, ushort> noteStyleIndexes) {
            var builder = new StringBuilder();
            var runs = new List<LegacyDocWritableRun>();
            var formattedParagraphs = new List<LegacyDocWritableParagraph>();
            var bookmarks = new LegacyDocWritableBookmarksBuilder();
            builder.Append(LegacyDocFootnoteReader.FootnoteReferenceCharacter);
            if (!HasExplicitNoteReferenceSeparator(footnote)) builder.Append(' ');
            bool hasBodyText = false;
            bool isFirstParagraph = true;
            foreach (LegacyDocElement child in footnote.Elements()) {
                AppendSimpleFootnoteStoryChild(
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
                    "footnote");
            }

            if (!hasBodyText) {
                throw new NotSupportedException($"Native DOC saving cannot write empty footnote id '{id}'.");
            }

            builder.Append('\r');
            return new LegacyDocWritableNoteStory(builder.ToString(), runs, formattedParagraphs, bookmarks.Create());
        }

        private static void AppendSimpleFootnoteStoryChild(
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
                    LegacyDocWritableParagraphFormatting paragraphFormatting = ReadSimpleFootnoteParagraph(child, id, runs, bookmarks, builder.Length, isFirstParagraph, relationshipOwner, pictures, noteStyleIndexes, out string paragraphText);
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
                    AppendSimpleFootnoteContentControl(
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
                    throw new NotSupportedException($"Native DOC saving supports simple footnote paragraphs, content controls, and bookmarks only. Unsupported {containerDescription} element: {child.LocalName}.");
            }
        }

        private static void AppendSimpleFootnoteContentControl(
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
                throw new NotSupportedException($"Native DOC saving supports footnote id '{id}' content controls only when they contain simple paragraphs and bookmarks.");
            }

            foreach (LegacyDocElement child in contentBlock.Elements()) {
                AppendSimpleFootnoteStoryChild(
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
                    "footnote content control");
            }
        }

        private static LegacyDocWritableParagraphFormatting ReadSimpleFootnoteParagraph(LegacyDocElement paragraph, long id, List<LegacyDocWritableRun> runs, LegacyDocWritableBookmarksBuilder bookmarks, int storyStart, bool isFirstParagraph, LegacyDocPart relationshipOwner, LegacyDocWritablePictures pictures, IReadOnlyDictionary<string, ushort> noteStyleIndexes, out string paragraphText) {
            var builder = new StringBuilder();
            LegacyDocWritableParagraphFormatting paragraphFormatting = ReadSupportedNoteParagraphFormatting(paragraph, id, "footnote", noteStyleIndexes);
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
                            AppendSimpleFootnoteRun(builder, runs, child, id, storyStart, relationshipOwner, pictures);
                        }

                        break;
                    case "ins":
                        AppendSimpleFootnoteRevision(builder, runs, child, LegacyDocRevisionKind.Inserted, id, storyStart, relationshipOwner, pictures);
                        break;
                    case "del":
                        AppendSimpleFootnoteRevision(builder, runs, child, LegacyDocRevisionKind.Deleted, id, storyStart, relationshipOwner, pictures);
                        break;
                    case "hyperlink":
                        AppendSupportedNoteHyperlinkText(builder, runs, bookmarks, child, relationshipOwner, id, "footnote", storyStart);
                        break;
                    case "fldSimple":
                        AppendSupportedNoteFieldFromSimpleField(builder, runs, bookmarks, child, storyStart);
                        break;
                    case "sdt":
                        AppendSupportedFootnoteInlineContentControl(builder, runs, bookmarks, child, relationshipOwner, pictures, id, storyStart);
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

                        throw new NotSupportedException($"Native DOC saving supports simple footnote paragraphs only with text runs, {SupportedFieldNames} simple fields, bookmarks, inline content controls, and simple hyperlinks. Unsupported footnote paragraph element: {child.LocalName}.");
                }
            }

            paragraphText = builder.ToString();
            return paragraphFormatting;
        }

        private static LegacyDocWritableParagraphFormatting ReadSupportedNoteParagraphFormatting(
            LegacyDocElement paragraph,
            long id,
            string noteKind,
            IReadOnlyDictionary<string, ushort> noteStyleIndexes) {
            try {
                return ReadSupportedParagraphFormatting(MaterializeNoteLineSpacing(paragraph, noteKind), noteStyleIndexes);
            } catch (NotSupportedException exception) {
                throw new NotSupportedException($"Native DOC saving supports simple {noteKind} id '{id}' only with supported paragraph formatting. {exception.Message}", exception);
            }
        }

        private static void AppendSimpleFootnoteRevision(
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
                    AppendSimpleFootnoteRun(builder, runs, child, id, storyStart, relationshipOwner, pictures, revision);
                    continue;
                }

                throw new NotSupportedException($"Native DOC saving supports tracked insertions and deletions in footnote id '{id}' only when they contain text runs. Unsupported revision element: {child.LocalName}.");
            }
        }

        private static void AppendSimpleFootnoteRun(StringBuilder builder, List<LegacyDocWritableRun> runs, LegacyDocElement run, long id, int storyStart, LegacyDocPart relationshipOwner, LegacyDocWritablePictures pictures, LegacyDocRevision revision = default) {
            if (IsFootnoteReferenceMarkRun(run)) {
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
                        AppendSimpleFootnoteBreak(builder, runs, child, id, formatting, storyStart);
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
                        throw new NotSupportedException($"Native DOC saving supports simple footnote id '{id}' only with text, embedded inline pictures, tabs, carriage returns, soft/no-break hyphens, and text-wrapping/page/column breaks. Unsupported footnote run element: {child.LocalName}.");
                }
            }
        }

        private static bool IsFootnoteReferenceMarkRun(LegacyDocElement run) {
            bool hasReferenceMark = false;
            foreach (LegacyDocElement child in run.Elements()) {
                switch (child.WName) {
                    case "rPr":
                    case "lastRenderedPageBreak":
                        break;
                    case "footnoteRef":
                        hasReferenceMark = true;
                        break;
                    default:
                        return false;
                }
            }

            return hasReferenceMark;
        }

        private static void AppendSimpleFootnoteBreak(StringBuilder builder, List<LegacyDocWritableRun> runs, LegacyDocElement breakNode, long id, LegacyDocWritableFormatting formatting, int storyStart) {
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

            throw new NotSupportedException($"Native DOC saving supports simple footnote id '{id}' only with text-wrapping, page, and column breaks.");
        }

        private static void AppendSupportedNoteHyperlinkText(
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            LegacyDocWritableBookmarksBuilder bookmarks,
            LegacyDocElement hyperlink,
            LegacyDocPart relationshipOwner,
            long id,
            string noteKind,
            int storyStart) {
            string instruction = CreateSupportedHyperlinkInstruction(hyperlink, relationshipOwner);
            AppendFormattedNoteText(text, runs, LegacyDocField.Begin.ToString(), LegacyDocWritableFormatting.SpecialCharacter, storyStart);
            AppendFormattedNoteText(text, runs, instruction, LegacyDocWritableFormatting.Plain, storyStart);
            AppendFormattedNoteText(text, runs, LegacyDocField.Separator.ToString(), LegacyDocWritableFormatting.SpecialCharacter, storyStart);

            int displayStart = text.Length;
            foreach (LegacyDocElement child in hyperlink.Elements()) {
                if (child.Is(OfficeXmlNamespaces.Math, "oMath") || child.Is(OfficeXmlNamespaces.Math, "oMathPara")) {
                    AppendMathEquationNoteField(text, runs, bookmarks, child, storyStart);
                    continue;
                }

                switch (child.WName) {
                    case "r":
                        EnsureSupportedHyperlinkRun(child);
                        AppendSupportedNoteHyperlinkRunText(text, runs, child, id, noteKind, storyStart);
                        break;
                    case "sdt":
                        AppendSupportedNoteHyperlinkInlineContentControlText(text, runs, bookmarks, child, id, noteKind, storyStart);
                        break;
                    case "fldSimple":
                        AppendSupportedNoteFieldFromSimpleField(
                            text,
                            runs,
                            bookmarks,
                            child,
                            storyStart,
                            allowHyperlinkRunStyle: true);
                        break;
                    case "bookmarkStart":
                        bookmarks.AddStart(child, storyStart + text.Length);
                        break;
                    case "bookmarkEnd":
                        bookmarks.AddEnd(child, storyStart + text.Length);
                        break;
                    default:
                        if (IsIgnorableParagraphMarkup(child)) {
                            break;
                        }

                        throw new NotSupportedException($"Native DOC saving supports simple {noteKind} hyperlinks only when they contain text runs, supported fields, equations, and inline content controls. Unsupported hyperlink element: {child.LocalName}.");
                }
            }

            if (text.Length == displayStart) {
                throw new NotSupportedException($"Native DOC saving supports {noteKind} hyperlinks only when they contain display text.");
            }

            AppendFormattedNoteText(text, runs, LegacyDocField.End.ToString(), LegacyDocWritableFormatting.SpecialCharacter, storyStart);
        }

        private static void AppendSupportedNoteHyperlinkRunText(StringBuilder text, List<LegacyDocWritableRun> runs, LegacyDocElement run, long id, string noteKind, int storyStart) {
            LegacyDocWritableFormatting formatting = ReadSupportedRunFormatting(run.Child("rPr"), allowHyperlinkRunStyle: true);

            foreach (LegacyDocElement child in run.Elements()) {
                switch (child.WName) {
                    case "rPr":
                        break;
                    case "lastRenderedPageBreak":
                        break;
                    case "t":
                        AppendFormattedNoteText(text, runs, child.Text, formatting, storyStart);
                        break;
                    case "tab":
                        AppendFormattedNoteText(text, runs, "\t", formatting, storyStart);
                        break;
                    case "cr":
                        AppendFormattedNoteText(text, runs, LegacyDocSpecialCharacters.TextWrappingBreak.ToString(), formatting, storyStart);
                        break;
                    case "noBreakHyphen":
                        AppendFormattedNoteText(text, runs, LegacyDocSpecialCharacters.NoBreakHyphen.ToString(), formatting, storyStart);
                        break;
                    case "softHyphen":
                        AppendFormattedNoteText(text, runs, LegacyDocSpecialCharacters.SoftHyphen.ToString(), formatting, storyStart);
                        break;
                    case "br":
                        AppendSupportedNoteHyperlinkBreak(text, runs, child, id, noteKind, storyStart, formatting);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple {noteKind} id '{id}' hyperlinks only with text, tabs, carriage returns, soft/no-break hyphens, and text-wrapping/page/column break display runs. Unsupported hyperlink run element: {child.LocalName}.");
                }
            }
        }

        private static void AppendSupportedNoteHyperlinkInlineContentControlText(
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            LegacyDocWritableBookmarksBuilder bookmarks,
            LegacyDocElement sdtRun,
            long id,
            string noteKind,
            int storyStart) {
            LegacyDocElement[] children = GetInlineContentControlChildren(sdtRun, $"{noteKind} id '{id}' hyperlink inline content control");
            foreach (LegacyDocElement child in children) {
                if (child.Is(OfficeXmlNamespaces.Math, "oMath") || child.Is(OfficeXmlNamespaces.Math, "oMathPara")) {
                    AppendMathEquationNoteField(text, runs, bookmarks, child, storyStart);
                    continue;
                }

                switch (child.WName) {
                    case "r":
                        EnsureSupportedHyperlinkRun(child);
                        AppendSupportedNoteHyperlinkRunText(text, runs, child, id, noteKind, storyStart);
                        break;
                    case "sdt":
                        AppendSupportedNoteHyperlinkInlineContentControlText(text, runs, bookmarks, child, id, noteKind, storyStart);
                        break;
                    case "fldSimple":
                        AppendSupportedNoteFieldFromSimpleField(
                            text,
                            runs,
                            bookmarks,
                            child,
                            storyStart,
                            allowHyperlinkRunStyle: true);
                        break;
                    case "bookmarkStart":
                        bookmarks.AddStart(child, storyStart + text.Length);
                        break;
                    case "bookmarkEnd":
                        bookmarks.AddEnd(child, storyStart + text.Length);
                        break;
                    default:
                        if (IsIgnorableParagraphMarkup(child)) {
                            break;
                        }

                        throw new NotSupportedException($"Native DOC saving supports {noteKind} id '{id}' hyperlink inline content controls only when they contain supported text runs, supported fields, equations, and nested inline content controls. Unsupported {noteKind} hyperlink inline content-control element: {child.LocalName}.");
                }
            }
        }

        private static void AppendSupportedNoteHyperlinkBreak(
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            LegacyDocElement breakNode,
            long id,
            string noteKind,
            int storyStart,
            LegacyDocWritableFormatting formatting) {
            string? breakType = breakNode.Attr("type");
            if (breakType == null || breakType == "textWrapping") {
                AppendFormattedNoteText(text, runs, LegacyDocSpecialCharacters.TextWrappingBreak.ToString(), formatting, storyStart);
                return;
            }

            if (breakType == "page") {
                AppendFormattedNoteText(text, runs, LegacyDocSpecialCharacters.PageBreak.ToString(), formatting, storyStart);
                return;
            }

            if (breakType == "column") {
                AppendFormattedNoteText(text, runs, LegacyDocSpecialCharacters.ColumnBreak.ToString(), formatting, storyStart);
                return;
            }

            throw new NotSupportedException($"Native DOC saving supports simple {noteKind} id '{id}' hyperlinks only with text-wrapping, page, and column breaks.");
        }

        private sealed class LegacyDocWritableFootnotes {
            internal static readonly LegacyDocWritableFootnotes Empty = new LegacyDocWritableFootnotes(new Dictionary<long, LegacyDocWritableNoteStory>());

            private readonly Dictionary<long, LegacyDocWritableNoteStory> _storiesById;
            private readonly List<LegacyDocWritableFootnoteReference> _references = new List<LegacyDocWritableFootnoteReference>();
            private readonly HashSet<long> _referencedIds = new HashSet<long>();

            internal LegacyDocWritableFootnotes(Dictionary<long, LegacyDocWritableNoteStory> storiesById) {
                _storiesById = storiesById;
            }

            internal bool HasReferences => _references.Count > 0;

            internal void AddReference(long id, int characterPosition) {
                if (!_storiesById.ContainsKey(id)) {
                    throw new NotSupportedException($"Native DOC saving cannot write footnote reference id '{id}' because the footnote body is missing.");
                }

                if (!_referencedIds.Add(id)) {
                    throw new NotSupportedException($"Native DOC saving currently supports one reference per footnote id. Duplicate reference id '{id}' was found.");
                }

                _references.Add(new LegacyDocWritableFootnoteReference(id, characterPosition));
            }

            internal void ThrowIfUnreferencedFootnotesRemain() {
                foreach (long id in _storiesById.Keys.OrderBy(item => item)) {
                    if (!_referencedIds.Contains(id)) {
                        throw new NotSupportedException($"Native DOC saving cannot write unreferenced footnote id '{id}'.");
                    }
                }
            }

            internal LegacyDocWritableFootnoteStories CreateStories(int bodyCharacterCount, int headerFooterCharacterCount, int terminalCharacterPadding) {
                if (_references.Count == 0) {
                    return LegacyDocWritableFootnoteStories.Empty;
                }

                var text = new StringBuilder();
                var runs = new List<LegacyDocWritableRun>();
                var paragraphs = new List<LegacyDocWritableParagraph>();
                var bookmarks = new LegacyDocWritableBookmarksBuilder();
                var textPositions = new List<int>(_references.Count + 2);
                var markerPositions = new List<int>(_references.Count);
                for (int index = 0; index < _references.Count; index++) {
                    LegacyDocWritableFootnoteReference reference = _references[index];
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
                return new LegacyDocWritableFootnoteStories(
                    text.ToString(),
                    CreateFootnoteReferencePlc(_references, bodyCharacterCount + text.Length + headerFooterCharacterCount + terminalCharacterPadding),
                    CreateFootnoteTextPlc(textPositions),
                    markerPositions,
                    runs,
                    paragraphs,
                    bookmarks.Create());
            }

            private static byte[] CreateFootnoteReferencePlc(IReadOnlyList<LegacyDocWritableFootnoteReference> references, int terminalCharacterPosition) {
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

            private static byte[] CreateFootnoteTextPlc(IReadOnlyList<int> textPositions) {
                byte[] plc = new byte[textPositions.Count * 4];
                for (int index = 0; index < textPositions.Count; index++) {
                    WriteInt32(plc, index * 4, textPositions[index]);
                }

                return plc;
            }
        }

        private readonly struct LegacyDocWritableFootnoteReference {
            internal LegacyDocWritableFootnoteReference(long id, int characterPosition) {
                Id = id;
                CharacterPosition = characterPosition;
            }

            internal long Id { get; }

            internal int CharacterPosition { get; }
        }

        private readonly struct LegacyDocWritableFootnoteStories {
            internal static readonly LegacyDocWritableFootnoteStories Empty = new LegacyDocWritableFootnoteStories(string.Empty, Array.Empty<byte>(), Array.Empty<byte>(), Array.Empty<int>(), Array.Empty<LegacyDocWritableRun>(), Array.Empty<LegacyDocWritableParagraph>(), LegacyDocWritableBookmarks.Empty);

            internal LegacyDocWritableFootnoteStories(string text, byte[] plcffndRef, byte[] plcffndTxt, IReadOnlyList<int> markerPositions, IReadOnlyList<LegacyDocWritableRun> formattedRuns, IReadOnlyList<LegacyDocWritableParagraph> formattedParagraphs, LegacyDocWritableBookmarks bookmarks) {
                Text = text;
                PlcffndRef = plcffndRef;
                PlcffndTxt = plcffndTxt;
                MarkerPositions = markerPositions;
                FormattedRuns = formattedRuns;
                FormattedParagraphs = formattedParagraphs;
                Bookmarks = bookmarks;
            }

            internal string Text { get; }

            internal byte[] PlcffndRef { get; }

            internal byte[] PlcffndTxt { get; }

            internal IReadOnlyList<int> MarkerPositions { get; }

            internal IReadOnlyList<LegacyDocWritableRun> FormattedRuns { get; }

            internal IReadOnlyList<LegacyDocWritableParagraph> FormattedParagraphs { get; }

            internal LegacyDocWritableBookmarks Bookmarks { get; }
        }

        private readonly struct LegacyDocWritableNoteStory {
            internal LegacyDocWritableNoteStory(string text, IReadOnlyList<LegacyDocWritableRun> formattedRuns, IReadOnlyList<LegacyDocWritableParagraph> formattedParagraphs, LegacyDocWritableBookmarks bookmarks) {
                Text = text;
                FormattedRuns = formattedRuns;
                FormattedParagraphs = formattedParagraphs;
                Bookmarks = bookmarks;
            }

            internal string Text { get; }

            internal IReadOnlyList<LegacyDocWritableRun> FormattedRuns { get; }

            internal IReadOnlyList<LegacyDocWritableParagraph> FormattedParagraphs { get; }

            internal LegacyDocWritableBookmarks Bookmarks { get; }
        }

        private static void AppendFormattedNoteText(
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            string? value,
            LegacyDocWritableFormatting formatting,
            int storyStart) {
            if (string.IsNullOrEmpty(value)) {
                return;
            }

            int start = storyStart + text.Length;
            text.Append(value);
            if (!formatting.HasFormatting) {
                return;
            }

            int length = value!.Length;
            if (runs.Count > 0) {
                LegacyDocWritableRun previous = runs[runs.Count - 1];
                if (!previous.FieldLocked && previous.EndCharacter == start && previous.Formatting.HasSameEncoding(formatting)) {
                    runs[runs.Count - 1] = previous.Extend(length);
                    return;
                }
            }

            runs.Add(new LegacyDocWritableRun(start, length, formatting));
        }
    }
}
