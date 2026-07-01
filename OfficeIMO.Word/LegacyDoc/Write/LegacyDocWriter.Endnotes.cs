using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static LegacyDocWritableEndnotes ReadSupportedEndnotes(MainDocumentPart mainPart) {
            Endnotes? endnotes = mainPart.EndnotesPart?.Endnotes;
            if (endnotes == null) {
                return LegacyDocWritableEndnotes.Empty;
            }

            var stories = new Dictionary<long, LegacyDocWritableNoteStory>();
            foreach (Endnote endnote in endnotes.Elements<Endnote>()) {
                if (!IsUserEndnote(endnote)) {
                    continue;
                }

                long? id = endnote.Id?.Value;
                if (id == null || id.Value <= 0) {
                    throw new NotSupportedException("Native DOC saving supports endnotes only when every user endnote has a positive identifier.");
                }

                if (stories.ContainsKey(id.Value)) {
                    throw new NotSupportedException($"Native DOC saving cannot write duplicate endnote id '{id.Value}'.");
                }

                stories.Add(id.Value, ReadSimpleEndnoteStory(endnote, id.Value, mainPart.EndnotesPart!));
            }

            return stories.Count == 0
                ? LegacyDocWritableEndnotes.Empty
                : new LegacyDocWritableEndnotes(stories);
        }

        private static LegacyDocWritableNoteStory ReadSimpleEndnoteStory(Endnote endnote, long id, EndnotesPart relationshipOwner) {
            var builder = new StringBuilder();
            var runs = new List<LegacyDocWritableRun>();
            var formattedParagraphs = new List<LegacyDocWritableParagraph>();
            var bookmarks = new LegacyDocWritableBookmarksBuilder();
            builder.Append(LegacyDocFootnoteReader.FootnoteReferenceCharacter);
            builder.Append(' ');
            bool hasBodyText = false;
            bool isFirstParagraph = true;
            foreach (OpenXmlElement child in endnote.ChildElements) {
                AppendSimpleEndnoteStoryChild(
                    child,
                    id,
                    relationshipOwner,
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
            OpenXmlElement child,
            long id,
            EndnotesPart relationshipOwner,
            StringBuilder builder,
            List<LegacyDocWritableRun> runs,
            List<LegacyDocWritableParagraph> formattedParagraphs,
            LegacyDocWritableBookmarksBuilder bookmarks,
            ref bool hasBodyText,
            ref bool isFirstParagraph,
            string containerDescription) {
            switch (child) {
                case Paragraph paragraph:
                    int paragraphStart = isFirstParagraph ? 0 : builder.Length;
                    LegacyDocWritableParagraphFormatting paragraphFormatting = ReadSimpleEndnoteParagraph(paragraph, id, runs, bookmarks, builder.Length, isFirstParagraph, relationshipOwner, out string paragraphText);
                    if (!string.IsNullOrEmpty(paragraphText)) {
                        hasBodyText = true;
                    }

                    builder.Append(paragraphText);
                    builder.Append('\r');
                    if (paragraphFormatting.HasFormatting) {
                        formattedParagraphs.Add(new LegacyDocWritableParagraph(paragraphStart, builder.Length - paragraphStart, paragraphFormatting));
                    }

                    isFirstParagraph = false;
                    break;
                case SdtBlock sdtBlock:
                    AppendSimpleEndnoteContentControl(
                        sdtBlock,
                        id,
                        relationshipOwner,
                        builder,
                        runs,
                        formattedParagraphs,
                        bookmarks,
                        ref hasBodyText,
                        ref isFirstParagraph);
                    break;
                case BookmarkStart bookmarkStart:
                    bookmarks.AddStart(bookmarkStart, builder.Length);
                    break;
                case BookmarkEnd bookmarkEnd:
                    bookmarks.AddEnd(bookmarkEnd, builder.Length);
                    break;
                default:
                    throw new NotSupportedException($"Native DOC saving supports simple endnote paragraphs, content controls, and bookmarks only. Unsupported {containerDescription} element: {child.LocalName}.");
            }
        }

        private static void AppendSimpleEndnoteContentControl(
            SdtBlock sdtBlock,
            long id,
            EndnotesPart relationshipOwner,
            StringBuilder builder,
            List<LegacyDocWritableRun> runs,
            List<LegacyDocWritableParagraph> formattedParagraphs,
            LegacyDocWritableBookmarksBuilder bookmarks,
            ref bool hasBodyText,
            ref bool isFirstParagraph) {
            SdtContentBlock? contentBlock = sdtBlock.SdtContentBlock;
            if (contentBlock == null) {
                throw new NotSupportedException($"Native DOC saving supports endnote id '{id}' content controls only when they contain simple paragraphs and bookmarks.");
            }

            foreach (OpenXmlElement child in contentBlock.ChildElements) {
                AppendSimpleEndnoteStoryChild(
                    child,
                    id,
                    relationshipOwner,
                    builder,
                    runs,
                    formattedParagraphs,
                    bookmarks,
                    ref hasBodyText,
                    ref isFirstParagraph,
                    "endnote content control");
            }
        }

        private static LegacyDocWritableParagraphFormatting ReadSimpleEndnoteParagraph(Paragraph paragraph, long id, List<LegacyDocWritableRun> runs, LegacyDocWritableBookmarksBuilder bookmarks, int storyStart, bool isFirstParagraph, EndnotesPart relationshipOwner, out string paragraphText) {
            var builder = new StringBuilder();
            LegacyDocWritableParagraphFormatting paragraphFormatting = ReadSupportedNoteParagraphFormatting(paragraph.ParagraphProperties, id, "endnote", EndnoteParagraphStyleIndexes);
            if (isFirstParagraph && paragraphFormatting.HasFormatting && paragraphFormatting.StyleIndex == null) {
                paragraphFormatting = paragraphFormatting.WithStyleIndex(NoteTextParagraphStyleIndex);
            }

            OpenXmlElement[] children = paragraph.ChildElements.ToArray();
            for (int index = 0; index < children.Length; index++) {
                OpenXmlElement child = children[index];
                switch (child) {
                    case ParagraphProperties:
                        break;
                    case Run run:
                        if (IsComplexFieldBeginRun(run)) {
                            AppendSupportedNoteComplexPageNumberField(children, ref index, builder, runs, storyStart);
                        } else {
                            AppendSimpleEndnoteRun(builder, runs, run, id, storyStart);
                        }

                        break;
                    case Hyperlink hyperlink:
                        AppendSupportedNoteHyperlinkText(builder, runs, hyperlink, relationshipOwner, id, "endnote", storyStart);
                        break;
                    case SimpleField simpleField:
                        AppendSupportedNoteFieldFromSimpleField(builder, runs, simpleField, storyStart);
                        break;
                    case BookmarkStart bookmarkStart:
                        bookmarks.AddStart(bookmarkStart, storyStart + builder.Length);
                        break;
                    case BookmarkEnd bookmarkEnd:
                        bookmarks.AddEnd(bookmarkEnd, storyStart + builder.Length);
                        break;
                    default:
                        if (IsIgnorableParagraphMarkup(child)) {
                            break;
                        }

                        throw new NotSupportedException($"Native DOC saving supports simple endnote paragraphs only with text runs, PAGE and NUMPAGES simple fields, bookmarks, and simple hyperlinks. Unsupported endnote paragraph element: {child.LocalName}.");
                }
            }

            paragraphText = builder.ToString();
            return paragraphFormatting;
        }

        private static void AppendSimpleEndnoteRun(StringBuilder builder, List<LegacyDocWritableRun> runs, Run run, long id, int storyStart) {
            if (IsEndnoteReferenceMarkRun(run)) {
                return;
            }

            LegacyDocWritableFormatting formatting = ReadSupportedRunFormatting(run.RunProperties);

            foreach (OpenXmlElement child in run.ChildElements) {
                switch (child) {
                    case RunProperties:
                        break;
                    case LastRenderedPageBreak:
                        break;
                    case Text text:
                        AppendFormattedNoteText(builder, runs, text.Text, formatting, storyStart);
                        break;
                    case TabChar:
                        AppendFormattedNoteText(builder, runs, "\t", formatting, storyStart);
                        break;
                    case CarriageReturn:
                        AppendFormattedNoteText(builder, runs, LegacyDocSpecialCharacters.TextWrappingBreak.ToString(), formatting, storyStart);
                        break;
                    case NoBreakHyphen:
                        AppendFormattedNoteText(builder, runs, LegacyDocSpecialCharacters.NoBreakHyphen.ToString(), formatting, storyStart);
                        break;
                    case SoftHyphen:
                        AppendFormattedNoteText(builder, runs, LegacyDocSpecialCharacters.SoftHyphen.ToString(), formatting, storyStart);
                        break;
                    case Break breakNode:
                        AppendSimpleEndnoteBreak(builder, runs, breakNode, id, formatting, storyStart);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple endnote id '{id}' only with text, tabs, carriage returns, soft/no-break hyphens, and text-wrapping/page/column breaks. Unsupported endnote run element: {child.LocalName}.");
                }
            }
        }

        private static bool IsEndnoteReferenceMarkRun(Run run) {
            bool hasReferenceMark = false;
            foreach (OpenXmlElement child in run.ChildElements) {
                switch (child) {
                    case RunProperties:
                    case LastRenderedPageBreak:
                        break;
                    case EndnoteReferenceMark:
                        hasReferenceMark = true;
                        break;
                    default:
                        return false;
                }
            }

            return hasReferenceMark;
        }

        private static void AppendSimpleEndnoteBreak(StringBuilder builder, List<LegacyDocWritableRun> runs, Break breakNode, long id, LegacyDocWritableFormatting formatting, int storyStart) {
            BreakValues? breakType = breakNode.Type?.Value;
            if (breakType == null || breakType == BreakValues.TextWrapping) {
                AppendFormattedNoteText(builder, runs, LegacyDocSpecialCharacters.TextWrappingBreak.ToString(), formatting, storyStart);
                return;
            }

            if (breakType == BreakValues.Page) {
                AppendFormattedNoteText(builder, runs, LegacyDocSpecialCharacters.PageBreak.ToString(), formatting, storyStart);
                return;
            }

            if (breakType == BreakValues.Column) {
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

            internal LegacyDocWritableEndnoteStories CreateStories(int bodyCharacterCount, int footnoteCharacterCount, int headerFooterCharacterCount, int terminalCharacterPadding) {
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
                        runs.Add(new LegacyDocWritableRun(text.Length + run.StartCharacter, run.Length, run.Formatting));
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
                    CreateEndnoteReferencePlc(_references, bodyCharacterCount + footnoteCharacterCount + headerFooterCharacterCount + text.Length + terminalCharacterPadding),
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
