using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static readonly IReadOnlyDictionary<string, ushort> FootnoteParagraphStyleIndexes = new Dictionary<string, ushort>(StringComparer.OrdinalIgnoreCase) {
            ["FootnoteText"] = NoteTextParagraphStyleIndex
        };

        private static readonly IReadOnlyDictionary<string, ushort> EndnoteParagraphStyleIndexes = new Dictionary<string, ushort>(StringComparer.OrdinalIgnoreCase) {
            ["EndnoteText"] = NoteTextParagraphStyleIndex
        };

        private static LegacyDocWritableFootnotes ReadSupportedFootnotes(MainDocumentPart mainPart) {
            Footnotes? footnotes = mainPart.FootnotesPart?.Footnotes;
            if (footnotes == null) {
                return LegacyDocWritableFootnotes.Empty;
            }

            var stories = new Dictionary<long, LegacyDocWritableNoteStory>();
            foreach (Footnote footnote in footnotes.Elements<Footnote>()) {
                if (!IsUserFootnote(footnote)) {
                    continue;
                }

                long? id = footnote.Id?.Value;
                if (id == null || id.Value <= 0) {
                    throw new NotSupportedException("Native DOC saving supports footnotes only when every user footnote has a positive identifier.");
                }

                if (stories.ContainsKey(id.Value)) {
                    throw new NotSupportedException($"Native DOC saving cannot write duplicate footnote id '{id.Value}'.");
                }

                stories.Add(id.Value, ReadSimpleFootnoteStory(footnote, id.Value, mainPart.FootnotesPart!));
            }

            return stories.Count == 0
                ? LegacyDocWritableFootnotes.Empty
                : new LegacyDocWritableFootnotes(stories);
        }

        private static bool IsUserFootnote(Footnote footnote) {
            return footnote.Type == null || footnote.Type.Value == FootnoteEndnoteValues.Normal;
        }

        private static LegacyDocWritableNoteStory ReadSimpleFootnoteStory(Footnote footnote, long id, FootnotesPart relationshipOwner) {
            var builder = new StringBuilder();
            var runs = new List<LegacyDocWritableRun>();
            var formattedParagraphs = new List<LegacyDocWritableParagraph>();
            var bookmarks = new LegacyDocWritableBookmarksBuilder();
            builder.Append(LegacyDocFootnoteReader.FootnoteReferenceCharacter);
            builder.Append(' ');
            bool hasBodyText = false;
            bool isFirstParagraph = true;
            foreach (OpenXmlElement child in footnote.ChildElements) {
                AppendSimpleFootnoteStoryChild(
                    child,
                    id,
                    relationshipOwner,
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
            OpenXmlElement child,
            long id,
            FootnotesPart relationshipOwner,
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
                    LegacyDocWritableParagraphFormatting paragraphFormatting = ReadSimpleFootnoteParagraph(paragraph, id, runs, bookmarks, builder.Length, isFirstParagraph, relationshipOwner, out string paragraphText);
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
                    AppendSimpleFootnoteContentControl(
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
                    throw new NotSupportedException($"Native DOC saving supports simple footnote paragraphs, content controls, and bookmarks only. Unsupported {containerDescription} element: {child.LocalName}.");
            }
        }

        private static void AppendSimpleFootnoteContentControl(
            SdtBlock sdtBlock,
            long id,
            FootnotesPart relationshipOwner,
            StringBuilder builder,
            List<LegacyDocWritableRun> runs,
            List<LegacyDocWritableParagraph> formattedParagraphs,
            LegacyDocWritableBookmarksBuilder bookmarks,
            ref bool hasBodyText,
            ref bool isFirstParagraph) {
            SdtContentBlock? contentBlock = sdtBlock.SdtContentBlock;
            if (contentBlock == null) {
                throw new NotSupportedException($"Native DOC saving supports footnote id '{id}' content controls only when they contain simple paragraphs and bookmarks.");
            }

            foreach (OpenXmlElement child in contentBlock.ChildElements) {
                AppendSimpleFootnoteStoryChild(
                    child,
                    id,
                    relationshipOwner,
                    builder,
                    runs,
                    formattedParagraphs,
                    bookmarks,
                    ref hasBodyText,
                    ref isFirstParagraph,
                    "footnote content control");
            }
        }

        private static LegacyDocWritableParagraphFormatting ReadSimpleFootnoteParagraph(Paragraph paragraph, long id, List<LegacyDocWritableRun> runs, LegacyDocWritableBookmarksBuilder bookmarks, int storyStart, bool isFirstParagraph, FootnotesPart relationshipOwner, out string paragraphText) {
            var builder = new StringBuilder();
            LegacyDocWritableParagraphFormatting paragraphFormatting = ReadSupportedNoteParagraphFormatting(paragraph.ParagraphProperties, id, "footnote", FootnoteParagraphStyleIndexes);
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
                            AppendSupportedNoteComplexPageNumberField(children, ref index, builder, runs, bookmarks, storyStart);
                        } else {
                            AppendSimpleFootnoteRun(builder, runs, run, id, storyStart);
                        }

                        break;
                    case Hyperlink hyperlink:
                        AppendSupportedNoteHyperlinkText(builder, runs, bookmarks, hyperlink, relationshipOwner, id, "footnote", storyStart);
                        break;
                    case SimpleField simpleField:
                        AppendSupportedNoteFieldFromSimpleField(builder, runs, bookmarks, simpleField, storyStart);
                        break;
                    case SdtRun sdtRun:
                        AppendSupportedFootnoteInlineContentControl(builder, runs, bookmarks, sdtRun, relationshipOwner, id, storyStart);
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

                        throw new NotSupportedException($"Native DOC saving supports simple footnote paragraphs only with text runs, PAGE and NUMPAGES simple fields, bookmarks, inline content controls, and simple hyperlinks. Unsupported footnote paragraph element: {child.LocalName}.");
                }
            }

            paragraphText = builder.ToString();
            return paragraphFormatting;
        }

        private static LegacyDocWritableParagraphFormatting ReadSupportedNoteParagraphFormatting(
            ParagraphProperties? paragraphProperties,
            long id,
            string noteKind,
            IReadOnlyDictionary<string, ushort> noteStyleIndexes) {
            try {
                return ReadSupportedParagraphFormatting(paragraphProperties, noteStyleIndexes);
            } catch (NotSupportedException exception) {
                throw new NotSupportedException($"Native DOC saving supports simple {noteKind} id '{id}' only with supported paragraph formatting. {exception.Message}", exception);
            }
        }

        private static void AppendSimpleFootnoteRun(StringBuilder builder, List<LegacyDocWritableRun> runs, Run run, long id, int storyStart) {
            if (IsFootnoteReferenceMarkRun(run)) {
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
                        AppendSimpleFootnoteBreak(builder, runs, breakNode, id, formatting, storyStart);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple footnote id '{id}' only with text, tabs, carriage returns, soft/no-break hyphens, and text-wrapping/page/column breaks. Unsupported footnote run element: {child.LocalName}.");
                }
            }
        }

        private static bool IsFootnoteReferenceMarkRun(Run run) {
            bool hasReferenceMark = false;
            foreach (OpenXmlElement child in run.ChildElements) {
                switch (child) {
                    case RunProperties:
                    case LastRenderedPageBreak:
                        break;
                    case FootnoteReferenceMark:
                        hasReferenceMark = true;
                        break;
                    default:
                        return false;
                }
            }

            return hasReferenceMark;
        }

        private static void AppendSimpleFootnoteBreak(StringBuilder builder, List<LegacyDocWritableRun> runs, Break breakNode, long id, LegacyDocWritableFormatting formatting, int storyStart) {
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

            throw new NotSupportedException($"Native DOC saving supports simple footnote id '{id}' only with text-wrapping, page, and column breaks.");
        }

        private static void AppendSupportedNoteFieldFromSimpleField(StringBuilder text, List<LegacyDocWritableRun> runs, LegacyDocWritableBookmarksBuilder bookmarks, SimpleField field, int storyStart) {
            if (!TryReadSupportedFieldKind(field.Instruction?.Value, out LegacyDocFieldKind fieldKind)) {
                throw new NotSupportedException("Native DOC saving currently supports only PAGE and NUMPAGES simple fields in note paragraphs. Other field types are not supported yet.");
            }

            LegacyDocSimpleFieldResult result = ReadSimpleFieldResult(field);
            AppendSupportedNoteField(text, runs, bookmarks, fieldKind, result.Formatting, result.BookmarkMarkers, storyStart);
        }

        private static void AppendSupportedNoteComplexPageNumberField(
            IReadOnlyList<OpenXmlElement> paragraphChildren,
            ref int childIndex,
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            LegacyDocWritableBookmarksBuilder bookmarks,
            int storyStart) {
            var instruction = new StringBuilder();
            LegacyDocWritableFormatting? resultFormatting = null;
            var bookmarkMarkers = new List<LegacyDocSimpleFieldBookmarkMarker>();
            bool sawSeparator = false;
            int resultOffset = 0;
            int index = childIndex;
            for (; index < paragraphChildren.Count; index++) {
                OpenXmlElement fieldChild = paragraphChildren[index];
                if (fieldChild is BookmarkStart bookmarkStart && sawSeparator) {
                    bookmarkMarkers.Add(new LegacyDocSimpleFieldBookmarkMarker(bookmarkStart, null, resultOffset));
                    continue;
                }

                if (fieldChild is BookmarkEnd bookmarkEnd && sawSeparator) {
                    bookmarkMarkers.Add(new LegacyDocSimpleFieldBookmarkMarker(null, bookmarkEnd, resultOffset));
                    continue;
                }

                if (fieldChild is not Run run) {
                    if (IsIgnorableParagraphMarkup(fieldChild)) {
                        continue;
                    }

                    throw new NotSupportedException("Native DOC saving supports PAGE and NUMPAGES complex fields in note paragraphs only when the whole field is represented by adjacent runs.");
                }

                LegacyDocWritableFormatting runFormatting = ReadSupportedRunFormatting(run.RunProperties);
                foreach (OpenXmlElement child in run.ChildElements) {
                    switch (child) {
                        case RunProperties:
                        case LastRenderedPageBreak:
                            break;
                        case FieldCode fieldCode when !sawSeparator:
                            instruction.Append(fieldCode.Text);
                            break;
                        case Text when sawSeparator:
                            resultFormatting ??= runFormatting;
                            if (!resultFormatting.Value.Equals(runFormatting)) {
                                throw new NotSupportedException("Native DOC saving supports PAGE and NUMPAGES complex fields in note paragraphs only when their display runs use one formatting set.");
                            }

                            resultOffset = 1;
                            break;
                        case FieldChar fieldChar:
                            FieldCharValues? fieldCharType = fieldChar.FieldCharType?.Value;
                            if (fieldCharType == FieldCharValues.Begin) {
                                if (index != childIndex) {
                                    throw new NotSupportedException("Native DOC saving does not support nested complex fields in note paragraph PAGE and NUMPAGES field runs.");
                                }

                                break;
                            }

                            if (fieldCharType == FieldCharValues.Separate) {
                                sawSeparator = true;
                                break;
                            }

                            if (fieldCharType == FieldCharValues.End) {
                                if (!TryReadSupportedFieldKind(instruction.ToString(), out LegacyDocFieldKind fieldKind)) {
                                    throw new NotSupportedException("Native DOC saving currently supports only PAGE and NUMPAGES complex fields in note paragraphs. Other field types are not supported yet.");
                                }

                                AppendSupportedNoteField(text, runs, bookmarks, fieldKind, resultFormatting ?? LegacyDocWritableFormatting.Plain, bookmarkMarkers, storyStart);
                                childIndex = index;
                                return;
                            }

                            throw new NotSupportedException("Native DOC saving supports PAGE and NUMPAGES complex fields in note paragraphs only with begin, separate, and end field characters.");
                        default:
                            throw new NotSupportedException($"Native DOC saving supports PAGE and NUMPAGES complex fields in note paragraphs only with field code and display text runs. Unsupported field run element: {child.LocalName}.");
                    }
                }
            }

            throw new NotSupportedException("Native DOC saving cannot write an unterminated PAGE or NUMPAGES complex field in a note paragraph.");
        }

        private static void AppendSupportedNoteField(StringBuilder text, List<LegacyDocWritableRun> runs, LegacyDocFieldKind fieldKind, LegacyDocWritableFormatting formatting, int storyStart) {
            AppendFormattedNoteText(text, runs, LegacyDocField.Begin.ToString(), LegacyDocWritableFormatting.SpecialCharacter, storyStart);
            AppendFormattedNoteText(text, runs, GetSupportedFieldInstruction(fieldKind), LegacyDocWritableFormatting.Plain, storyStart);
            AppendFormattedNoteText(text, runs, LegacyDocField.Separator.ToString(), LegacyDocWritableFormatting.SpecialCharacter, storyStart);
            AppendFormattedNoteText(text, runs, "1", formatting, storyStart);
            AppendFormattedNoteText(text, runs, LegacyDocField.End.ToString(), LegacyDocWritableFormatting.SpecialCharacter, storyStart);
        }

        private static void AppendSupportedNoteField(
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            LegacyDocWritableBookmarksBuilder bookmarks,
            LegacyDocFieldKind fieldKind,
            LegacyDocWritableFormatting formatting,
            IReadOnlyList<LegacyDocSimpleFieldBookmarkMarker> bookmarkMarkers,
            int storyStart) {
            AppendFormattedNoteText(text, runs, LegacyDocField.Begin.ToString(), LegacyDocWritableFormatting.SpecialCharacter, storyStart);
            AppendFormattedNoteText(text, runs, GetSupportedFieldInstruction(fieldKind), LegacyDocWritableFormatting.Plain, storyStart);
            AppendFormattedNoteText(text, runs, LegacyDocField.Separator.ToString(), LegacyDocWritableFormatting.SpecialCharacter, storyStart);
            int resultStartCharacter = storyStart + text.Length;
            AddSimpleFieldBookmarkMarkers(bookmarks, bookmarkMarkers, resultStartCharacter, resultOffset: 0);
            AppendFormattedNoteText(text, runs, "1", formatting, storyStart);
            AddSimpleFieldBookmarkMarkers(bookmarks, bookmarkMarkers, resultStartCharacter, resultOffset: 1);
            AppendFormattedNoteText(text, runs, LegacyDocField.End.ToString(), LegacyDocWritableFormatting.SpecialCharacter, storyStart);
        }

        private static void AppendSupportedNoteHyperlinkText(
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            LegacyDocWritableBookmarksBuilder bookmarks,
            Hyperlink hyperlink,
            OpenXmlPartContainer relationshipOwner,
            long id,
            string noteKind,
            int storyStart) {
            string instruction = CreateSupportedHyperlinkInstruction(hyperlink, relationshipOwner);
            AppendFormattedNoteText(text, runs, LegacyDocField.Begin.ToString(), LegacyDocWritableFormatting.SpecialCharacter, storyStart);
            AppendFormattedNoteText(text, runs, instruction, LegacyDocWritableFormatting.Plain, storyStart);
            AppendFormattedNoteText(text, runs, LegacyDocField.Separator.ToString(), LegacyDocWritableFormatting.SpecialCharacter, storyStart);

            int displayStart = text.Length;
            foreach (OpenXmlElement child in hyperlink.ChildElements) {
                switch (child) {
                    case Run run:
                        EnsureSupportedHyperlinkRun(run);
                        AppendSupportedNoteHyperlinkRunText(text, runs, run, id, noteKind, storyStart);
                        break;
                    case SdtRun sdtRun:
                        AppendSupportedNoteHyperlinkInlineContentControlText(text, runs, bookmarks, sdtRun, id, noteKind, storyStart);
                        break;
                    case BookmarkStart bookmarkStart:
                        bookmarks.AddStart(bookmarkStart, storyStart + text.Length);
                        break;
                    case BookmarkEnd bookmarkEnd:
                        bookmarks.AddEnd(bookmarkEnd, storyStart + text.Length);
                        break;
                    default:
                        if (IsIgnorableParagraphMarkup(child)) {
                            break;
                        }

                        throw new NotSupportedException($"Native DOC saving supports simple {noteKind} hyperlinks only when they contain text runs and inline content controls. Unsupported hyperlink element: {child.LocalName}.");
                }
            }

            if (text.Length == displayStart) {
                throw new NotSupportedException($"Native DOC saving supports {noteKind} hyperlinks only when they contain display text.");
            }

            AppendFormattedNoteText(text, runs, LegacyDocField.End.ToString(), LegacyDocWritableFormatting.SpecialCharacter, storyStart);
        }

        private static void AppendSupportedNoteHyperlinkRunText(StringBuilder text, List<LegacyDocWritableRun> runs, Run run, long id, string noteKind, int storyStart) {
            LegacyDocWritableFormatting formatting = ReadSupportedRunFormatting(run.RunProperties, allowHyperlinkRunStyle: true);

            foreach (OpenXmlElement child in run.ChildElements) {
                switch (child) {
                    case RunProperties:
                        break;
                    case LastRenderedPageBreak:
                        break;
                    case Text textNode:
                        AppendFormattedNoteText(text, runs, textNode.Text, formatting, storyStart);
                        break;
                    case TabChar:
                        AppendFormattedNoteText(text, runs, "\t", formatting, storyStart);
                        break;
                    case CarriageReturn:
                        AppendFormattedNoteText(text, runs, LegacyDocSpecialCharacters.TextWrappingBreak.ToString(), formatting, storyStart);
                        break;
                    case NoBreakHyphen:
                        AppendFormattedNoteText(text, runs, LegacyDocSpecialCharacters.NoBreakHyphen.ToString(), formatting, storyStart);
                        break;
                    case SoftHyphen:
                        AppendFormattedNoteText(text, runs, LegacyDocSpecialCharacters.SoftHyphen.ToString(), formatting, storyStart);
                        break;
                    case Break breakNode:
                        AppendSupportedNoteHyperlinkBreak(text, runs, breakNode, id, noteKind, storyStart, formatting);
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
            SdtRun sdtRun,
            long id,
            string noteKind,
            int storyStart) {
            OpenXmlElement[] children = GetInlineContentControlChildren(sdtRun, $"{noteKind} id '{id}' hyperlink inline content control");
            foreach (OpenXmlElement child in children) {
                switch (child) {
                    case Run run:
                        EnsureSupportedHyperlinkRun(run);
                        AppendSupportedNoteHyperlinkRunText(text, runs, run, id, noteKind, storyStart);
                        break;
                    case SdtRun nestedSdtRun:
                        AppendSupportedNoteHyperlinkInlineContentControlText(text, runs, bookmarks, nestedSdtRun, id, noteKind, storyStart);
                        break;
                    case BookmarkStart bookmarkStart:
                        bookmarks.AddStart(bookmarkStart, storyStart + text.Length);
                        break;
                    case BookmarkEnd bookmarkEnd:
                        bookmarks.AddEnd(bookmarkEnd, storyStart + text.Length);
                        break;
                    default:
                        if (IsIgnorableParagraphMarkup(child)) {
                            break;
                        }

                        throw new NotSupportedException($"Native DOC saving supports {noteKind} id '{id}' hyperlink inline content controls only when they contain supported text runs and nested inline content controls. Unsupported {noteKind} hyperlink inline content-control element: {child.LocalName}.");
                }
            }
        }

        private static void AppendSupportedNoteHyperlinkBreak(
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            Break breakNode,
            long id,
            string noteKind,
            int storyStart,
            LegacyDocWritableFormatting formatting) {
            BreakValues? breakType = breakNode.Type?.Value;
            if (breakType == null || breakType == BreakValues.TextWrapping) {
                AppendFormattedNoteText(text, runs, LegacyDocSpecialCharacters.TextWrappingBreak.ToString(), formatting, storyStart);
                return;
            }

            if (breakType == BreakValues.Page) {
                AppendFormattedNoteText(text, runs, LegacyDocSpecialCharacters.PageBreak.ToString(), formatting, storyStart);
                return;
            }

            if (breakType == BreakValues.Column) {
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
                if (previous.EndCharacter == start && previous.Formatting.Equals(formatting)) {
                    runs[runs.Count - 1] = previous.Extend(length);
                    return;
                }
            }

            runs.Add(new LegacyDocWritableRun(start, length, formatting));
        }
    }
}
