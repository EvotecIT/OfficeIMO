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
                    case BookmarkStart bookmarkStart:
                        bookmarks.AddStart(bookmarkStart, builder.Length);
                        break;
                    case BookmarkEnd bookmarkEnd:
                        bookmarks.AddEnd(bookmarkEnd, builder.Length);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple footnote paragraphs and bookmarks only. Unsupported footnote element: {child.LocalName}.");
                }
            }

            if (!hasBodyText) {
                throw new NotSupportedException($"Native DOC saving cannot write empty footnote id '{id}'.");
            }

            builder.Append('\r');
            return new LegacyDocWritableNoteStory(builder.ToString(), runs, formattedParagraphs, bookmarks.Create());
        }

        private static LegacyDocWritableParagraphFormatting ReadSimpleFootnoteParagraph(Paragraph paragraph, long id, List<LegacyDocWritableRun> runs, LegacyDocWritableBookmarksBuilder bookmarks, int storyStart, bool isFirstParagraph, FootnotesPart relationshipOwner, out string paragraphText) {
            var builder = new StringBuilder();
            LegacyDocWritableParagraphFormatting paragraphFormatting = ReadSupportedNoteParagraphFormatting(paragraph.ParagraphProperties, id, "footnote", FootnoteParagraphStyleIndexes);
            if (isFirstParagraph && paragraphFormatting.HasFormatting && paragraphFormatting.StyleIndex == null) {
                paragraphFormatting = paragraphFormatting.WithStyleIndex(NoteTextParagraphStyleIndex);
            }

            foreach (OpenXmlElement child in paragraph.ChildElements) {
                switch (child) {
                    case ParagraphProperties:
                        break;
                    case Run run:
                        AppendSimpleFootnoteRun(builder, runs, run, id, storyStart);
                        break;
                    case Hyperlink hyperlink:
                        AppendSupportedNoteHyperlinkText(builder, runs, hyperlink, relationshipOwner, id, "footnote", storyStart);
                        break;
                    case BookmarkStart bookmarkStart:
                        bookmarks.AddStart(bookmarkStart, storyStart + builder.Length);
                        break;
                    case BookmarkEnd bookmarkEnd:
                        bookmarks.AddEnd(bookmarkEnd, storyStart + builder.Length);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple footnote paragraphs only with text runs, bookmarks, and simple hyperlinks. Unsupported footnote paragraph element: {child.LocalName}.");
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
                    case Text text:
                        AppendFormattedNoteText(builder, runs, text.Text, formatting, storyStart);
                        break;
                    case TabChar:
                        AppendFormattedNoteText(builder, runs, "\t", formatting, storyStart);
                        break;
                    case Break breakNode:
                        AppendSimpleFootnoteBreak(builder, runs, breakNode, id, formatting, storyStart);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple footnote id '{id}' only with text, tabs, and simple line breaks. Unsupported footnote run element: {child.LocalName}.");
                }
            }
        }

        private static bool IsFootnoteReferenceMarkRun(Run run) {
            bool hasReferenceMark = false;
            foreach (OpenXmlElement child in run.ChildElements) {
                switch (child) {
                    case RunProperties:
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
                AppendFormattedNoteText(builder, runs, "\v", formatting, storyStart);
                return;
            }

            if (breakType == BreakValues.Page) {
                AppendFormattedNoteText(builder, runs, "\f", formatting, storyStart);
                return;
            }

            throw new NotSupportedException($"Native DOC saving supports simple footnote id '{id}' only with text-wrapping and page breaks.");
        }

        private static void AppendSupportedNoteHyperlinkText(
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
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
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple {noteKind} hyperlinks only when they contain text runs. Unsupported hyperlink element: {child.LocalName}.");
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
                    case Text textNode:
                        AppendFormattedNoteText(text, runs, textNode.Text, formatting, storyStart);
                        break;
                    case TabChar:
                        AppendFormattedNoteText(text, runs, "\t", formatting, storyStart);
                        break;
                    case Break breakNode:
                        AppendSupportedNoteHyperlinkBreak(text, runs, breakNode, id, noteKind, storyStart, formatting);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple {noteKind} id '{id}' hyperlinks only with text, tabs, and text-wrapping break display runs. Unsupported hyperlink run element: {child.LocalName}.");
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
                AppendFormattedNoteText(text, runs, "\v", formatting, storyStart);
                return;
            }

            if (breakType == BreakValues.Page) {
                AppendFormattedNoteText(text, runs, "\f", formatting, storyStart);
                return;
            }

            throw new NotSupportedException($"Native DOC saving supports simple {noteKind} id '{id}' hyperlinks only with text-wrapping and page breaks.");
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
