using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
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

                stories.Add(id.Value, ReadSimpleFootnoteStory(footnote, id.Value));
            }

            return stories.Count == 0
                ? LegacyDocWritableFootnotes.Empty
                : new LegacyDocWritableFootnotes(stories);
        }

        private static bool IsUserFootnote(Footnote footnote) {
            return footnote.Type == null || footnote.Type.Value == FootnoteEndnoteValues.Normal;
        }

        private static LegacyDocWritableNoteStory ReadSimpleFootnoteStory(Footnote footnote, long id) {
            var builder = new StringBuilder();
            var runs = new List<LegacyDocWritableRun>();
            builder.Append(LegacyDocFootnoteReader.FootnoteReferenceCharacter);
            builder.Append(' ');
            bool hasBodyText = false;
            foreach (OpenXmlElement child in footnote.ChildElements) {
                switch (child) {
                    case Paragraph paragraph:
                        string paragraphText = ReadSimpleFootnoteParagraph(paragraph, id, runs, builder.Length);
                        if (!string.IsNullOrEmpty(paragraphText)) {
                            hasBodyText = true;
                        }

                        builder.Append(paragraphText);
                        builder.Append('\r');
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple footnote paragraphs only. Unsupported footnote element: {child.LocalName}.");
                }
            }

            if (!hasBodyText) {
                throw new NotSupportedException($"Native DOC saving cannot write empty footnote id '{id}'.");
            }

            builder.Append('\r');
            return new LegacyDocWritableNoteStory(builder.ToString(), runs);
        }

        private static string ReadSimpleFootnoteParagraph(Paragraph paragraph, long id, List<LegacyDocWritableRun> runs, int storyStart) {
            var builder = new StringBuilder();
            foreach (OpenXmlElement child in paragraph.ChildElements) {
                switch (child) {
                    case ParagraphProperties paragraphProperties:
                        ThrowIfUnsupportedFootnoteParagraphProperties(paragraphProperties, id);
                        break;
                    case Run run:
                        AppendSimpleFootnoteRun(builder, runs, run, id, storyStart);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple footnote paragraphs only. Unsupported footnote paragraph element: {child.LocalName}.");
                }
            }

            return builder.ToString();
        }

        private static void ThrowIfUnsupportedFootnoteParagraphProperties(ParagraphProperties paragraphProperties, long id) {
            foreach (OpenXmlElement property in paragraphProperties.ChildElements) {
                switch (property) {
                    case ParagraphStyleId paragraphStyleId:
                        string? value = paragraphStyleId.Val?.Value;
                        if (string.IsNullOrWhiteSpace(value)
                            || string.Equals(value, "FootnoteText", StringComparison.OrdinalIgnoreCase)) {
                            break;
                        }

                        throw new NotSupportedException($"Native DOC saving supports simple footnote id '{id}' only with the FootnoteText paragraph style.");
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple footnote id '{id}' only without paragraph formatting. Unsupported footnote paragraph property: {property.LocalName}.");
                }
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

            throw new NotSupportedException($"Native DOC saving supports simple footnote id '{id}' only with text-wrapping breaks.");
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

                    text.Append(story.Text);
                }

                textPositions.Add(text.Length - 1);
                textPositions.Add(text.Length + 2);
                return new LegacyDocWritableFootnoteStories(
                    text.ToString(),
                    CreateFootnoteReferencePlc(_references, bodyCharacterCount + text.Length + headerFooterCharacterCount + terminalCharacterPadding),
                    CreateFootnoteTextPlc(textPositions),
                    markerPositions,
                    runs);
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
            internal static readonly LegacyDocWritableFootnoteStories Empty = new LegacyDocWritableFootnoteStories(string.Empty, Array.Empty<byte>(), Array.Empty<byte>(), Array.Empty<int>(), Array.Empty<LegacyDocWritableRun>());

            internal LegacyDocWritableFootnoteStories(string text, byte[] plcffndRef, byte[] plcffndTxt, IReadOnlyList<int> markerPositions, IReadOnlyList<LegacyDocWritableRun> formattedRuns) {
                Text = text;
                PlcffndRef = plcffndRef;
                PlcffndTxt = plcffndTxt;
                MarkerPositions = markerPositions;
                FormattedRuns = formattedRuns;
            }

            internal string Text { get; }

            internal byte[] PlcffndRef { get; }

            internal byte[] PlcffndTxt { get; }

            internal IReadOnlyList<int> MarkerPositions { get; }

            internal IReadOnlyList<LegacyDocWritableRun> FormattedRuns { get; }
        }

        private readonly struct LegacyDocWritableNoteStory {
            internal LegacyDocWritableNoteStory(string text, IReadOnlyList<LegacyDocWritableRun> formattedRuns) {
                Text = text;
                FormattedRuns = formattedRuns;
            }

            internal string Text { get; }

            internal IReadOnlyList<LegacyDocWritableRun> FormattedRuns { get; }
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
