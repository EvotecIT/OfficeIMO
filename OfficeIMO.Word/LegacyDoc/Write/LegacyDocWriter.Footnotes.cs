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

            var stories = new Dictionary<long, string>();
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

        private static string ReadSimpleFootnoteStory(Footnote footnote, long id) {
            var builder = new StringBuilder();
            builder.Append(LegacyDocFootnoteReader.FootnoteReferenceCharacter);
            builder.Append(' ');
            bool hasBodyText = false;
            foreach (OpenXmlElement child in footnote.ChildElements) {
                switch (child) {
                    case Paragraph paragraph:
                        string paragraphText = ReadSimpleFootnoteParagraph(paragraph, id);
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
            return builder.ToString();
        }

        private static string ReadSimpleFootnoteParagraph(Paragraph paragraph, long id) {
            var builder = new StringBuilder();
            foreach (OpenXmlElement child in paragraph.ChildElements) {
                switch (child) {
                    case ParagraphProperties paragraphProperties:
                        ThrowIfUnsupportedFootnoteParagraphProperties(paragraphProperties, id);
                        break;
                    case Run run:
                        AppendSimpleFootnoteRun(builder, run, id);
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

        private static void AppendSimpleFootnoteRun(StringBuilder builder, Run run, long id) {
            if (IsFootnoteReferenceMarkRun(run)) {
                return;
            }

            if (run.RunProperties != null && run.RunProperties.HasChildren) {
                throw new NotSupportedException($"Native DOC saving supports simple footnote id '{id}' only without formatted footnote text runs.");
            }

            foreach (OpenXmlElement child in run.ChildElements) {
                switch (child) {
                    case RunProperties:
                        break;
                    case Text text:
                        builder.Append(text.Text);
                        break;
                    case TabChar:
                        builder.Append('\t');
                        break;
                    case Break breakNode:
                        AppendSimpleFootnoteBreak(builder, breakNode, id);
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

        private static void AppendSimpleFootnoteBreak(StringBuilder builder, Break breakNode, long id) {
            BreakValues? breakType = breakNode.Type?.Value;
            if (breakType == null || breakType == BreakValues.TextWrapping) {
                builder.Append('\v');
                return;
            }

            throw new NotSupportedException($"Native DOC saving supports simple footnote id '{id}' only with text-wrapping breaks.");
        }

        private sealed class LegacyDocWritableFootnotes {
            internal static readonly LegacyDocWritableFootnotes Empty = new LegacyDocWritableFootnotes(new Dictionary<long, string>());

            private readonly Dictionary<long, string> _storiesById;
            private readonly List<LegacyDocWritableFootnoteReference> _references = new List<LegacyDocWritableFootnoteReference>();
            private readonly HashSet<long> _referencedIds = new HashSet<long>();

            internal LegacyDocWritableFootnotes(Dictionary<long, string> storiesById) {
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
                var textPositions = new List<int>(_references.Count + 2);
                var markerPositions = new List<int>(_references.Count);
                for (int index = 0; index < _references.Count; index++) {
                    LegacyDocWritableFootnoteReference reference = _references[index];
                    string story = _storiesById[reference.Id];
                    textPositions.Add(text.Length);
                    markerPositions.Add(text.Length);
                    text.Append(story);
                }

                textPositions.Add(text.Length - 1);
                textPositions.Add(text.Length + 2);
                return new LegacyDocWritableFootnoteStories(
                    text.ToString(),
                    CreateFootnoteReferencePlc(_references, bodyCharacterCount + text.Length + headerFooterCharacterCount + terminalCharacterPadding),
                    CreateFootnoteTextPlc(textPositions),
                    markerPositions);
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
            internal static readonly LegacyDocWritableFootnoteStories Empty = new LegacyDocWritableFootnoteStories(string.Empty, Array.Empty<byte>(), Array.Empty<byte>(), Array.Empty<int>());

            internal LegacyDocWritableFootnoteStories(string text, byte[] plcffndRef, byte[] plcffndTxt, IReadOnlyList<int> markerPositions) {
                Text = text;
                PlcffndRef = plcffndRef;
                PlcffndTxt = plcffndTxt;
                MarkerPositions = markerPositions;
            }

            internal string Text { get; }

            internal byte[] PlcffndRef { get; }

            internal byte[] PlcffndTxt { get; }

            internal IReadOnlyList<int> MarkerPositions { get; }
        }
    }
}
