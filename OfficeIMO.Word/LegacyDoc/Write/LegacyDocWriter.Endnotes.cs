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

            var stories = new Dictionary<long, string>();
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

                stories.Add(id.Value, ReadSimpleEndnoteStory(endnote, id.Value));
            }

            return stories.Count == 0
                ? LegacyDocWritableEndnotes.Empty
                : new LegacyDocWritableEndnotes(stories);
        }

        private static string ReadSimpleEndnoteStory(Endnote endnote, long id) {
            var builder = new StringBuilder();
            builder.Append(LegacyDocFootnoteReader.FootnoteReferenceCharacter);
            builder.Append(' ');
            bool hasBodyText = false;
            foreach (OpenXmlElement child in endnote.ChildElements) {
                switch (child) {
                    case Paragraph paragraph:
                        string paragraphText = ReadSimpleEndnoteParagraph(paragraph, id);
                        if (!string.IsNullOrEmpty(paragraphText)) {
                            hasBodyText = true;
                        }

                        builder.Append(paragraphText);
                        builder.Append('\r');
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple endnote paragraphs only. Unsupported endnote element: {child.LocalName}.");
                }
            }

            if (!hasBodyText) {
                throw new NotSupportedException($"Native DOC saving cannot write empty endnote id '{id}'.");
            }

            builder.Append('\r');
            return builder.ToString();
        }

        private static string ReadSimpleEndnoteParagraph(Paragraph paragraph, long id) {
            var builder = new StringBuilder();
            foreach (OpenXmlElement child in paragraph.ChildElements) {
                switch (child) {
                    case ParagraphProperties paragraphProperties:
                        ThrowIfUnsupportedEndnoteParagraphProperties(paragraphProperties, id);
                        break;
                    case Run run:
                        AppendSimpleEndnoteRun(builder, run, id);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple endnote paragraphs only. Unsupported endnote paragraph element: {child.LocalName}.");
                }
            }

            return builder.ToString();
        }

        private static void ThrowIfUnsupportedEndnoteParagraphProperties(ParagraphProperties paragraphProperties, long id) {
            foreach (OpenXmlElement property in paragraphProperties.ChildElements) {
                switch (property) {
                    case ParagraphStyleId paragraphStyleId:
                        string? value = paragraphStyleId.Val?.Value;
                        if (string.IsNullOrWhiteSpace(value)
                            || string.Equals(value, "EndnoteText", StringComparison.OrdinalIgnoreCase)) {
                            break;
                        }

                        throw new NotSupportedException($"Native DOC saving supports simple endnote id '{id}' only with the EndnoteText paragraph style.");
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple endnote id '{id}' only without paragraph formatting. Unsupported endnote paragraph property: {property.LocalName}.");
                }
            }
        }

        private static void AppendSimpleEndnoteRun(StringBuilder builder, Run run, long id) {
            if (IsEndnoteReferenceMarkRun(run)) {
                return;
            }

            if (run.RunProperties != null && run.RunProperties.HasChildren) {
                throw new NotSupportedException($"Native DOC saving supports simple endnote id '{id}' only without formatted endnote text runs.");
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
                        AppendSimpleEndnoteBreak(builder, breakNode, id);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple endnote id '{id}' only with text, tabs, and simple line breaks. Unsupported endnote run element: {child.LocalName}.");
                }
            }
        }

        private static bool IsEndnoteReferenceMarkRun(Run run) {
            bool hasReferenceMark = false;
            foreach (OpenXmlElement child in run.ChildElements) {
                switch (child) {
                    case RunProperties:
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

        private static void AppendSimpleEndnoteBreak(StringBuilder builder, Break breakNode, long id) {
            BreakValues? breakType = breakNode.Type?.Value;
            if (breakType == null || breakType == BreakValues.TextWrapping) {
                builder.Append('\v');
                return;
            }

            throw new NotSupportedException($"Native DOC saving supports simple endnote id '{id}' only with text-wrapping breaks.");
        }

        private sealed class LegacyDocWritableEndnotes {
            internal static readonly LegacyDocWritableEndnotes Empty = new LegacyDocWritableEndnotes(new Dictionary<long, string>());

            private readonly Dictionary<long, string> _storiesById;
            private readonly List<LegacyDocWritableEndnoteReference> _references = new List<LegacyDocWritableEndnoteReference>();
            private readonly HashSet<long> _referencedIds = new HashSet<long>();

            internal LegacyDocWritableEndnotes(Dictionary<long, string> storiesById) {
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
                var textPositions = new List<int>(_references.Count + 2);
                var markerPositions = new List<int>(_references.Count);
                for (int index = 0; index < _references.Count; index++) {
                    LegacyDocWritableEndnoteReference reference = _references[index];
                    string story = _storiesById[reference.Id];
                    textPositions.Add(text.Length);
                    markerPositions.Add(text.Length);
                    text.Append(story);
                }

                textPositions.Add(text.Length - 1);
                textPositions.Add(text.Length + 2);
                return new LegacyDocWritableEndnoteStories(
                    text.ToString(),
                    CreateEndnoteReferencePlc(_references, bodyCharacterCount + footnoteCharacterCount + headerFooterCharacterCount + text.Length + terminalCharacterPadding),
                    CreateEndnoteTextPlc(textPositions),
                    markerPositions);
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
            internal static readonly LegacyDocWritableEndnoteStories Empty = new LegacyDocWritableEndnoteStories(string.Empty, Array.Empty<byte>(), Array.Empty<byte>(), Array.Empty<int>());

            internal LegacyDocWritableEndnoteStories(string text, byte[] plcfendRef, byte[] plcfendTxt, IReadOnlyList<int> markerPositions) {
                Text = text;
                PlcfendRef = plcfendRef;
                PlcfendTxt = plcfendTxt;
                MarkerPositions = markerPositions;
            }

            internal string Text { get; }

            internal byte[] PlcfendRef { get; }

            internal byte[] PlcfendTxt { get; }

            internal IReadOnlyList<int> MarkerPositions { get; }
        }
    }
}
