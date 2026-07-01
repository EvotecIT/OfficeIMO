using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private const int HeaderFooterSeparatorStoryCount = 6;
        private const int HeaderFooterStoriesPerSection = 6;
        private const string HeaderFooterAutoNumberSeparatorStory = "\u0003\r\r";
        private const string HeaderFooterContinuationSeparatorStory = "\u0004\r\r";
        private const string HeaderFooterSeparatorTerminator = "\r";

        private static LegacyDocWritableHeaderFooterStories BuildHeaderFooterStories(WordDocument document, MainDocumentPart mainPart, bool includeDefaultSeparators, IReadOnlyDictionary<string, ushort> styleIndexes) {
            if (!includeDefaultSeparators && !mainPart.HeaderParts.Any() && !mainPart.FooterParts.Any()) {
                return LegacyDocWritableHeaderFooterStories.Empty;
            }

            if (!includeDefaultSeparators && document.Sections.Count == 0) {
                return LegacyDocWritableHeaderFooterStories.Empty;
            }

            int sectionCount = Math.Max(document.Sections.Count, 1);
            LegacyDocWritableHeaderFooterStory[] stories = new LegacyDocWritableHeaderFooterStory[HeaderFooterSeparatorStoryCount + (sectionCount * HeaderFooterStoriesPerSection)];
            for (int storyIndex = 0; storyIndex < stories.Length; storyIndex++) {
                stories[storyIndex] = LegacyDocWritableHeaderFooterStory.Empty;
            }

            if (includeDefaultSeparators) {
                stories[0] = LegacyDocWritableHeaderFooterStory.Plain(HeaderFooterAutoNumberSeparatorStory);
                stories[1] = LegacyDocWritableHeaderFooterStory.Plain(HeaderFooterContinuationSeparatorStory);
                stories[3] = LegacyDocWritableHeaderFooterStory.Plain(HeaderFooterAutoNumberSeparatorStory);
                stories[4] = LegacyDocWritableHeaderFooterStory.Plain(HeaderFooterContinuationSeparatorStory);
            }

            for (int sectionIndex = 0; sectionIndex < document.Sections.Count; sectionIndex++) {
                SectionProperties sectionProperties = document.Sections[sectionIndex]._sectionProperties;
                int sectionStoryOffset = HeaderFooterSeparatorStoryCount + (sectionIndex * HeaderFooterStoriesPerSection);
                stories[sectionStoryOffset] = ReadHeaderStory(mainPart, sectionProperties, HeaderFooterValues.Even, styleIndexes) ?? LegacyDocWritableHeaderFooterStory.Empty;
                stories[sectionStoryOffset + 1] = ReadHeaderStory(mainPart, sectionProperties, HeaderFooterValues.Default, styleIndexes) ?? LegacyDocWritableHeaderFooterStory.Empty;
                stories[sectionStoryOffset + 2] = ReadFooterStory(mainPart, sectionProperties, HeaderFooterValues.Even, styleIndexes) ?? LegacyDocWritableHeaderFooterStory.Empty;
                stories[sectionStoryOffset + 3] = ReadFooterStory(mainPart, sectionProperties, HeaderFooterValues.Default, styleIndexes) ?? LegacyDocWritableHeaderFooterStory.Empty;
                stories[sectionStoryOffset + 4] = ReadHeaderStory(mainPart, sectionProperties, HeaderFooterValues.First, styleIndexes) ?? LegacyDocWritableHeaderFooterStory.Empty;
                stories[sectionStoryOffset + 5] = ReadFooterStory(mainPart, sectionProperties, HeaderFooterValues.First, styleIndexes) ?? LegacyDocWritableHeaderFooterStory.Empty;
            }

            ThrowIfUnreferencedHeaderFooterContent(mainPart, document.Sections.Select(section => section._sectionProperties));
            if (!includeDefaultSeparators && stories.All(story => story.Text.Length == 0)) {
                return LegacyDocWritableHeaderFooterStories.Empty;
            }

            var text = new StringBuilder();
            var characterPositions = new List<int>(stories.Length + 2);
            var markerPositions = new List<int>();
            var formattedRuns = new List<LegacyDocWritableRun>();
            var formattedParagraphs = new List<LegacyDocWritableParagraph>();
            var bookmarks = new LegacyDocWritableBookmarksBuilder();
            foreach (LegacyDocWritableHeaderFooterStory story in stories) {
                characterPositions.Add(text.Length);
                AddHeaderFooterSpecialCharacterPositions(text.Length, story.Text, markerPositions);
                foreach (LegacyDocWritableRun run in story.FormattedRuns) {
                    formattedRuns.Add(new LegacyDocWritableRun(text.Length + run.StartCharacter, run.Length, run.Formatting));
                }

                foreach (LegacyDocWritableParagraph paragraph in story.FormattedParagraphs) {
                    formattedParagraphs.Add(new LegacyDocWritableParagraph(text.Length + paragraph.StartCharacter, paragraph.Length, paragraph.Formatting));
                }

                bookmarks.AddRange(story.Bookmarks, text.Length);
                text.Append(story.Text);
            }

            if (includeDefaultSeparators) {
                text.Append(HeaderFooterSeparatorTerminator);
            }

            characterPositions.Add(includeDefaultSeparators ? text.Length - HeaderFooterSeparatorTerminator.Length : text.Length);
            characterPositions.Add(includeDefaultSeparators ? text.Length + 2 : text.Length);
            byte[] plcfHdd = new byte[characterPositions.Count * 4];
            for (int index = 0; index < characterPositions.Count; index++) {
                WriteInt32(plcfHdd, index * 4, characterPositions[index]);
            }

            return new LegacyDocWritableHeaderFooterStories(text.ToString(), plcfHdd, markerPositions, formattedRuns, formattedParagraphs, bookmarks.Create());
        }

        private static void AddHeaderFooterSpecialCharacterPositions(int storyStart, string story, List<int> markerPositions) {
            for (int index = 0; index < story.Length; index++) {
                char character = story[index];
                if (character == '\u0003' || character == '\u0004') {
                    markerPositions.Add(storyStart + index);
                }
            }
        }

        private static LegacyDocWritableHeaderFooterStory? ReadHeaderStory(MainDocumentPart mainPart, SectionProperties sectionProperties, HeaderFooterValues type, IReadOnlyDictionary<string, ushort> styleIndexes) {
            HeaderReference[] references = sectionProperties.Elements<HeaderReference>().ToArray();
            if (references.Length == 0) {
                return null;
            }

            string kind = GetHeaderFooterStoryDescription(type, "header");
            HeaderReference? reference = GetSingleReference(references, type, "header");
            if (reference == null) {
                return null;
            }

            HeaderPart headerPart = GetReferencedPart<HeaderPart>(mainPart, reference.Id?.Value, kind);
            return ReadSimpleHeaderFooterStory(headerPart.Header, headerPart, kind, styleIndexes);
        }

        private static LegacyDocWritableHeaderFooterStory? ReadFooterStory(MainDocumentPart mainPart, SectionProperties sectionProperties, HeaderFooterValues type, IReadOnlyDictionary<string, ushort> styleIndexes) {
            FooterReference[] references = sectionProperties.Elements<FooterReference>().ToArray();
            if (references.Length == 0) {
                return null;
            }

            string kind = GetHeaderFooterStoryDescription(type, "footer");
            FooterReference? reference = GetSingleReference(references, type, "footer");
            if (reference == null) {
                return null;
            }

            FooterPart footerPart = GetReferencedPart<FooterPart>(mainPart, reference.Id?.Value, kind);
            return ReadSimpleHeaderFooterStory(footerPart.Footer, footerPart, kind, styleIndexes);
        }

        private static TPart GetReferencedPart<TPart>(MainDocumentPart mainPart, string? relationshipId, string kind)
            where TPart : OpenXmlPart {
            if (string.IsNullOrWhiteSpace(relationshipId)) {
                throw new NotSupportedException($"Native DOC saving cannot write a {kind} reference without a relationship id.");
            }

            OpenXmlPart part;
            try {
                part = mainPart.GetPartById(relationshipId!);
            } catch (ArgumentOutOfRangeException exception) {
                throw new NotSupportedException($"Native DOC saving cannot write a {kind} reference that points to a missing part.", exception);
            }

            if (part is TPart typedPart) {
                return typedPart;
            }

            throw new NotSupportedException($"Native DOC saving cannot write a {kind} reference that points to an unexpected part type.");
        }

        private static TReference? GetSingleReference<TReference>(IReadOnlyList<TReference> references, HeaderFooterValues requestedType, string kind)
            where TReference : HeaderFooterReferenceType {
            TReference? requestedReference = default;
            foreach (TReference reference in references) {
                HeaderFooterValues type = reference.Type?.Value ?? HeaderFooterValues.Default;
                if (!IsSupportedHeaderFooterType(type)) {
                    throw new NotSupportedException($"Native DOC saving supports only default, first-page, and even-page {kind}s.");
                }

                if (type != requestedType) {
                    continue;
                }

                if (requestedReference != null) {
                    throw new NotSupportedException($"Native DOC saving cannot write multiple {GetHeaderFooterDescription(requestedType, kind)} references in one section.");
                }

                requestedReference = reference;
            }

            return requestedReference;
        }

        private static bool IsSupportedHeaderFooterType(HeaderFooterValues type) {
            return type == HeaderFooterValues.Default
                || type == HeaderFooterValues.First
                || type == HeaderFooterValues.Even;
        }

        private static string GetHeaderFooterDescription(HeaderFooterValues type, string kind) {
            if (type == HeaderFooterValues.First) {
                return $"first-page {kind}";
            }

            if (type == HeaderFooterValues.Even) {
                return $"even-page {kind}";
            }

            return $"default {kind}";
        }

        private static string GetHeaderFooterStoryDescription(HeaderFooterValues type, string kind) {
            if (type == HeaderFooterValues.Default) {
                return kind;
            }

            return GetHeaderFooterDescription(type, kind);
        }

        private static LegacyDocWritableHeaderFooterStory? ReadSimpleHeaderFooterStory(OpenXmlCompositeElement? container, OpenXmlPartContainer relationshipOwner, string kind, IReadOnlyDictionary<string, ushort> styleIndexes) {
            if (container == null || !container.HasChildren) {
                return null;
            }

            var paragraphs = new List<string>();
            var storyText = new StringBuilder();
            var formattedRuns = new List<LegacyDocWritableRun>();
            var formattedParagraphs = new List<LegacyDocWritableParagraph>();
            var bookmarks = new LegacyDocWritableBookmarksBuilder();
            foreach (OpenXmlElement child in container.ChildElements) {
                switch (child) {
                    case Paragraph paragraph:
                        int paragraphStart = storyText.Length;
                        LegacyDocWritableParagraphFormatting paragraphFormatting = ReadSimpleHeaderFooterParagraph(storyText, formattedRuns, bookmarks, paragraph, relationshipOwner, kind, styleIndexes, out string paragraphText);
                        paragraphs.Add(paragraphText);
                        storyText.Append('\r');
                        if (paragraphFormatting.HasFormatting) {
                            formattedParagraphs.Add(new LegacyDocWritableParagraph(paragraphStart, storyText.Length - paragraphStart, paragraphFormatting));
                        }

                        break;
                    case BookmarkStart bookmarkStart:
                        bookmarks.AddStart(bookmarkStart, storyText.Length);
                        break;
                    case BookmarkEnd bookmarkEnd:
                        bookmarks.AddEnd(bookmarkEnd, storyText.Length);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving currently supports only text paragraphs and bookmarks in {kind}s. Unsupported {kind} element: {child.LocalName}.");
                }
            }

            bool hasVisibleText = paragraphs.Any(paragraph => paragraph.Length > 0);
            if (!hasVisibleText) {
                return null;
            }

            storyText.Append('\r');
            return new LegacyDocWritableHeaderFooterStory(storyText.ToString(), formattedRuns, formattedParagraphs, bookmarks.Create());
        }

        private static LegacyDocWritableParagraphFormatting ReadSimpleHeaderFooterParagraph(StringBuilder storyText, List<LegacyDocWritableRun> formattedRuns, LegacyDocWritableBookmarksBuilder bookmarks, Paragraph paragraph, OpenXmlPartContainer relationshipOwner, string kind, IReadOnlyDictionary<string, ushort> styleIndexes, out string paragraphText) {
            var text = new StringBuilder();
            LegacyDocWritableParagraphFormatting paragraphFormatting = ReadSupportedParagraphFormatting(paragraph.ParagraphProperties, styleIndexes);
            foreach (OpenXmlElement child in paragraph.ChildElements) {
                switch (child) {
                    case ParagraphProperties:
                        break;
                    case Run run:
                        AppendFormattedHeaderFooterRun(storyText, formattedRuns, text, run, kind);
                        break;
                    case Hyperlink hyperlink:
                        AppendFormattedHeaderFooterHyperlink(storyText, formattedRuns, text, hyperlink, relationshipOwner, kind);
                        break;
                    case BookmarkStart bookmarkStart:
                        bookmarks.AddStart(bookmarkStart, storyText.Length);
                        break;
                    case BookmarkEnd bookmarkEnd:
                        bookmarks.AddEnd(bookmarkEnd, storyText.Length);
                        break;
                    default:
                        if (IsIgnorableParagraphMarkup(child)) {
                            break;
                        }

                        throw new NotSupportedException($"Native DOC saving currently supports only text runs, bookmarks, and simple hyperlinks with supported direct formatting in {kind}s. Unsupported {kind} paragraph element: {child.LocalName}.");
                }
            }

            paragraphText = text.ToString();
            return paragraphFormatting;
        }

        private static void AppendFormattedHeaderFooterHyperlink(StringBuilder storyText, List<LegacyDocWritableRun> formattedRuns, StringBuilder paragraphText, Hyperlink hyperlink, OpenXmlPartContainer relationshipOwner, string kind) {
            int before = storyText.Length;
            try {
                AppendSupportedHyperlinkText(storyText, formattedRuns, hyperlink, relationshipOwner, LegacyDocWritableFootnotes.Empty, LegacyDocWritableEndnotes.Empty);
            } catch (NotSupportedException exception) {
                throw new NotSupportedException($"Native DOC saving supports simple {kind} hyperlinks only when they are external plain-text hyperlinks. {exception.Message}", exception);
            }

            if (storyText.Length > before) {
                paragraphText.Append(storyText.ToString(before, storyText.Length - before));
            }
        }

        private static void AppendFormattedHeaderFooterRun(StringBuilder storyText, List<LegacyDocWritableRun> formattedRuns, StringBuilder paragraphText, Run run, string kind) {
            LegacyDocWritableFormatting formatting = ReadSupportedRunFormatting(run.RunProperties);
            foreach (OpenXmlElement child in run.ChildElements) {
                switch (child) {
                    case RunProperties:
                        break;
                    case LastRenderedPageBreak:
                        break;
                    case Text textNode:
                        AppendFormattedHeaderFooterText(storyText, formattedRuns, paragraphText, textNode.Text, formatting);
                        break;
                    case TabChar:
                        AppendFormattedHeaderFooterText(storyText, formattedRuns, paragraphText, "\t", formatting);
                        break;
                    case CarriageReturn:
                        AppendFormattedHeaderFooterText(storyText, formattedRuns, paragraphText, LegacyDocSpecialCharacters.TextWrappingBreak.ToString(), formatting);
                        break;
                    case NoBreakHyphen:
                        AppendFormattedHeaderFooterText(storyText, formattedRuns, paragraphText, LegacyDocSpecialCharacters.NoBreakHyphen.ToString(), formatting);
                        break;
                    case SoftHyphen:
                        AppendFormattedHeaderFooterText(storyText, formattedRuns, paragraphText, LegacyDocSpecialCharacters.SoftHyphen.ToString(), formatting);
                        break;
                    case Break breakNode:
                        AppendFormattedHeaderFooterBreak(storyText, formattedRuns, paragraphText, breakNode, kind, formatting);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving currently supports only text, tabs, carriage returns, soft/no-break hyphens, and text-wrapping/page/column breaks in {kind}s. Unsupported {kind} run element: {child.LocalName}.");
                }
            }
        }

        private static void AppendFormattedHeaderFooterText(StringBuilder storyText, List<LegacyDocWritableRun> formattedRuns, StringBuilder paragraphText, string? value, LegacyDocWritableFormatting formatting) {
            int before = storyText.Length;
            AppendFormattedNoteText(storyText, formattedRuns, value, formatting, storyStart: 0);
            if (storyText.Length > before) {
                paragraphText.Append(storyText.ToString(before, storyText.Length - before));
            }
        }

        private static void AppendFormattedHeaderFooterBreak(StringBuilder storyText, List<LegacyDocWritableRun> formattedRuns, StringBuilder paragraphText, Break breakNode, string kind, LegacyDocWritableFormatting formatting) {
            BreakValues? breakType = breakNode.Type?.Value;
            if (breakType == null || breakType == BreakValues.TextWrapping) {
                AppendFormattedHeaderFooterText(storyText, formattedRuns, paragraphText, LegacyDocSpecialCharacters.TextWrappingBreak.ToString(), formatting);
                return;
            }

            if (breakType == BreakValues.Page) {
                AppendFormattedHeaderFooterText(storyText, formattedRuns, paragraphText, LegacyDocSpecialCharacters.PageBreak.ToString(), formatting);
                return;
            }

            if (breakType == BreakValues.Column) {
                AppendFormattedHeaderFooterText(storyText, formattedRuns, paragraphText, LegacyDocSpecialCharacters.ColumnBreak.ToString(), formatting);
                return;
            }

            throw new NotSupportedException($"Native DOC saving supports simple {kind}s only with text-wrapping, page, and column breaks.");
        }

        private static void ThrowIfUnreferencedHeaderFooterContent(MainDocumentPart mainPart, IEnumerable<SectionProperties> sectionPropertiesCollection) {
            var referencedIds = new HashSet<string>(StringComparer.Ordinal);
            foreach (SectionProperties sectionProperties in sectionPropertiesCollection) {
                foreach (HeaderReference reference in sectionProperties.Elements<HeaderReference>()) {
                    string? id = reference.Id?.Value;
                    if (!string.IsNullOrWhiteSpace(id)) {
                        referencedIds.Add(id!);
                    }
                }

                foreach (FooterReference reference in sectionProperties.Elements<FooterReference>()) {
                    string? id = reference.Id?.Value;
                    if (!string.IsNullOrWhiteSpace(id)) {
                        referencedIds.Add(id!);
                    }
                }
            }

            foreach (HeaderPart part in mainPart.HeaderParts) {
                string id = mainPart.GetIdOfPart(part);
                if (!referencedIds.Contains(id) && HasVisibleHeaderFooterContent(part.Header)) {
                    throw new NotSupportedException("Native DOC saving cannot preserve unreferenced header content.");
                }
            }

            foreach (FooterPart part in mainPart.FooterParts) {
                string id = mainPart.GetIdOfPart(part);
                if (!referencedIds.Contains(id) && HasVisibleHeaderFooterContent(part.Footer)) {
                    throw new NotSupportedException("Native DOC saving cannot preserve unreferenced footer content.");
                }
            }
        }

        private static bool HasVisibleHeaderFooterContent(OpenXmlCompositeElement? container) {
            return container != null && container.Descendants<Text>().Any(text => !string.IsNullOrEmpty(text.Text));
        }

        private readonly struct LegacyDocWritableHeaderFooterStories {
            internal static LegacyDocWritableHeaderFooterStories Empty { get; } = new LegacyDocWritableHeaderFooterStories(string.Empty, Array.Empty<byte>(), Array.Empty<int>(), Array.Empty<LegacyDocWritableRun>(), Array.Empty<LegacyDocWritableParagraph>(), LegacyDocWritableBookmarks.Empty);

            internal LegacyDocWritableHeaderFooterStories(string text, byte[] plcfHdd, IReadOnlyList<int> markerPositions, IReadOnlyList<LegacyDocWritableRun> formattedRuns, IReadOnlyList<LegacyDocWritableParagraph> formattedParagraphs, LegacyDocWritableBookmarks bookmarks) {
                Text = text;
                PlcfHdd = plcfHdd;
                MarkerPositions = markerPositions;
                FormattedRuns = formattedRuns;
                FormattedParagraphs = formattedParagraphs;
                Bookmarks = bookmarks;
            }

            internal string Text { get; }

            internal byte[] PlcfHdd { get; }

            internal IReadOnlyList<int> MarkerPositions { get; }

            internal IReadOnlyList<LegacyDocWritableRun> FormattedRuns { get; }

            internal IReadOnlyList<LegacyDocWritableParagraph> FormattedParagraphs { get; }

            internal LegacyDocWritableBookmarks Bookmarks { get; }
        }

        private readonly struct LegacyDocWritableHeaderFooterStory {
            internal static LegacyDocWritableHeaderFooterStory Empty { get; } = new LegacyDocWritableHeaderFooterStory(string.Empty, Array.Empty<LegacyDocWritableRun>(), Array.Empty<LegacyDocWritableParagraph>(), LegacyDocWritableBookmarks.Empty);

            internal LegacyDocWritableHeaderFooterStory(string text, IReadOnlyList<LegacyDocWritableRun> formattedRuns, IReadOnlyList<LegacyDocWritableParagraph> formattedParagraphs, LegacyDocWritableBookmarks bookmarks) {
                Text = text;
                FormattedRuns = formattedRuns;
                FormattedParagraphs = formattedParagraphs;
                Bookmarks = bookmarks;
            }

            internal string Text { get; }

            internal IReadOnlyList<LegacyDocWritableRun> FormattedRuns { get; }

            internal IReadOnlyList<LegacyDocWritableParagraph> FormattedParagraphs { get; }

            internal LegacyDocWritableBookmarks Bookmarks { get; }

            internal static LegacyDocWritableHeaderFooterStory Plain(string text) {
                return new LegacyDocWritableHeaderFooterStory(text, Array.Empty<LegacyDocWritableRun>(), Array.Empty<LegacyDocWritableParagraph>(), LegacyDocWritableBookmarks.Empty);
            }
        }
    }
}
