using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private const int HeaderFooterSeparatorStoryCount = 6;
        private const int HeaderFooterStoriesPerSection = 6;

        private static LegacyDocWritableHeaderFooterStories BuildHeaderFooterStories(WordDocument document, MainDocumentPart mainPart) {
            if (!mainPart.HeaderParts.Any() && !mainPart.FooterParts.Any()) {
                return LegacyDocWritableHeaderFooterStories.Empty;
            }

            if (document.Sections.Count != 1) {
                throw new NotSupportedException("Native DOC saving currently supports headers and footers only for single-section documents.");
            }

            SectionProperties sectionProperties = document.Sections[0]._sectionProperties;
            string? defaultHeader = ReadDefaultHeaderStory(mainPart, sectionProperties);
            string? defaultFooter = ReadDefaultFooterStory(mainPart, sectionProperties);
            ThrowIfUnreferencedHeaderFooterContent(mainPart, sectionProperties);
            if (string.IsNullOrEmpty(defaultHeader) && string.IsNullOrEmpty(defaultFooter)) {
                return LegacyDocWritableHeaderFooterStories.Empty;
            }

            string[] stories = new string[HeaderFooterSeparatorStoryCount + HeaderFooterStoriesPerSection];
            stories[HeaderFooterSeparatorStoryCount + 1] = defaultHeader ?? string.Empty;
            stories[HeaderFooterSeparatorStoryCount + 3] = defaultFooter ?? string.Empty;

            var text = new StringBuilder();
            var characterPositions = new List<int>(stories.Length + 2);
            foreach (string story in stories) {
                characterPositions.Add(text.Length);
                text.Append(story);
            }

            characterPositions.Add(text.Length);
            characterPositions.Add(text.Length);
            byte[] plcfHdd = new byte[characterPositions.Count * 4];
            for (int index = 0; index < characterPositions.Count; index++) {
                WriteInt32(plcfHdd, index * 4, characterPositions[index]);
            }

            return new LegacyDocWritableHeaderFooterStories(text.ToString(), plcfHdd);
        }

        private static string? ReadDefaultHeaderStory(MainDocumentPart mainPart, SectionProperties sectionProperties) {
            HeaderReference[] references = sectionProperties.Elements<HeaderReference>().ToArray();
            if (references.Length == 0) {
                return null;
            }

            HeaderReference reference = GetSingleDefaultReference(references, "header");
            HeaderPart headerPart = GetReferencedPart<HeaderPart>(mainPart, reference.Id?.Value, "header");
            return ReadSimpleHeaderFooterStory(headerPart.Header, "header");
        }

        private static string? ReadDefaultFooterStory(MainDocumentPart mainPart, SectionProperties sectionProperties) {
            FooterReference[] references = sectionProperties.Elements<FooterReference>().ToArray();
            if (references.Length == 0) {
                return null;
            }

            FooterReference reference = GetSingleDefaultReference(references, "footer");
            FooterPart footerPart = GetReferencedPart<FooterPart>(mainPart, reference.Id?.Value, "footer");
            return ReadSimpleHeaderFooterStory(footerPart.Footer, "footer");
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

        private static TReference GetSingleDefaultReference<TReference>(IReadOnlyList<TReference> references, string kind)
            where TReference : HeaderFooterReferenceType {
            TReference? defaultReference = default;
            foreach (TReference reference in references) {
                HeaderFooterValues type = reference.Type?.Value ?? HeaderFooterValues.Default;
                if (type != HeaderFooterValues.Default) {
                    throw new NotSupportedException($"Native DOC saving currently supports only default {kind}s. First-page and even-page {kind}s are not supported yet.");
                }

                if (defaultReference != null) {
                    throw new NotSupportedException($"Native DOC saving cannot write multiple default {kind} references in one section.");
                }

                defaultReference = reference;
            }

            return defaultReference ?? throw new NotSupportedException($"Native DOC saving cannot write an empty {kind} reference list.");
        }

        private static string? ReadSimpleHeaderFooterStory(OpenXmlCompositeElement? container, string kind) {
            if (container == null || !container.HasChildren) {
                return null;
            }

            var paragraphs = new List<string>();
            foreach (OpenXmlElement child in container.ChildElements) {
                if (child is not Paragraph paragraph) {
                    throw new NotSupportedException($"Native DOC saving currently supports only text paragraphs in {kind}s. Unsupported {kind} element: {child.LocalName}.");
                }

                paragraphs.Add(ReadSimpleHeaderFooterParagraph(paragraph, kind));
            }

            bool hasVisibleText = paragraphs.Any(paragraph => paragraph.Length > 0);
            if (!hasVisibleText) {
                return null;
            }

            if (paragraphs.Any(paragraph => paragraph.Length == 0)) {
                throw new NotSupportedException($"Native DOC saving currently supports only non-empty text paragraphs in {kind}s when the {kind} contains visible text.");
            }

            return string.Join("\r", paragraphs) + "\r\r";
        }

        private static string ReadSimpleHeaderFooterParagraph(Paragraph paragraph, string kind) {
            var text = new StringBuilder();
            foreach (OpenXmlElement child in paragraph.ChildElements) {
                switch (child) {
                    case ParagraphProperties paragraphProperties:
                        ThrowIfUnsupportedHeaderFooterParagraphProperties(paragraphProperties, kind);
                        break;
                    case Run run:
                        AppendPlainHeaderFooterRun(text, run, kind);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving currently supports only plain text runs in {kind}s. Unsupported {kind} paragraph element: {child.LocalName}.");
                }
            }

            return text.ToString();
        }

        private static void ThrowIfUnsupportedHeaderFooterParagraphProperties(ParagraphProperties paragraphProperties, string kind) {
            if (paragraphProperties.ChildElements.Count == 0) {
                return;
            }

            throw new NotSupportedException($"Native DOC saving currently supports only unformatted text paragraphs in {kind}s.");
        }

        private static void AppendPlainHeaderFooterRun(StringBuilder text, Run run, string kind) {
            foreach (OpenXmlElement child in run.ChildElements) {
                switch (child) {
                    case RunProperties runProperties:
                        if (runProperties.ChildElements.Count > 0) {
                            throw new NotSupportedException($"Native DOC saving currently supports only unformatted text runs in {kind}s.");
                        }

                        break;
                    case Text textNode:
                        text.Append(textNode.Text);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving currently supports only plain text in {kind}s. Unsupported {kind} run element: {child.LocalName}.");
                }
            }
        }

        private static void ThrowIfUnreferencedHeaderFooterContent(MainDocumentPart mainPart, SectionProperties sectionProperties) {
            var referencedIds = new HashSet<string>(StringComparer.Ordinal);
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
            internal static LegacyDocWritableHeaderFooterStories Empty { get; } = new LegacyDocWritableHeaderFooterStories(string.Empty, Array.Empty<byte>());

            internal LegacyDocWritableHeaderFooterStories(string text, byte[] plcfHdd) {
                Text = text;
                PlcfHdd = plcfHdd;
            }

            internal string Text { get; }

            internal byte[] PlcfHdd { get; }
        }
    }
}
