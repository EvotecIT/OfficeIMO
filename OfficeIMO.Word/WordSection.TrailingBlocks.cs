using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordSection {
        // Last paragraph found by TryGetLastBlocks. It is only a starting point: every walk checks
        // that it still belongs to this section, so edits made elsewhere cannot make it wrong.
        private Paragraph? _lastParagraphHint;

        /// <summary>
        /// Finds the last body table or paragraph of this section and its last paragraph. The search
        /// continues forward from the previous answer, because Open XML siblings can only be walked
        /// forward cheaply and documents are mostly built by appending.
        /// </summary>
        /// <returns><c>false</c> when the body layout cannot be matched to this section cheaply.</returns>
        internal bool TryGetLastBlocks(out OpenXmlElement? lastBlock, out Paragraph? lastParagraph) {
            lastBlock = null;
            lastParagraph = null;
            if (!TryGetBodyRegion(out Body? body, out OpenXmlElement? regionStart, out Paragraph? regionEnd)) {
                return false;
            }

            Paragraph? hint = _lastParagraphHint;
            bool found = hint != null &&
                         ReferenceEquals(hint.Parent, body) &&
                         TryFindLastBlocks(hint, regionEnd, out lastBlock, out lastParagraph);
            if (!found && !TryFindLastBlocks(regionStart, regionEnd, out lastBlock, out lastParagraph)) {
                _lastParagraphHint = null;
                return false;
            }

            _lastParagraphHint = lastParagraph;
            return true;
        }

        private static bool TryFindLastBlocks(OpenXmlElement? start, Paragraph? regionEnd, out OpenXmlElement? lastBlock, out Paragraph? lastParagraph) {
            lastBlock = null;
            lastParagraph = null;
            for (OpenXmlElement? current = start; ; current = current.NextSibling()) {
                if (current == null) {
                    return regionEnd == null;
                }
                if (ReferenceEquals(current, regionEnd)) {
                    return true;
                }

                if (current is Paragraph paragraph) {
                    if (IsSectionBoundaryParagraph(paragraph)) {
                        return false;
                    }
                    lastBlock = paragraph;
                    lastParagraph = paragraph;
                } else if (current is Table) {
                    lastBlock = current;
                }
            }
        }

        /// <summary>
        /// Collects the body paragraphs that follow <paramref name="lastParagraph"/> and the body tables
        /// that follow <paramref name="lastBlock"/> in this section, in document order. The anchors come
        /// from <see cref="TryGetLastBlocks"/>; a <c>null</c> anchor collects from the start of the section.
        /// </summary>
        /// <returns><c>false</c> when an anchor is no longer in this section or the layout cannot be matched.</returns>
        internal bool TryGetBlocksAfter(OpenXmlElement? lastBlock, Paragraph? lastParagraph, List<Paragraph> paragraphs, List<Table> tables) {
            paragraphs.Clear();
            tables.Clear();
            if (!TryGetBodyRegion(out Body? body, out OpenXmlElement? regionStart, out Paragraph? regionEnd)) {
                return false;
            }

            OpenXmlElement? current = regionStart;
            if (lastParagraph != null) {
                if (!ReferenceEquals(lastParagraph.Parent, body)) {
                    return false;
                }
                current = lastParagraph;
            }

            bool afterLastBlock = lastBlock == null;
            for (; ; current = current.NextSibling()) {
                if (current == null) {
                    if (regionEnd != null) {
                        return false;
                    }
                    break;
                }
                if (ReferenceEquals(current, regionEnd)) {
                    break;
                }

                if (current is Paragraph paragraph) {
                    if (IsSectionBoundaryParagraph(paragraph)) {
                        return false;
                    }
                    if (!ReferenceEquals(paragraph, lastParagraph)) {
                        paragraphs.Add(paragraph);
                    }
                } else if (current is Table table && afterLastBlock) {
                    tables.Add(table);
                }

                if (ReferenceEquals(current, lastBlock)) {
                    afterLastBlock = true;
                }
            }

            return afterLastBlock;
        }

        /// <summary>
        /// Resolves the first body element of this section and the break paragraph that ends it, or
        /// <c>null</c> for the final section, which runs to the end of the body. A non-final section is
        /// only matched when its break paragraph carries nothing but the section break, because content is
        /// added before that paragraph and every other paragraph of the section precedes it.
        /// </summary>
        private bool TryGetBodyRegion(out Body? body, out OpenXmlElement? regionStart, out Paragraph? regionEnd) {
            body = _document._document?.Body;
            regionStart = null;
            regionEnd = null;
            List<WordSection> sections = _document.Sections;
            int index = sections.IndexOf(this);
            if (body == null || index < 0) {
                return false;
            }

            if (index < sections.Count - 1) {
                regionEnd = GetBodyBoundaryParagraph(body, _sectionProperties);
                if (regionEnd == null || !IsPureSectionBreakParagraph(regionEnd)) {
                    return false;
                }
            }

            if (index == 0) {
                regionStart = body.FirstChild;
            } else {
                Paragraph? previousBoundary = GetBodyBoundaryParagraph(body, sections[index - 1]._sectionProperties);
                if (previousBoundary == null) {
                    return false;
                }
                regionStart = previousBoundary.NextSibling();
            }

            return true;
        }

        private static Paragraph? GetBodyBoundaryParagraph(Body body, SectionProperties sectionProperties) =>
            sectionProperties.Parent is ParagraphProperties properties &&
            properties.Parent is Paragraph paragraph &&
            ReferenceEquals(paragraph.Parent, body)
                ? paragraph
                : null;
    }
}
