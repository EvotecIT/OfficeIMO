using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using System.Threading;

namespace OfficeIMO.Word {
    internal static partial class WordDocumentImageRenderer {
        internal static int EstimatePageCount(WordDocument document) =>
            Math.Max(1, EstimateSectionPageCounts(document).Sum());

        private static IReadOnlyList<int> EstimateSectionPageCounts(
            WordDocument document,
            CancellationToken cancellationToken = default,
            Action<WordImageCancellationCheckpoint>? cancellationCheckpoint = null) {
            cancellationToken.ThrowIfCancellationRequested();
            int sectionCount = Math.Max(1, document.Sections.Count);
            int[] pageCounts = new int[sectionCount];
            int sectionGroupStart = 0;
            int sectionIndex = 0;
            int firstPageInSection = 0;
            var sectionElements = new List<OpenXmlElement>();

            foreach (OpenXmlElement element in document.BodyRoot.ChildElements) {
                cancellationToken.ThrowIfCancellationRequested();
                if (element is SectionProperties) {
                    continue;
                }

                sectionElements.Add(element);
                if (HasSectionBoundary(element) && sectionIndex < sectionCount - 1) {
                    WordSection currentSection = document.Sections[sectionIndex];
                    WordSection nextSection = document.Sections[sectionIndex + 1];
                    SectionProperties? boundaryProperties = GetSectionBoundaryProperties(element);
                    if (CanMergeImageSectionOnSamePage(boundaryProperties, currentSection, nextSection)) {
                        sectionIndex++;
                        continue;
                    }

                    WordSection groupSection = document.Sections[sectionGroupStart];
                    int contentPages = EstimateSectionContentPageCount(
                        document,
                        groupSection,
                        sectionElements,
                        cancellationToken,
                        cancellationCheckpoint);
                    int breakPages = CountSectionBreakPageAdvance(firstPageInSection + contentPages - 1, boundaryProperties);
                    pageCounts[sectionGroupStart] = Math.Max(1, contentPages - 1 + breakPages);
                    firstPageInSection += pageCounts[sectionGroupStart];
                    sectionElements.Clear();
                    sectionIndex++;
                    sectionGroupStart = sectionIndex;
                }
            }

            if (sectionGroupStart < sectionCount) {
                pageCounts[sectionGroupStart] = EstimateSectionContentPageCount(
                    document,
                    document.Sections[sectionGroupStart],
                    sectionElements,
                    cancellationToken,
                    cancellationCheckpoint);
            }

            return pageCounts;
        }

        private static IReadOnlyList<OpenXmlElement> GetSectionBodyElements(WordDocument document, int targetSectionIndex) {
            return GetSectionBodyElementEntries(document, targetSectionIndex)
                .Select(entry => entry.Element)
                .ToList();
        }

        private static IReadOnlyList<WordSectionBodyElement> GetSectionBodyElementEntries(
            WordDocument document,
            int targetSectionIndex,
            CancellationToken cancellationToken = default) {
            int sectionCount = Math.Max(1, document.Sections.Count);
            int normalizedTarget = Math.Min(Math.Max(0, targetSectionIndex), sectionCount - 1);
            int sectionIndex = 0;
            var sectionElements = new List<WordSectionBodyElement>();

            foreach (OpenXmlElement element in document.BodyRoot.ChildElements) {
                cancellationToken.ThrowIfCancellationRequested();
                if (element is SectionProperties) {
                    continue;
                }

                sectionElements.Add(new WordSectionBodyElement(element, sectionIndex));
                if (HasSectionBoundary(element) && sectionIndex < sectionCount - 1) {
                    if (sectionIndex >= normalizedTarget &&
                        CanMergeImageSectionOnSamePage(GetSectionBoundaryProperties(element), document.Sections[sectionIndex], document.Sections[sectionIndex + 1])) {
                        sectionIndex++;
                        continue;
                    }

                    if (sectionIndex >= normalizedTarget) {
                        return sectionElements;
                    }

                    sectionElements = new List<WordSectionBodyElement>();
                    sectionIndex++;
                }
            }

            return sectionIndex >= normalizedTarget
                ? sectionElements
                : new List<WordSectionBodyElement>();
        }

        private readonly struct WordSectionBodyElement {
            internal WordSectionBodyElement(OpenXmlElement element, int sectionIndex) {
                Element = element;
                SectionIndex = sectionIndex;
            }

            internal OpenXmlElement Element { get; }

            internal int SectionIndex { get; }
        }

        private static int EstimateSectionContentPageCount(
            WordDocument document,
            WordSection section,
            IReadOnlyList<OpenXmlElement> sectionElements,
            CancellationToken cancellationToken,
            Action<WordImageCancellationCheckpoint>? cancellationCheckpoint) {
            cancellationToken.ThrowIfCancellationRequested();
            (double width, double height) = GetPageSizePoints(section);
            var drawing = new OfficeDrawing(width, height);
            WordHeaderFooterPageFrame headerFooterFrame = CreateHeaderFooterPageFrame(section, drawing, 0, document.Sections.IndexOf(section), 1, 0, 1, 1);
            WordImageFlowContext context = CreateFlowContext(
                section,
                drawing,
                int.MaxValue,
                contentTop: headerFooterFrame.BodyTop,
                contentBottom: headerFooterFrame.BodyBottom,
                bodyFrameProvider: CreateBodyFrameProvider(section, drawing, document.Sections.IndexOf(section), 1, 1, 1, 0, headerFooterFrame),
                cancellationToken: cancellationToken,
                cancellationCheckpoint: cancellationCheckpoint);
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers = DocumentTraversal.BuildListMarkers(document);
            var diagnostics = new List<OfficeImageExportDiagnostic>();

            for (int index = 0; index < sectionElements.Count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                OpenXmlElement element = sectionElements[index];
                TryAdvanceForKeepWithNext(document, sectionElements, index, context, listMarkers);
                if (context.StoppedForPagination) {
                    break;
                }

                bool added = AddBodyElementContentForSectionPageEstimate(document, element, context, diagnostics, listMarkers);
                if (context.StoppedForPagination) {
                    break;
                }

                if (!added && element is Paragraph) {
                    context.ClearParagraphSpacingState();
                    context.Y += ParagraphGapPoints;
                }
            }

            return Math.Max(1, context.PageIndex + 1);
        }

        private static bool AddBodyElementContentForSectionPageEstimate(
            WordDocument document,
            OpenXmlElement element,
            WordImageFlowContext context,
            List<OfficeImageExportDiagnostic> diagnostics,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers) {
            if (element is Paragraph paragraph) {
                if (ResolvePageBreakBefore(document, paragraph)) {
                    AdvanceForPageBreakBefore(context);
                }

                bool added = AddParagraphContent(document, paragraph, context, diagnostics, listMarkers);
                if (IsNextColumnSectionBreak(paragraph.ParagraphProperties?.SectionProperties)) {
                    context.AdvanceColumnOrPage();
                    return true;
                }

                return added;
            }

            return AddBodyElementContent(document, element, context, diagnostics, listMarkers);
        }

        private static SectionProperties? GetSectionBoundaryProperties(OpenXmlElement element) =>
            element is Paragraph paragraph ? paragraph.ParagraphProperties?.SectionProperties : null;

        private static bool CanMergeImageSectionOnSamePage(SectionProperties? boundaryProperties, WordSection previous, WordSection current) {
            SectionMarkValues? sectionMark = ResolveSectionMark(boundaryProperties);
            if (sectionMark == SectionMarkValues.Continuous) {
                return CanMergeImageContinuousSection(previous, current);
            }

            return sectionMark == SectionMarkValues.NextColumn &&
                CanMergeImageNextColumnSection(previous, current);
        }

        private static bool CanMergeImageContinuousSection(WordSection previous, WordSection current) {
            if (GetSectionColumnCount(previous) > 1 || GetSectionColumnCount(current) > 1) {
                return false;
            }

            return SectionPageSetupEquivalent(previous, current) &&
                SectionChildElementsEquivalent<HeaderReference>(previous, current) &&
                SectionChildElementsEquivalent<FooterReference>(previous, current) &&
                SectionChildElementEquivalent<PageNumberType>(previous, current);
        }

        private static bool CanMergeImageNextColumnSection(WordSection previous, WordSection current) {
            if (GetSectionColumnCount(previous) <= 1) {
                return false;
            }

            return SectionChildElementEquivalent<PageSize>(previous, current) &&
                SectionChildElementEquivalent<PageMargin>(previous, current) &&
                SectionChildElementEquivalent<Columns>(previous, current) &&
                SectionChildElementsEquivalent<HeaderReference>(previous, current) &&
                SectionChildElementsEquivalent<FooterReference>(previous, current) &&
                SectionChildElementEquivalent<PageNumberType>(previous, current);
        }

        private static bool SectionPageSetupEquivalent(WordSection first, WordSection second) {
            (double firstWidth, double firstHeight) = GetPageSizePoints(first);
            (double secondWidth, double secondHeight) = GetPageSizePoints(second);
            (double firstLeft, double firstRight, double firstTop, double firstBottom, double firstHeader, double firstFooter, double firstGutter) = GetSectionMarginPoints(first);
            (double secondLeft, double secondRight, double secondTop, double secondBottom, double secondHeader, double secondFooter, double secondGutter) = GetSectionMarginPoints(second);

            return PointsEquivalent(firstWidth, secondWidth) &&
                PointsEquivalent(firstHeight, secondHeight) &&
                PointsEquivalent(firstLeft, secondLeft) &&
                PointsEquivalent(firstRight, secondRight) &&
                PointsEquivalent(firstTop, secondTop) &&
                PointsEquivalent(firstBottom, secondBottom) &&
                PointsEquivalent(firstHeader, secondHeader) &&
                PointsEquivalent(firstFooter, secondFooter) &&
                PointsEquivalent(firstGutter, secondGutter);
        }

        private static (double Left, double Right, double Top, double Bottom, double Header, double Footer, double Gutter) GetSectionMarginPoints(WordSection section) {
            WordMargins margins = section.Margins;
            return (
                ToPoints(margins.Left?.Value, DefaultMarginPoints),
                ToPoints(margins.Right?.Value, DefaultMarginPoints),
                ToPoints(margins.Top, DefaultMarginPoints),
                ToPoints(margins.Bottom, DefaultMarginPoints),
                ToPoints(margins.HeaderDistance?.Value, DefaultMarginPoints / 2D),
                ToPoints(margins.FooterDistance?.Value, DefaultMarginPoints / 2D),
                ToPoints(margins.Gutter?.Value, 0D));
        }

        private static bool PointsEquivalent(double first, double second) =>
            Math.Abs(first - second) < 0.01D;

        private static bool SectionChildElementsEquivalent<TElement>(WordSection first, WordSection second)
            where TElement : OpenXmlElement =>
            string.Join("|", first._sectionProperties.Elements<TElement>().Select(element => element.OuterXml).OrderBy(value => value)) ==
            string.Join("|", second._sectionProperties.Elements<TElement>().Select(element => element.OuterXml).OrderBy(value => value));

        private static bool SectionChildElementEquivalent<TElement>(WordSection first, WordSection second)
            where TElement : OpenXmlElement =>
            (first._sectionProperties.GetFirstChild<TElement>()?.OuterXml ?? string.Empty) ==
            (second._sectionProperties.GetFirstChild<TElement>()?.OuterXml ?? string.Empty);

        private static int CountExplicitPageAdvances(WordDocument document, OpenXmlElement element) {
            int count = 0;
            foreach (Paragraph paragraph in EnumerateParagraphs(element)) {
                if (ResolvePageBreakBefore(document, paragraph)) {
                    count++;
                }

                count += paragraph
                    .Descendants<Break>()
                    .Count(documentBreak => documentBreak.Type?.Value == BreakValues.Page);
            }

            return count;
        }

        private static IEnumerable<Paragraph> EnumerateParagraphs(OpenXmlElement element) {
            if (element is Paragraph paragraph) {
                yield return paragraph;
            }

            foreach (Paragraph descendant in element.Descendants<Paragraph>()) {
                yield return descendant;
            }
        }

        private static bool HasSectionBoundary(OpenXmlElement element) =>
            element is Paragraph paragraph && paragraph.ParagraphProperties?.SectionProperties != null;
    }
}
