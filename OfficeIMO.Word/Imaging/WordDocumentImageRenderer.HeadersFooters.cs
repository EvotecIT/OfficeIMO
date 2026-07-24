using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;

namespace OfficeIMO.Word {
    internal static partial class WordDocumentImageRenderer {
        private const double DefaultHeaderFooterLineHeightPoints = 18D;

        private static WordHeaderFooterPageFrame AddSupportedHeaderFooterContent(WordSection section, OfficeDrawing drawing, List<OfficeImageExportDiagnostic> diagnostics, int pageIndex, int sectionIndex, int sectionPageNumberStart, int sectionPageIndex, int totalPageCount, int sectionPageCount) {
            WordHeaderFooterPageFrame frame = CreateHeaderFooterPageFrame(section, drawing, pageIndex, sectionIndex, sectionPageNumberStart, sectionPageIndex, totalPageCount, sectionPageCount);
            if (frame.Header != null) {
                WordImageFlowContext context = CreateFlowContext(
                    drawing,
                    frame.ContentLeft,
                    frame.HeaderTop,
                    frame.ContentWidth,
                    frame.HeaderRenderBottom,
                    "unsupported-word-header-overflow",
                    "Stopped rendering Word header content because it does not fit in the page header band.",
                    targetPageIndex: pageIndex,
                    initialPageIndex: pageIndex,
                    resolveDynamicPageFields: true,
                    totalPageCount: totalPageCount,
                    sectionNumber: sectionIndex + 1,
                    sectionPageCount: sectionPageCount,
                    pageNumberValue: frame.PageNumberValue,
                    pageNumberText: frame.PageNumberText);
                AddHeaderFooterContent(frame.Header, context, diagnostics, "header");
            }

            if (frame.Footer != null) {
                WordImageFlowContext context = CreateFlowContext(
                    drawing,
                    frame.ContentLeft,
                    frame.FooterTop,
                    frame.ContentWidth,
                    drawing.Height,
                    "unsupported-word-footer-overflow",
                    "Stopped rendering Word footer content because it does not fit in the page footer band.",
                    targetPageIndex: pageIndex,
                    initialPageIndex: pageIndex,
                    resolveDynamicPageFields: true,
                    totalPageCount: totalPageCount,
                    sectionNumber: sectionIndex + 1,
                    sectionPageCount: sectionPageCount,
                    pageNumberValue: frame.PageNumberValue,
                    pageNumberText: frame.PageNumberText);
                AddHeaderFooterContent(frame.Footer, context, diagnostics, "footer");
            }

            return frame;
        }

        private static WordHeaderFooterPageFrame CreateHeaderFooterPageFrame(WordSection section, OfficeDrawing drawing, int pageIndex, int sectionIndex, int sectionPageNumberStart, int sectionPageIndex, int totalPageCount, int sectionPageCount) {
            WordMargins margins = section.Margins;
            double left = ToPoints(margins.Left?.Value, DefaultMarginPoints);
            double right = ToPoints(margins.Right?.Value, DefaultMarginPoints);
            double topMargin = ToPoints(margins.Top, DefaultMarginPoints);
            double bottomMargin = ToPoints(margins.Bottom, DefaultMarginPoints);
            double contentWidth = Math.Max(1D, drawing.Width - left - right);
            double bodyTop = topMargin;
            double bodyBottom = Math.Max(bodyTop, drawing.Height - bottomMargin);

            (int pageNumberValue, string pageNumberText) = ResolveSectionPageNumber(section, sectionPageNumberStart, sectionPageIndex);
            WordHeaderFooter? header = SelectPageHeader(section, sectionPageIndex, pageNumberValue);
            double headerTop = 0D;
            double headerRenderBottom = 0D;
            if (header != null) {
                double headerDistance = ToPoints(margins.HeaderDistance?.Value, DefaultMarginPoints / 2D);
                double headerHeight = EstimateHeaderFooterContentHeight(header, drawing.Width, left, contentWidth, pageIndex, sectionIndex, sectionPageCount, pageNumberValue, pageNumberText, totalPageCount);
                headerTop = Math.Max(0D, Math.Min(headerDistance, topMargin) - (DefaultHeaderFooterLineHeightPoints / 2D));
                double headerContentBottom = Math.Min(drawing.Height, headerTop + headerHeight);
                headerRenderBottom = Math.Min(drawing.Height, Math.Max(headerContentBottom, topMargin + DefaultHeaderFooterLineHeightPoints + ParagraphGapPoints));
                if (headerContentBottom > topMargin) {
                    bodyTop = Math.Min(drawing.Height, Math.Max(bodyTop, headerContentBottom + ParagraphGapPoints));
                }
            }

            WordHeaderFooter? footer = SelectPageFooter(section, sectionPageIndex, pageNumberValue);
            double footerTop = drawing.Height;
            double footerRenderBottom = drawing.Height;
            if (footer != null) {
                double footerDistance = ToPoints(margins.FooterDistance?.Value, DefaultMarginPoints / 2D);
                double footerHeight = EstimateHeaderFooterContentHeight(footer, drawing.Width, left, contentWidth, pageIndex, sectionIndex, sectionPageCount, pageNumberValue, pageNumberText, totalPageCount);
                double footerTopFromDistance = drawing.Height - footerDistance - footerHeight;
                footerTop = Math.Min(Math.Max(0D, footerTopFromDistance), Math.Max(0D, drawing.Height - footerHeight));
                footerRenderBottom = Math.Min(drawing.Height, footerTop + footerHeight + ParagraphGapPoints);
                if (footerTop < bodyBottom) {
                    bodyBottom = Math.Max(bodyTop, Math.Min(bodyBottom, footerTop - ParagraphGapPoints));
                }
            }

            bodyBottom = Math.Max(bodyTop, bodyBottom);
            return new WordHeaderFooterPageFrame(
                header,
                footer,
                left,
                contentWidth,
                bodyTop,
                bodyBottom,
                headerTop,
                headerRenderBottom,
                footerTop,
                footerRenderBottom,
                pageNumberValue,
                pageNumberText);
        }

        private static Func<int, WordImageBodyFrame> CreateBodyFrameProvider(
            WordSection section,
            OfficeDrawing drawing,
            int sectionIndex,
            int sectionPageNumberStart,
            int totalPageCount,
            int sectionPageCount,
            int knownSectionPageIndex,
            WordHeaderFooterPageFrame? knownFrame = null) =>
            sectionPageIndex => {
                int normalizedSectionPageIndex = Math.Max(0, sectionPageIndex);
                if (knownFrame.HasValue && normalizedSectionPageIndex == knownSectionPageIndex) {
                    return knownFrame.Value.BodyFrame;
                }

                WordHeaderFooterPageFrame frame = CreateHeaderFooterPageFrame(
                    section,
                    drawing,
                    normalizedSectionPageIndex,
                    sectionIndex,
                    sectionPageNumberStart,
                    normalizedSectionPageIndex,
                    totalPageCount,
                    sectionPageCount);
                return frame.BodyFrame;
            };

        private static (int Value, string Text) ResolveSectionPageNumber(WordSection section, int sectionPageNumberStart, int sectionPageIndex) {
            PageNumberType? pageNumberType = section._sectionProperties.GetFirstChild<PageNumberType>();
            int start = pageNumberType?.Start?.Value ?? sectionPageNumberStart;
            int value = Math.Max(1, start + Math.Max(0, sectionPageIndex));
            return (value, FormatPageNumber(value, pageNumberType?.Format?.Value));
        }

        private static WordHeaderFooter? SelectPageHeader(WordSection section, int sectionPageIndex, int pageNumberValue) {
            if (sectionPageIndex == 0 && section.DifferentFirstPage && section.Header.First != null) {
                return section.Header.First;
            }

            if (IsEvenPageNumber(pageNumberValue) &&
                section.DifferentOddAndEvenPages &&
                section.Header.Even != null) {
                return section.Header.Even;
            }

            return section.Header.Default;
        }

        private static WordHeaderFooter? SelectPageFooter(WordSection section, int sectionPageIndex, int pageNumberValue) {
            if (sectionPageIndex == 0 && section.DifferentFirstPage && section.Footer.First != null) {
                return section.Footer.First;
            }

            if (IsEvenPageNumber(pageNumberValue) &&
                section.DifferentOddAndEvenPages &&
                section.Footer.Even != null) {
                return section.Footer.Even;
            }

            return section.Footer.Default;
        }

        private static bool IsEvenPageNumber(int pageNumber) =>
            (pageNumber % 2) == 0;

        private static double EstimateHeaderFooterContentHeight(
            WordHeaderFooter headerFooter,
            double pageWidth,
            double left,
            double contentWidth,
            int pageIndex,
            int sectionIndex,
            int sectionPageCount,
            int pageNumberValue,
            string pageNumberText,
            int totalPageCount) {
            var measurementDrawing = new OfficeDrawing(Math.Max(1D, pageWidth), double.MaxValue);
            WordImageFlowContext measurementContext = CreateFlowContext(
                measurementDrawing,
                left,
                0D,
                contentWidth,
                double.MaxValue,
                "unsupported-word-header-footer-measurement-overflow",
                "Skipped Word header/footer measurement because content does not fit within the measurement frame.",
                targetPageIndex: pageIndex,
                initialPageIndex: pageIndex,
                resolveDynamicPageFields: true,
                totalPageCount: totalPageCount,
                sectionNumber: sectionIndex + 1,
                sectionPageCount: sectionPageCount,
                pageNumberValue: pageNumberValue,
                pageNumberText: pageNumberText);
            using (measurementDrawing.DeferBehindContentOrdering()) {
                AddHeaderFooterContent(headerFooter, measurementContext, new List<OfficeImageExportDiagnostic>(), "header-footer");
            }
            return Math.Max(DefaultHeaderFooterLineHeightPoints, measurementContext.Y);
        }

        private static void AddHeaderFooterContent(WordHeaderFooter headerFooter, WordImageFlowContext context, List<OfficeImageExportDiagnostic> diagnostics, string kind) {
            WordDocument document = headerFooter.Document;
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers = DocumentTraversal.BuildListMarkers(document);
            foreach (OpenXmlElement element in headerFooter.ChildElements) {
                bool added = AddHeaderFooterElementContent(document, element, context, diagnostics, listMarkers, kind);

                if (context.StoppedForPagination) {
                    break;
                }

                if (!added && element is Paragraph) {
                    context.Y += ParagraphGapPoints;
                }
            }
        }

        private readonly struct WordHeaderFooterPageFrame {
            internal WordHeaderFooterPageFrame(
                WordHeaderFooter? header,
                WordHeaderFooter? footer,
                double contentLeft,
                double contentWidth,
                double bodyTop,
                double bodyBottom,
                double headerTop,
                double headerRenderBottom,
                double footerTop,
                double footerRenderBottom,
                int pageNumberValue,
                string pageNumberText) {
                Header = header;
                Footer = footer;
                ContentLeft = contentLeft;
                ContentWidth = contentWidth;
                BodyTop = bodyTop;
                BodyBottom = bodyBottom;
                HeaderTop = headerTop;
                HeaderRenderBottom = headerRenderBottom;
                FooterTop = footerTop;
                FooterRenderBottom = footerRenderBottom;
                PageNumberValue = pageNumberValue;
                PageNumberText = pageNumberText;
            }

            internal WordHeaderFooter? Header { get; }

            internal WordHeaderFooter? Footer { get; }

            internal double ContentLeft { get; }

            internal double ContentWidth { get; }

            internal double BodyTop { get; }

            internal double BodyBottom { get; }

            internal WordImageBodyFrame BodyFrame => new WordImageBodyFrame(BodyTop, BodyBottom);

            internal double HeaderTop { get; }

            internal double HeaderRenderBottom { get; }

            internal double FooterTop { get; }

            internal double FooterRenderBottom { get; }

            internal int PageNumberValue { get; }

            internal string PageNumberText { get; }
        }

        private readonly struct WordImageBodyFrame {
            internal WordImageBodyFrame(double top, double bottom) {
                Top = top;
                Bottom = bottom;
            }

            internal double Top { get; }

            internal double Bottom { get; }
        }
    }
}
