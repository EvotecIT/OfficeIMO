using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;

namespace OfficeIMO.Word {
    internal static partial class WordDocumentImageRenderer {
        private const double DefaultHeaderFooterLineHeightPoints = 18D;

        private static void AddSupportedHeaderFooterContent(WordDocument document, OfficeDrawing drawing, List<OfficeImageExportDiagnostic> diagnostics) {
            WordSection? section = document.Sections.FirstOrDefault();
            if (section == null) {
                return;
            }

            WordMargins margins = section.Margins;
            double left = ToPoints(margins.Left?.Value, DefaultMarginPoints);
            double right = ToPoints(margins.Right?.Value, DefaultMarginPoints);
            double topMargin = ToPoints(margins.Top, DefaultMarginPoints);
            double bottomMargin = ToPoints(margins.Bottom, DefaultMarginPoints);
            double contentWidth = Math.Max(1D, drawing.Width - left - right);

            WordHeaderFooter? header = SelectFirstPageHeader(section);
            if (header != null) {
                double headerDistance = ToPoints(margins.HeaderDistance?.Value, DefaultMarginPoints / 2D);
                double headerTop = Math.Max(0D, Math.Min(headerDistance, topMargin) - (DefaultHeaderFooterLineHeightPoints / 2D));
                double headerBottom = Math.Max(headerTop + DefaultHeaderFooterLineHeightPoints, topMargin + DefaultHeaderFooterLineHeightPoints + ParagraphGapPoints);
                WordImageFlowContext context = CreateFlowContext(
                    drawing,
                    left,
                    headerTop,
                    contentWidth,
                    headerBottom,
                    "unsupported-word-header-overflow",
                    "Stopped rendering Word header content because it does not fit in the first-page header band.");
                AddHeaderFooterContent(header, context, diagnostics, "header");
            }

            WordHeaderFooter? footer = SelectFirstPageFooter(section);
            if (footer != null) {
                double footerDistance = ToPoints(margins.FooterDistance?.Value, DefaultMarginPoints / 2D);
                double bodyBottom = Math.Max(topMargin, drawing.Height - bottomMargin);
                double footerTopFromDistance = drawing.Height - footerDistance - DefaultHeaderFooterLineHeightPoints;
                double footerTop = Math.Max(bodyBottom + ParagraphGapPoints, footerTopFromDistance);
                footerTop = Math.Min(Math.Max(0D, footerTop), Math.Max(0D, drawing.Height - DefaultHeaderFooterLineHeightPoints));
                WordImageFlowContext context = CreateFlowContext(
                    drawing,
                    left,
                    footerTop,
                    contentWidth,
                    drawing.Height - 1D,
                    "unsupported-word-footer-overflow",
                    "Stopped rendering Word footer content because it does not fit in the first-page footer band.");
                AddHeaderFooterContent(footer, context, diagnostics, "footer");
            }
        }

        private static WordHeaderFooter? SelectFirstPageHeader(WordSection section) {
            if (section.DifferentFirstPage) {
                return section.Header.First;
            }

            return section.Header.Default;
        }

        private static WordHeaderFooter? SelectFirstPageFooter(WordSection section) {
            if (section.DifferentFirstPage) {
                return section.Footer.First;
            }

            return section.Footer.Default;
        }

        private static void AddHeaderFooterContent(WordHeaderFooter headerFooter, WordImageFlowContext context, List<OfficeImageExportDiagnostic> diagnostics, string kind) {
            WordDocument document = headerFooter.Document;
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers = DocumentTraversal.BuildListMarkers(document);
            foreach (OpenXmlElement element in headerFooter.ChildElements) {
                bool added = false;
                if (element is Paragraph paragraph) {
                    added = AddParagraphContent(document, paragraph, context, diagnostics, listMarkers);
                } else if (element is Table table) {
                    added = AddTable(new WordTable(document, table), context, diagnostics, listMarkers);
                } else {
                    AddDiagnostic(diagnostics, "unsupported-word-" + kind + "-element", "Skipped a Word " + kind + " element that is not yet projected through OfficeIMO.Drawing.", element.GetType().Name);
                }

                if (context.StoppedForPagination) {
                    break;
                }

                if (!added && element is Paragraph) {
                    context.Y += ParagraphGapPoints;
                }
            }
        }
    }
}
