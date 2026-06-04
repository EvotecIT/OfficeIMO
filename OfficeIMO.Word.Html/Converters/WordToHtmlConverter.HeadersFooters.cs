using AngleSharp.Dom;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;
using System.Threading;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        private delegate void AppendWordParagraphHtml(IElement parent, WordParagraph paragraph);

        private delegate void AppendWordTableHtml(IElement parent, WordTable table);

        private static void AppendHeaderFooterRegions(
            IDocument htmlDoc,
            IElement parent,
            WordSection section,
            int sectionIndex,
            bool headers,
            AppendWordParagraphHtml appendParagraph,
            AppendWordTableHtml appendTable,
            WordToHtmlOptions options,
            CancellationToken cancellationToken) {
            if (!options.ExportHeadersAndFooters) {
                return;
            }

            if (headers) {
                AppendHeaderFooterRegion(htmlDoc, parent, section.Header.Default, "header", "default", sectionIndex, appendParagraph, appendTable, cancellationToken);
                AppendHeaderFooterRegion(htmlDoc, parent, section.Header.First, "header", "first", sectionIndex, appendParagraph, appendTable, cancellationToken);
                AppendHeaderFooterRegion(htmlDoc, parent, section.Header.Even, "header", "even", sectionIndex, appendParagraph, appendTable, cancellationToken);
            } else {
                AppendHeaderFooterRegion(htmlDoc, parent, section.Footer.Default, "footer", "default", sectionIndex, appendParagraph, appendTable, cancellationToken);
                AppendHeaderFooterRegion(htmlDoc, parent, section.Footer.First, "footer", "first", sectionIndex, appendParagraph, appendTable, cancellationToken);
                AppendHeaderFooterRegion(htmlDoc, parent, section.Footer.Even, "footer", "even", sectionIndex, appendParagraph, appendTable, cancellationToken);
            }
        }

        private static void AppendHeaderFooterRegion(
            IDocument htmlDoc,
            IElement parent,
            WordHeaderFooter? headerFooter,
            string tagName,
            string type,
            int sectionIndex,
            AppendWordParagraphHtml appendParagraph,
            AppendWordTableHtml appendTable,
            CancellationToken cancellationToken) {
            if (headerFooter == null || !HasRenderableHeaderFooterContent(headerFooter)) {
                return;
            }

            var element = htmlDoc.CreateElement(tagName);
            var kind = string.Equals(tagName, "header", StringComparison.OrdinalIgnoreCase) ? "header" : "footer";
            element.SetAttribute("class", $"word-{kind} word-{kind}-{type}");
            element.SetAttribute("data-section-index", sectionIndex.ToString(CultureInfo.InvariantCulture));
            element.SetAttribute("data-type", type);

            foreach (var child in headerFooter.Elements) {
                cancellationToken.ThrowIfCancellationRequested();
                if (child is WordParagraph paragraph) {
                    if (!IsRenderableHeaderFooterParagraph(paragraph)) {
                        continue;
                    }
                    appendParagraph(element, paragraph);
                } else if (child is WordTable table) {
                    appendTable(element, table);
                }
            }

            parent.AppendChild(element);
        }

        private static bool HasRenderableHeaderFooterContent(WordHeaderFooter headerFooter) {
            foreach (var element in headerFooter.Elements) {
                if (element is WordParagraph paragraph) {
                    if (IsRenderableHeaderFooterParagraph(paragraph)) {
                        return true;
                    }
                } else if (element is WordTable table && table.Rows.Count > 0) {
                    return true;
                }
            }

            return false;
        }

        private static bool IsRenderableHeaderFooterParagraph(WordParagraph paragraph) =>
            !string.IsNullOrWhiteSpace(paragraph.Text) ||
            paragraph.GetRuns().Any(run => run.IsImage || run.IsStructuredDocumentTag || run.IsCheckBox || run.IsDropDownList || run.IsComboBox || run.IsDatePicker);
    }
}
