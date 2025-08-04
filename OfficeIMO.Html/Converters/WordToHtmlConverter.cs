using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System;
using System.IO;
using System.Text;

namespace OfficeIMO.Html {
    /// <summary>
    /// Converts WordprocessingDocument content into simple HTML fragments.
    /// </summary>
    public static class WordToHtmlConverter {
        /// <summary>
        /// Converts a DOCX contained in the provided stream into HTML.
        /// </summary>
        /// <param name="docxStream">Stream containing DOCX content.</param>
        /// <param name="options">Conversion options.</param>
        /// <returns>Generated HTML string.</returns>
        public static string Convert(Stream docxStream, WordToHtmlOptions? options = null) {
            if (docxStream == null) {
                throw new ArgumentNullException(nameof(docxStream));
            }

            options ??= new WordToHtmlOptions();

            using WordprocessingDocument document = WordprocessingDocument.Open(docxStream, false);
            StringBuilder sb = new StringBuilder();
            sb.Append("<html><body>");

            foreach (var evt in WordListTraversal.Traverse(document)) {
                switch (evt.EventType) {
                    case WordListEventType.StartList:
                        string tagOpen = evt.Ordered ? "<ol>" : "<ul>";
                        if (options.PreserveListStyles) {
                            string listStyle = evt.Ordered ? "decimal" : "disc";
                            tagOpen = evt.Ordered ? $"<ol style=\"list-style-type:{listStyle}\">" : $"<ul style=\"list-style-type:{listStyle}\">";
                        }
                        sb.Append(tagOpen);
                        break;
                    case WordListEventType.EndList:
                        sb.Append(evt.Ordered ? "</ol>" : "</ul>");
                        break;
                    case WordListEventType.StartItem:
                        sb.Append("<li>");
                        AppendRuns(sb, evt.Paragraph!, options);
                        break;
                    case WordListEventType.EndItem:
                        sb.Append("</li>");
                        break;
                    case WordListEventType.Paragraph:
                        string tag = "p";
                        string? styleId = evt.Paragraph!.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                        if (styleId != null && Enum.TryParse(styleId, true, out WordParagraphStyles style)) {
                            if (style >= WordParagraphStyles.Heading1 && style <= WordParagraphStyles.Heading6) {
                                int levelHeading = (int)style - (int)WordParagraphStyles.Heading1 + 1;
                                tag = $"h{levelHeading}";
                            }
                        }

                        sb.Append('<').Append(tag).Append('>');
                        AppendRuns(sb, evt.Paragraph!, options);
                        sb.Append("</").Append(tag).Append('>');
                        break;
                }
            }

            sb.Append("</body></html>");
            return sb.ToString();
        }

        private static void AppendRuns(StringBuilder sb, Paragraph paragraph, WordToHtmlOptions options) {
            foreach (Run run in paragraph.Elements<Run>()) {
                string text = run.InnerText;
                string encoded = System.Net.WebUtility.HtmlEncode(text);
                RunProperties? runProps = run.RunProperties;
                string result = encoded;

                if (options.IncludeStyles && runProps?.RunFonts?.Ascii != null) {
                    result = $"<span style=\"font-family:{runProps.RunFonts.Ascii}\">{result}</span>";
                }

                if (runProps?.Underline != null && runProps.Underline.Val != UnderlineValues.None) {
                    result = $"<u>{result}</u>";
                }
                if (runProps?.Italic != null) {
                    result = $"<i>{result}</i>";
                }
                if (runProps?.Bold != null) {
                    result = $"<b>{result}</b>";
                }

                sb.Append(result);
            }
        }
    }
}