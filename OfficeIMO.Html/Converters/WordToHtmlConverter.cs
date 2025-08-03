using System;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

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

            foreach (Paragraph paragraph in document.MainDocumentPart!.Document.Body!.Elements<Paragraph>()) {
                string tag = "p";
                string? styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                if (styleId != null && Enum.TryParse(styleId, true, out WordParagraphStyles style)) {
                    if (style >= WordParagraphStyles.Heading1 && style <= WordParagraphStyles.Heading6) {
                        int level = (int)style - (int)WordParagraphStyles.Heading1 + 1;
                        tag = $"h{level}";
                    }
                }

                sb.Append('<').Append(tag).Append('>');
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
                sb.Append("</").Append(tag).Append('>');
            }

            sb.Append("</body></html>");
            return sb.ToString();
        }
    }
}
