using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Html {
    /// <summary>
    /// Converts WordprocessingDocument content into simple HTML fragments.
    /// </summary>
    public class WordToHtmlConverter : IWordConverter {
        /// <summary>
        /// Converts a DOCX contained in the provided stream into HTML.
        /// </summary>
        /// <param name="docxStream">Stream containing DOCX content.</param>
        /// <param name="options">Conversion options.</param>
        /// <returns>Generated HTML string.</returns>
        public static string Convert(Stream docxStream, WordToHtmlOptions? options = null) {
            if (docxStream == null) {
                throw new ConversionException($"{nameof(docxStream)} cannot be null.");
            }

            options ??= new WordToHtmlOptions();

            using WordprocessingDocument document = WordprocessingDocument.Open(docxStream, false);
            StringBuilder sb = new StringBuilder();
            sb.Append("<html><body>");

            Dictionary<int, bool> listTypes = ListParser.GetListTypes(document.MainDocumentPart!);
            AppendElements(document.MainDocumentPart!.Document.Body!.ChildElements, sb, options, listTypes, document.MainDocumentPart);

            sb.Append("</body></html>");
            return sb.ToString();
        }

        private static void AppendElements(IEnumerable<OpenXmlElement> elements, StringBuilder sb, WordToHtmlOptions options, Dictionary<int, bool> listTypes, MainDocumentPart mainPart) {
            Stack<(int numId, bool ordered)> listStack = new Stack<(int numId, bool ordered)>();

            foreach (OpenXmlElement element in elements) {
                if (element is Paragraph paragraph) {
                    NumberingProperties? numProps = paragraph.ParagraphProperties?.NumberingProperties;
                    if (numProps != null) {
                        int level = numProps.NumberingLevelReference?.Val ?? 0;
                        int numId = numProps.NumberingId?.Val ?? 0;
                        bool ordered = listTypes.ContainsKey(numId) && listTypes[numId];

                        if (listStack.Count == 0) {
                            for (int i = 0; i <= level; i++) {
                                string tagOpen = ordered ? "<ol>" : "<ul>";
                                if (options.IncludeListStyles) {
                                    string listStyle = ordered ? "decimal" : "disc";
                                    tagOpen = ordered ? $"<ol style=\"list-style-type:{listStyle}\">" : $"<ul style=\"list-style-type:{listStyle}\">";
                                }
                                sb.Append(tagOpen);
                                listStack.Push((numId, ordered));
                            }
                        } else {
                            int currentLevel = listStack.Count - 1;
                            if (level > currentLevel) {
                                for (int i = currentLevel + 1; i <= level; i++) {
                                    string tagOpen = ordered ? "<ol>" : "<ul>";
                                    if (options.IncludeListStyles) {
                                        string listStyle = ordered ? "decimal" : "disc";
                                        tagOpen = ordered ? $"<ol style=\"list-style-type:{listStyle}\">" : $"<ul style=\"list-style-type:{listStyle}\">";
                                    }
                                    sb.Append(tagOpen);
                                    listStack.Push((numId, ordered));
                                }
                            } else {
                                while (currentLevel > level) {
                                    var closing = listStack.Pop();
                                    sb.Append(closing.ordered ? "</ol>" : "</ul>");
                                    currentLevel--;
                                }
                                if (listStack.Count > 0 && listStack.Peek().numId != numId) {
                                    var closing = listStack.Pop();
                                    sb.Append(closing.ordered ? "</ol>" : "</ul>");
                                    string tagOpen = ordered ? "<ol>" : "<ul>";
                                    if (options.IncludeListStyles) {
                                        string listStyle = ordered ? "decimal" : "disc";
                                        tagOpen = ordered ? $"<ol style=\"list-style-type:{listStyle}\">" : $"<ul style=\"list-style-type:{listStyle}\">";
                                    }
                                    sb.Append(tagOpen);
                                    listStack.Push((numId, ordered));
                                }
                            }
                        }

                        sb.Append("<li>");
                        AppendRuns(sb, paragraph, options, mainPart);
                        sb.Append("</li>");
                    } else {
                        while (listStack.Count > 0) {
                            var closing = listStack.Pop();
                            sb.Append(closing.ordered ? "</ol>" : "</ul>");
                        }

                        string tag = "p";
                        string? styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                        if (styleId != null && Enum.TryParse(styleId, true, out WordParagraphStyles style)) {
                            if (style >= WordParagraphStyles.Heading1 && style <= WordParagraphStyles.Heading6) {
                                int levelHeading = (int)style - (int)WordParagraphStyles.Heading1 + 1;
                                tag = $"h{levelHeading}";
                            }
                        }

                        sb.Append('<').Append(tag).Append('>');
                        AppendRuns(sb, paragraph, options, mainPart);
                        sb.Append("</").Append(tag).Append('>');
                    }
                } else if (element is Table table) {
                    while (listStack.Count > 0) {
                        var closing = listStack.Pop();
                        sb.Append(closing.ordered ? "</ol>" : "</ul>");
                    }

                    AppendTable(sb, table, options, listTypes, mainPart);
                }
            }

            while (listStack.Count > 0) {
                var closing = listStack.Pop();
                sb.Append(closing.ordered ? "</ol>" : "</ul>");
            }
        }

        private static void AppendTable(StringBuilder sb, Table table, WordToHtmlOptions options, Dictionary<int, bool> listTypes, MainDocumentPart mainPart) {
            sb.Append("<table>");
            foreach (TableRow row in table.Elements<TableRow>()) {
                sb.Append("<tr>");
                foreach (TableCell cell in row.Elements<TableCell>()) {
                    sb.Append("<td>");
                    AppendElements(cell.ChildElements, sb, options, listTypes, mainPart);
                    sb.Append("</td>");
                }
                sb.Append("</tr>");
            }
            sb.Append("</table>");
        }


        private static void AppendRuns(StringBuilder sb, Paragraph paragraph, WordToHtmlOptions options, MainDocumentPart mainPart) {
            foreach (Run run in paragraph.Elements<Run>()) {
                Drawing? drawing = run.GetFirstChild<Drawing>();
                if (drawing != null) {
                    A.Blip? blip = drawing.Descendants<A.Blip>().FirstOrDefault();
                    string? embed = blip?.Embed;
                    if (embed != null) {
                        ImagePart part = (ImagePart)mainPart.GetPartById(embed);
                        using Stream imgStream = part.GetStream();
                        using MemoryStream ms = new MemoryStream();
                        imgStream.CopyTo(ms);
                        string base64 = System.Convert.ToBase64String(ms.ToArray());
                        sb.Append($"<img src=\"data:{part.ContentType};base64,{base64}\" />");
                    }
                    continue;
                }

                string text = run.InnerText;
                string encoded = System.Net.WebUtility.HtmlEncode(text);
                RunProperties? runProps = run.RunProperties;
                string result = encoded;

                if (options.IncludeFontStyles && runProps?.RunFonts?.Ascii != null) {
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
        public void Convert(Stream input, Stream output, IConversionOptions options) {
            string html = Convert(input, options as WordToHtmlOptions);
            using StreamWriter writer = new StreamWriter(
                output,
                Encoding.UTF8,
                bufferSize: 1024,
                leaveOpen: true);
            writer.Write(html);
            writer.Flush();
        }

        public async Task ConvertAsync(Stream input, Stream output, IConversionOptions options) {
            string html = Convert(input, options as WordToHtmlOptions);
            using StreamWriter writer = new StreamWriter(
                output,
                Encoding.UTF8,
                bufferSize: 1024,
                leaveOpen: true);
            await writer.WriteAsync(html).ConfigureAwait(false);
            await writer.FlushAsync().ConfigureAwait(false);
        }
    }
}