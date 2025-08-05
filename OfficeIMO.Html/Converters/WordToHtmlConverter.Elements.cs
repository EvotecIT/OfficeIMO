using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeIMO.Html {
    public partial class WordToHtmlConverter {
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
    }
}
