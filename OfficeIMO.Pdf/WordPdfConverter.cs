using OfficeIMO.Word;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System.Collections.Generic;
using System.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Provides extension methods for converting <see cref="WordDocument"/> instances to PDF files.
/// </summary>
public static class WordPdfConverter {
    /// <summary>
    /// Saves the specified <see cref="WordDocument"/> as a PDF at the given <paramref name="path"/>.
    /// </summary>
    /// <param name="document">The document to convert.</param>
    /// <param name="path">The output PDF file path.</param>
    public static void SaveAsPdf(this WordDocument document, string path) {
        QuestPDF.Settings.License = LicenseType.Community;

        Dictionary<WordParagraph, string> listPrefixes = BuildListPrefixes(document);

        Document pdf = Document.Create(container => {
            container.Page(page => {
                page.Margin(1, Unit.Centimetre);

                WordHeaderFooter header = document.Header?.Default;
                if (header != null && (header.Paragraphs.Count > 0 || header.Tables.Count > 0)) {
                    page.Header().Column(col => {
                        RenderElements(col, header.Paragraphs, header.Tables);
                    });
                }

                WordHeaderFooter footer = document.Footer?.Default;
                if (footer != null && (footer.Paragraphs.Count > 0 || footer.Tables.Count > 0)) {
                    page.Footer().Column(col => {
                        RenderElements(col, footer.Paragraphs, footer.Tables);
                    });
                }

                page.Content().Column(column => {
                    foreach (WordElement element in document.Elements) {
                        if (element is WordParagraph paragraph) {
                            column.Item().Element(e => RenderParagraph(e, paragraph, GetPrefix(paragraph)));
                        } else if (element is WordTable table) {
                            column.Item().Element(e => RenderTable(e, table));
                        }
                    }
                });
            });
        });

        pdf.GeneratePdf(path);

        string GetPrefix(WordParagraph paragraph) {
            if (listPrefixes.TryGetValue(paragraph, out string value)) {
                return value;
            }

            return string.Empty;
        }

        void RenderElements(ColumnDescriptor column, IEnumerable<WordParagraph> paragraphs, IEnumerable<WordTable> tables) {
            foreach (WordParagraph paragraph in paragraphs) {
                column.Item().Element(e => RenderParagraph(e, paragraph, GetPrefix(paragraph)));
            }

            foreach (WordTable table in tables) {
                column.Item().Element(e => RenderTable(e, table));
            }
        }

        IContainer RenderTable(IContainer container, WordTable table) {
            container.Table(tableContainer => {
                int columnCount = table.Rows.Max(r => r.CellsCount);
                tableContainer.ColumnsDefinition(columns => {
                    for (int i = 0; i < columnCount; i++) {
                        columns.RelativeColumn();
                    }
                });

                foreach (WordTableRow row in table.Rows) {
                    foreach (WordTableCell cell in row.Cells) {
                        tableContainer.Cell().Column(cellColumn => {
                            foreach (WordParagraph paragraph in cell.Paragraphs) {
                                cellColumn.Item().Element(e => RenderParagraph(e, paragraph, GetPrefix(paragraph)));
                            }

                            foreach (WordTable nested in cell.NestedTables) {
                                cellColumn.Item().Element(e => RenderTable(e, nested));
                            }
                        });
                    }
                }
            });

            return container;
        }

        static IContainer RenderParagraph(IContainer container, WordParagraph paragraph, string prefix) {
            if (paragraph == null) {
                return container;
            }

            if (paragraph.ParagraphAlignment == W.JustificationValues.Center) {
                container = container.AlignCenter();
            } else if (paragraph.ParagraphAlignment == W.JustificationValues.Right) {
                container = container.AlignRight();
            } else if (paragraph.ParagraphAlignment == W.JustificationValues.Both) {
                container = container.AlignLeft();
            }

            container.Column(col => {
                if (paragraph.Image != null) {
                    col.Item().Image(paragraph.Image.GetBytes());
                }

                if (!string.IsNullOrEmpty(paragraph.Text) || !string.IsNullOrEmpty(prefix)) {
                    col.Item().Text(text => {
                        TextSpanDescriptor span = text.Span(prefix + paragraph.Text);
                        if (paragraph.Bold) {
                            span = span.Bold();
                        }
                        if (paragraph.Italic) {
                            span = span.Italic();
                        }
                        if (paragraph.Underline != null) {
                            span = span.Underline();
                        }
                        if (paragraph.Style.HasValue) {
                            switch (paragraph.Style.Value) {
                                case WordParagraphStyles.Heading1:
                                    span.FontSize(24).Bold();
                                    break;
                                case WordParagraphStyles.Heading2:
                                    span.FontSize(20).Bold();
                                    break;
                                case WordParagraphStyles.Heading3:
                                    span.FontSize(16).Bold();
                                    break;
                                case WordParagraphStyles.Heading4:
                                    span.FontSize(14).Bold();
                                    break;
                                case WordParagraphStyles.Heading5:
                                    span.FontSize(13).Bold();
                                    break;
                                case WordParagraphStyles.Heading6:
                                    span.FontSize(12).Bold();
                                    break;
                            }
                        }
                    });
                }
            });

            return container;
        }

        static Dictionary<WordParagraph, string> BuildListPrefixes(WordDocument document) {
            Dictionary<WordParagraph, string> result = new Dictionary<WordParagraph, string>();

            foreach (WordList list in document.Lists) {
                Dictionary<int, int> indices = new Dictionary<int, int>();
                bool bullet = list.Style.ToString().IndexOf("Bullet", System.StringComparison.OrdinalIgnoreCase) >= 0;
                foreach (WordParagraph item in list.ListItems) {
                    int level = item.ListItemLevel ?? 0;
                    if (!indices.ContainsKey(level)) {
                        indices[level] = 1;
                    }

                    int index = indices[level];
                    indices[level] = index + 1;
                    string prefix = bullet ? "• " : $"{index}. ";
                    string indent = new string(' ', level * 2);
                    result[item] = indent + prefix;
                }
            }

            return result;
        }
    }
}