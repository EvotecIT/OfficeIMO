using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System;
using System.Collections.Generic;
using System.Linq;




namespace OfficeIMO.Pdf;

/// <summary>
/// Provides extension methods for converting <see cref="WordDocument"/> instances to PDF files.
/// </summary>
    /// <summary>
    /// Saves the specified <see cref="WordDocument"/> as a PDF at the given <paramref name="path"/>.
    /// </summary>
    /// <param name="document">The document to convert.</param>
    /// <param name="path">The output PDF file path.</param>
                        bool bullet = list.Style.ToString().IndexOf("Bullet", StringComparison.OrdinalIgnoreCase) >= 0;
namespace OfficeIMO.Pdf;

public static class WordPdfConverter {
    public static void SaveAsPdf(this WordDocument document, string path) {
        QuestPDF.Settings.License = LicenseType.Community;
        var pdf = QuestPDF.Fluent.Document.Create(container => {
            container.Page(page => {
                page.Margin(1, Unit.Centimetre);

                var headerParagraphs = document.Header?.Default?.Paragraphs ?? new List<WordParagraph>();
                if (headerParagraphs.Count > 0) {
                    page.Header().Column(header => {
                        foreach (var paragraph in headerParagraphs) {
                            header.Item().Element(e => RenderParagraph(e, paragraph));
                        }
                    });
                }

                var footerParagraphs = document.Footer?.Default?.Paragraphs ?? new List<WordParagraph>();
                if (footerParagraphs.Count > 0) {
                    page.Footer().Column(footer => {
                        foreach (var paragraph in footerParagraphs) {
                            footer.Item().Element(e => RenderParagraph(e, paragraph));
                        }
                    });
                }

                page.Content().Column(column => {
                    foreach (var paragraph in document.Paragraphs) {
                        column.Item().Element(e => RenderParagraph(e, paragraph));
                    }

                    foreach (var list in document.Lists) {
                        var index = 1;
                        bool bullet = list.Style.ToString().Contains("Bullet", StringComparison.OrdinalIgnoreCase);
                        foreach (var item in list.ListItems) {
                            var prefix = bullet ? "• " : $"{index++}. ";
                            column.Item().Element(e => RenderParagraph(e, item, prefix));
                        }
                    }

                    foreach (var table in document.Tables) {
                        column.Item().Table(tableContainer => {
                            var columnCount = table.Rows.Max(r => r.CellsCount);
                            tableContainer.ColumnsDefinition(columns => {
                                for (int i = 0; i < columnCount; i++) {
                                    columns.RelativeColumn();
                                }
                            });

                            foreach (var row in table.Rows) {
                                foreach (var cell in row.Cells) {
                                    var firstParagraph = cell.Paragraphs.FirstOrDefault();
                                    if (firstParagraph != null) {
                                        tableContainer.Cell().Element(e => RenderParagraph(e, firstParagraph));
                                    } else {
                                        tableContainer.Cell();
                                    }
                                }
                            }
                        });
                    }

                    foreach (var image in document.Images) {
                        column.Item().Image(image.GetBytes());
                    }
                });
            });
        });

        pdf.GeneratePdf(path);

        static IContainer RenderParagraph(IContainer container, WordParagraph paragraph, string prefix = "") {
            if (paragraph == null) {
                return container;
            }

            if (paragraph.ParagraphAlignment == JustificationValues.Center) {
                container = container.AlignCenter();
            } else if (paragraph.ParagraphAlignment == JustificationValues.Right) {
                container = container.AlignRight();
            } else if (paragraph.ParagraphAlignment == JustificationValues.Both) {
                container = container.AlignLeft();
            }

            container.Text(text => {
                var span = text.Span(prefix + paragraph.Text);

                if (paragraph.Bold) {
                    span = span.Bold();
                }

                if (paragraph.Italic) {
                    span = span.Italic();
                }

                if (paragraph.Underline != null) {
                    span = span.Underline();
                }

}
                    switch (paragraph.Style.Value) {
                        case WordParagraphStyles.Heading1:
                            span.FontSize(24);
                            span.Bold();
                            break;
                        case WordParagraphStyles.Heading2:
                            span.FontSize(20);
                            span.Bold();
                            break;
                        case WordParagraphStyles.Heading3:
                            span.FontSize(16);
                            span.Bold();
                            break;
                        case WordParagraphStyles.Heading4:
                            span.FontSize(14);
                            span.Bold();
                            break;
                        case WordParagraphStyles.Heading5:
                            span.FontSize(13);
                            span.Bold();
                            break;
                        case WordParagraphStyles.Heading6:
                            span.FontSize(12);
                            span.Bold();
                            break;
                    }
                }
            });

            return container;
        }
    }
}