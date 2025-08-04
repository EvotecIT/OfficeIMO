using OfficeIMO.Word;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System;
using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

#nullable enable annotations

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
    /// <param name="options">Optional PDF configuration.</param>
    public static void SaveAsPdf(this WordDocument document, string path, PdfSaveOptions? options = null) {
        Document pdf = CreatePdfDocument(document, options);
        pdf.GeneratePdf(path);
    }

    /// <summary>
    /// Saves the specified <see cref="WordDocument"/> as a PDF to the provided <paramref name="stream"/>.
    /// </summary>
    /// <param name="document">The document to convert.</param>
    /// <param name="stream">The output stream to receive the PDF data.</param>
    /// <param name="options">Optional PDF configuration.</param>
    public static void SaveAsPdf(this WordDocument document, Stream stream, PdfSaveOptions? options = null) {
        Document pdf = CreatePdfDocument(document, options);
        pdf.GeneratePdf(stream);
    }

    private static Document CreatePdfDocument(WordDocument document, PdfSaveOptions? options) {
        QuestPDF.Settings.License = LicenseType.Community;

        Dictionary<WordParagraph, ListItemInfo> listItems = BuildListItems(document);

        Document pdf = Document.Create(container => {
            container.Page(page => {
                float margin = options?.Margin ?? 1;
                Unit unit = options?.MarginUnit ?? Unit.Centimetre;
                page.Margin(margin, unit);

                if (options != null) {
                    PageSize size = options.PageSize ?? PageSizes.A4;
                    if (options.Orientation == PdfPageOrientation.Landscape) {
                        size = size.Landscape();
                    } else if (options.Orientation == PdfPageOrientation.Portrait) {
                        size = size.Portrait();
                    }

                    page.Size(size);
                }

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
                            column.Item().Element(e => RenderParagraph(e, paragraph, GetListItem(paragraph)));
                        } else if (element is WordTable table) {
                            column.Item().Element(e => RenderTable(e, table));
                        }
                    }
                });
            });
        });

        return pdf;

        ListItemInfo? GetListItem(WordParagraph paragraph) {
            if (listItems.TryGetValue(paragraph, out ListItemInfo value)) {
                return value;
            }

            return null;
        }

        void RenderElements(ColumnDescriptor column, IEnumerable<WordParagraph> paragraphs, IEnumerable<WordTable> tables) {
            foreach (WordParagraph paragraph in paragraphs) {
                column.Item().Element(e => RenderParagraph(e, paragraph, GetListItem(paragraph)));
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
                        tableContainer.Cell().Element(cellContainer => {
                            cellContainer = ApplyCellStyle(cellContainer, cell);

                            cellContainer.Column(cellColumn => {
                                foreach (WordParagraph paragraph in cell.Paragraphs) {
                                    cellColumn.Item().Element(e => RenderParagraph(e, paragraph, GetListItem(paragraph)));
                                }

                                foreach (WordTable nested in cell.NestedTables) {
                                    cellColumn.Item().Element(e => RenderTable(e, nested));
                                }
                            });

                            return cellContainer;
                        });
                    }
                }
            });

            return container;
        }

        static IContainer ApplyCellStyle(IContainer container, WordTableCell cell) {
            if (!string.IsNullOrEmpty(cell.ShadingFillColorHex)) {
                container = container.Background("#" + cell.ShadingFillColorHex);
            }

            WordTableCellBorder borders = cell.Borders;

            List<string> colors = new() {
                borders.TopColorHex,
                borders.BottomColorHex,
                borders.LeftColorHex,
                borders.RightColorHex
            };
            colors.RemoveAll(string.IsNullOrEmpty);
            if (colors.Count > 0 && colors.Distinct(StringComparer.OrdinalIgnoreCase).Count() == 1) {
                container = container.BorderColor("#" + colors[0]);
            }

            if (HasBorder(borders.TopStyle)) {
                container = container.BorderTop(GetBorderWidth(borders.TopSize));
            }
            if (HasBorder(borders.BottomStyle)) {
                container = container.BorderBottom(GetBorderWidth(borders.BottomSize));
            }
            if (HasBorder(borders.LeftStyle)) {
                container = container.BorderLeft(GetBorderWidth(borders.LeftSize));
            }
            if (HasBorder(borders.RightStyle)) {
                container = container.BorderRight(GetBorderWidth(borders.RightSize));
            }

            return container;
        }

        static bool HasBorder(W.BorderValues? style) => style != null && style != W.BorderValues.Nil && style != W.BorderValues.None;

        static float GetBorderWidth(UInt32Value size) => size != null ? size.Value / 8f : 1f;

        static IContainer RenderParagraph(IContainer container, WordParagraph paragraph, ListItemInfo? listInfo) {
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

            if (listInfo != null) {
                container.Row(row => {
                    const float indentSize = 12f;
                    row.ConstantItem(listInfo.Level * indentSize);
                    row.ConstantItem(indentSize).Text(listInfo.Marker);
                    row.RelativeItem().Column(RenderContent);
                });
            } else {
                container.Column(RenderContent);
            }

            return container;

        void RenderContent(ColumnDescriptor col) {
            if (paragraph.Image != null) {
                col.Item().Image(paragraph.Image.GetBytes());
            }

            string content = paragraph.IsHyperLink && paragraph.Hyperlink != null ? paragraph.Hyperlink.Text : paragraph.Text;
            if (!string.IsNullOrEmpty(content)) {
                col.Item().Text(text => {
                    if (paragraph.IsHyperLink && paragraph.Hyperlink != null) {
                        ApplyFormatting(text.Hyperlink(content, paragraph.Hyperlink.Uri.ToString()));
                    } else {
                        ApplyFormatting(text.Span(content));
                    }
                });
            }
        }

        void ApplyFormatting(TextSpanDescriptor span) {
            if (paragraph.Bold) {
                span = span.Bold();
            }
            if (paragraph.Italic) {
                span = span.Italic();
            }
            if (paragraph.Underline != null) {
                span = span.Underline();
            }
            if (!string.IsNullOrEmpty(paragraph.ColorHex)) {
                span = span.FontColor("#" + paragraph.ColorHex);
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
        }
        }
        }

    private static Dictionary<WordParagraph, ListItemInfo> BuildListItems(WordDocument document) {
        Dictionary<WordParagraph, ListItemInfo> result = new Dictionary<WordParagraph, ListItemInfo>();

        foreach (WordList list in document.Lists) {
            Dictionary<int, int> indices = new Dictionary<int, int>();
            bool bullet = list.Style.ToString().IndexOf("Bullet", StringComparison.OrdinalIgnoreCase) >= 0;

            foreach (WordParagraph item in list.ListItems) {
                int level = item.ListItemLevel ?? 0;

                foreach (int key in indices.Keys.Where(k => k > level).ToList()) {
                    indices.Remove(key);
                }

                if (!indices.ContainsKey(level)) {
                    indices[level] = 1;
                }

                int index = indices[level];

                string marker;
                if (bullet) {
                    marker = "•";
                } else {
                    var parts = indices.Where(kv => kv.Key <= level)
                        .OrderBy(kv => kv.Key)
                        .Select(kv => kv.Value.ToString());
                    marker = string.Join('.', parts) + ".";
                }

                result[item] = new ListItemInfo {
                    Marker = marker,
                    Level = level,
                    IsBullet = bullet
                };

                indices[level] = index + 1;
            }
        }

        return result;
    }

    private class ListItemInfo {
        public int Level { get; set; }
        public string Marker { get; set; } = string.Empty;
        public bool IsBullet { get; set; }
    }
}