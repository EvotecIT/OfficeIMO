using OfficeIMO.Word;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Converters;

namespace OfficeIMO.Pdf;

/// <summary>
/// Provides stream based conversion from Word documents to PDF.
/// </summary>
public class WordPdfConverter : IWordConverter {
    public void Convert(Stream input, Stream output, IConversionOptions options) {
        using WordDocument document = WordDocument.Load(input);
        document.SaveAsPdf(output, options as PdfSaveOptions);
    }

    public async Task ConvertAsync(Stream input, Stream output, IConversionOptions options) {
        using WordDocument document = WordDocument.Load(input);
        document.SaveAsPdf(output, options as PdfSaveOptions);
        await output.FlushAsync().ConfigureAwait(false);
    }
}

/// <summary>
/// Provides extension methods for converting <see cref="WordDocument"/> instances to PDF files.
/// </summary>
public static class WordPdfConverterExtensions {
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

        Dictionary<WordParagraph, (int Level, string Marker)> listMarkers = BuildListMarkers(document);

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
                            column.Item().Element(e => RenderParagraph(e, paragraph, GetMarker(paragraph)));
                        } else if (element is WordTable table) {
                            column.Item().Element(e => RenderTable(e, table));
                        }
                    }
                });
            });
        });

        return pdf;

        (int Level, string Marker)? GetMarker(WordParagraph paragraph) {
            if (listMarkers.TryGetValue(paragraph, out var value)) {
                return value;
            }

            return null;
        }

        void RenderElements(ColumnDescriptor column, IEnumerable<WordParagraph> paragraphs, IEnumerable<WordTable> tables) {
            foreach (WordParagraph paragraph in paragraphs) {
                column.Item().Element(e => RenderParagraph(e, paragraph, GetMarker(paragraph)));
            }

            foreach (WordTable table in tables) {
                column.Item().Element(e => RenderTable(e, table));
            }
        }

        IContainer RenderTable(IContainer container, WordTable table) {
            container.Table(tableContainer => {
                var rows = TableBuilder.Map(table).ToList();
                int columnCount = rows.Max(r => r.Count);
                tableContainer.ColumnsDefinition(columns => {
                    for (int i = 0; i < columnCount; i++) {
                        columns.RelativeColumn();
                    }
                });

                foreach (var row in rows) {
                    foreach (WordTableCell cell in row) {
                        tableContainer.Cell().Element(cellContainer => {
                            cellContainer = ApplyCellStyle(cellContainer, cell);

                            cellContainer.Column(cellColumn => {
                                foreach (WordParagraph paragraph in cell.Paragraphs) {
                                    cellColumn.Item().Element(e => RenderParagraph(e, paragraph, GetMarker(paragraph)));
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

        static IContainer RenderParagraph(IContainer container, WordParagraph paragraph, (int Level, string Marker)? marker) {
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
                    col.Item().Element(imageContainer => {
                        var img = paragraph.Image;
                        var sized = imageContainer;
                        if (img.Width.HasValue) {
                            sized = sized.Width((float)(img.Width.Value * 72 / 96));
                        }
                        if (img.Height.HasValue) {
                            sized = sized.Height((float)(img.Height.Value * 72 / 96));
                        }
                        sized.Image(ImageEmbedder.GetImageBytes(img));
                    });
                }

                string content = paragraph.IsHyperLink && paragraph.Hyperlink != null ? paragraph.Hyperlink.Text : paragraph.Text;
                if (!string.IsNullOrEmpty(content) || marker != null) {
                    if (marker != null) {
                        const float indentSize = 15f;
                        col.Item().Row(row => {
                            if (marker.Value.Level > 0) {
                                row.ConstantItem(indentSize * marker.Value.Level);
                            }
                            row.ConstantItem(indentSize).Text(marker.Value.Marker);
                            row.RelativeItem().Text(text => {
                                if (paragraph.IsHyperLink && paragraph.Hyperlink != null) {
                                    ApplyFormatting(text.Hyperlink(content, paragraph.Hyperlink.Uri.ToString()));
                                } else {
                                    ApplyFormatting(text.Span(content));
                                }
                            });
                        });
                    } else {
                        col.Item().Text(text => {
                            if (paragraph.IsHyperLink && paragraph.Hyperlink != null) {
                                ApplyFormatting(text.Hyperlink(content, paragraph.Hyperlink.Uri.ToString()));
                            } else {
                                ApplyFormatting(text.Span(content));
                            }
                        });
                    }
                }
            });

            return container;
        
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

    private static Dictionary<WordParagraph, (int Level, string Marker)> BuildListMarkers(WordDocument document) {
        Dictionary<WordParagraph, (int, string)> result = new Dictionary<WordParagraph, (int, string)>();

        foreach (WordList list in document.Lists) {
            Dictionary<int, int> indices = new Dictionary<int, int>();
            bool bullet = list.Style.ToString().IndexOf("Bullet", StringComparison.OrdinalIgnoreCase) >= 0;
            foreach (WordParagraph item in list.ListItems) {
                int level = item.ListItemLevel ?? 0;
                if (!indices.ContainsKey(level)) {
                    indices[level] = 1;
                }

                int index = indices[level];
                indices[level] = index + 1;
                string marker = bullet ? "•" : $"{index}.";
                result[item] = (level, marker);
            }
        }

        return result;
    }
}