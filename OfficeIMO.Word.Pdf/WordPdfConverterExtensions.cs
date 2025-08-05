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

namespace OfficeIMO.Word.Pdf {

/// <summary>
/// Provides extension methods for converting <see cref="WordDocument"/> instances to PDF files.
/// </summary>
public static partial class WordPdfConverterExtensions {
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

                PageSize size = PageSizes.A4;
                PdfPageOrientation? orientation = null;
                if (options != null) {
                    if (options.PageSize != null) {
                        size = options.PageSize;
                    } else if (options.DefaultPageSize.HasValue) {
                        size = MapToPageSize(options.DefaultPageSize.Value);
                    }

                    orientation = options.Orientation;
                    if (orientation == null && options.DefaultOrientation.HasValue) {
                        orientation = options.DefaultOrientation == W.PageOrientationValues.Landscape ? PdfPageOrientation.Landscape : PdfPageOrientation.Portrait;
                    }
                }

                if (orientation == PdfPageOrientation.Landscape) {
                    size = size.Landscape();
                } else if (orientation == PdfPageOrientation.Portrait) {
                    size = size.Portrait();
                }

                page.Size(size);

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
                TableLayout layout = TableLayoutCache.GetLayout(table);
                tableContainer.ColumnsDefinition(columns => {
                    foreach (float width in layout.ColumnWidths) {
                        if (width > 0) {
                            columns.ConstantColumn(width);
                        } else {
                            columns.RelativeColumn();
                        }
                    }
                });

                foreach (IReadOnlyList<WordTableCell> row in layout.Rows) {
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

    }
}
}
