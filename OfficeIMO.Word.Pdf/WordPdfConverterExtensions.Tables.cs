using DocumentFormat.OpenXml;
using OfficeIMO.Word;
using QuestPDF.Fluent;
using QuestPDF.Infrastructure;
using System;
using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static IContainer RenderTable(IContainer container, WordTable table, Func<WordParagraph, (int Level, string Marker)?> getMarker) {
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
                                    cellColumn.Item().Element(e => RenderParagraph(e, paragraph, getMarker(paragraph)));
                                }

                                foreach (WordTable nested in cell.NestedTables) {
                                    cellColumn.Item().Element(e => RenderTable(e, nested, getMarker));
                                }
                            });

                            return cellContainer;
                        });
                    }
                }
            });

            return container;
        }

        private static IContainer ApplyCellStyle(IContainer container, WordTableCell cell) {
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

        private static bool HasBorder(W.BorderValues? style) => style != null && style != W.BorderValues.Nil && style != W.BorderValues.None;

        private static float GetBorderWidth(UInt32Value size) => size != null ? size.Value / 8f : 1f;
    }
}