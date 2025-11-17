using DocumentFormat.OpenXml;
using QuestPDF.Fluent;
using QuestPDF.Infrastructure;
using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static IContainer RenderTable(IContainer container, WordTable table, Func<WordParagraph, (int Level, string Marker)?> getMarker, PdfSaveOptions? options, Dictionary<WordParagraph, int> footnoteMap) {
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
                            cellContainer = ApplyCellStyle(cellContainer, cell, options);

                            cellContainer.Column(cellColumn => {
                                foreach (WordParagraph paragraph in cell.Paragraphs) {
                                    cellColumn.Item().Element(e => {
                                        return RenderParagraph(e, paragraph, getMarker(paragraph), options, footnoteMap);
                                    });
                                }

                                foreach (WordTable nested in cell.NestedTables) {
                                    cellColumn.Item().Element(e => RenderTable(e, nested, getMarker, options, footnoteMap));
                                }
                            });

                            return cellContainer;
                        });
                    }
                }
            });

            return container;
        }

        private static IContainer ApplyCellStyle(IContainer container, WordTableCell cell, PdfSaveOptions? options) {
            if (!string.IsNullOrEmpty(cell.ShadingFillColorHex)) {
                // Ignore automatic color to let QuestPDF use its own defaults.
                if (!cell.ShadingFillColorHex.Equals("auto", StringComparison.OrdinalIgnoreCase)) {
                    container = container.Background("#" + cell.ShadingFillColorHex);
                }
            }

            WordTableCellBorder borders = cell.Borders;

            List<string?> colors = new()
            {
                borders.TopColorHex,
                borders.BottomColorHex,
                borders.LeftColorHex,
                borders.RightColorHex
            };
	            // Filter out empty and automatic colors – QuestPDF expects real hex values.
	            colors.RemoveAll(c =>
	                string.IsNullOrEmpty(c) || string.Equals(c, "auto", StringComparison.OrdinalIgnoreCase));
            if (colors.Count > 0 && colors.Distinct(StringComparer.OrdinalIgnoreCase).Count() == 1) {
                container = container.BorderColor("#" + colors[0]!);
            }

            bool anyBorder = false;
            if (HasBorder(borders.TopStyle)) {
                container = container.BorderTop(GetBorderWidth(borders.TopSize));
                anyBorder = true;
            }
            if (HasBorder(borders.BottomStyle)) {
                container = container.BorderBottom(GetBorderWidth(borders.BottomSize));
                anyBorder = true;
            }
            if (HasBorder(borders.LeftStyle)) {
                container = container.BorderLeft(GetBorderWidth(borders.LeftSize));
                anyBorder = true;
            }
            if (HasBorder(borders.RightStyle)) {
                container = container.BorderRight(GetBorderWidth(borders.RightSize));
                anyBorder = true;
            }

            if (!anyBorder && options?.DefaultTableBorders == true) {
                container = container.Border(0.75f).BorderColor("#d6d6d6");
            }

            return container;
        }

        private static bool HasBorder(W.BorderValues? style) => style != null && style != W.BorderValues.Nil && style != W.BorderValues.None;

        private static float GetBorderWidth(UInt32Value? size) => size != null ? size.Value / 8f : 1f;
    }
}
