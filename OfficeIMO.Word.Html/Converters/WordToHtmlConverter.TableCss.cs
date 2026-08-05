using AngleSharp.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        private static IEnumerable<WordElement> ExpandTableCellBlockContent(WordTableCell cell) =>
            ExpandTableCellBlockContent(cell.Document, cell._tableCell);

        private static IEnumerable<WordElement> ExpandTableCellBlockContent(
            WordDocument document,
            OpenXmlCompositeElement container) {
            foreach (OpenXmlElement child in container.ChildElements) {
                if (child is Paragraph paragraph) {
                    foreach (WordParagraph converted in WordSection.ConvertParagraphToWordParagraphs(document, paragraph)) {
                        yield return converted;
                    }
                } else if (child is Table table) {
                    yield return new WordTable(document, table);
                } else if (child is OpenXmlCompositeElement blockWrapper && child is not TableCellProperties) {
                    foreach (WordElement nested in ExpandTableCellBlockContent(document, blockWrapper)) {
                        yield return nested;
                    }
                }
            }
        }

        private static bool IsMarkedListItemTable(WordTable table) {
            if (table._table.Parent is not SdtContentBlock content ||
                content.Parent is not SdtBlock control) {
                return false;
            }
            return string.Equals(
                control.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value,
                HtmlWordRoundTripMarkers.ListItemTableTag,
                StringComparison.Ordinal);
        }

        private static bool IsFollowedByContinuingListParagraph(
            IReadOnlyList<WordElement> elements,
            int tableIndex) {
            int? precedingNumberId = null;
            for (int index = tableIndex - 1; index >= 0; index--) {
                if (elements[index] is WordParagraph paragraph && DocumentTraversal.GetListInfo(paragraph) != null) {
                    precedingNumberId = paragraph._listNumberId;
                    break;
                }
            }
            if (!precedingNumberId.HasValue) return false;

            for (int index = tableIndex + 1; index < elements.Count; index++) {
                if (elements[index] is not WordParagraph paragraph || paragraph.IsEmpty) continue;
                return DocumentTraversal.GetListInfo(paragraph) != null &&
                       paragraph._listNumberId == precedingNumberId;
            }
            return false;
        }


            string? GetWidthCss(TableWidthUnitValues? type, int? width) {
                if (type == null || width == null) {
                    return null;
                }

                if (type == TableWidthUnitValues.Pct) {
                    return $"{FormatCssNumber(width.Value / 50.0)}%";
                }

                if (type == TableWidthUnitValues.Dxa) {
                    double points = width.Value / 20.0;
                    double pixels = points * 96 / 72;
                    return $"{Math.Round(pixels)}px";
                }

                return null;
            }

            string? GetTableCellSpacingCss(WordTable table) {
                var cellSpacing = table.StyleDetails?.CellSpacing;
                if (cellSpacing == null || cellSpacing.Value <= 0) {
                    return null;
                }

                return FormatTwips(cellSpacing.Value);
            }

            void AppendColumnGroup(IDocument htmlDoc, IElement tableElement, WordTable table) {
                var columns = GetColumnWidths(table);
                if (columns.Count == 0) {
                    return;
                }

                var colGroup = CreateOutputElement(htmlDoc, "colgroup");
                foreach (var (type, width) in columns) {
                    var col = CreateOutputElement(htmlDoc, "col");
                    var widthCss = GetWidthCss(type, width);
                    if (!string.IsNullOrEmpty(widthCss)) {
                        SetOutputAttribute(col, "style", $"width:{widthCss}", "TableColumn:style");
                    }
                    colGroup.AppendChild(col);
                }
                tableElement.AppendChild(colGroup);
            }

            List<(TableWidthUnitValues? Type, int Width)> GetColumnWidths(WordTable table) {
                if (table.Rows.Count == 0) {
                    return new List<(TableWidthUnitValues? Type, int Width)>();
                }

                var gridColumns = GetGridColumnWidths(table);
                var firstRow = table.Rows[0];
                var columns = new List<(TableWidthUnitValues? Type, int Width)>();
                foreach (var cell in firstRow.Cells) {
                    if (cell.HorizontalMerge == WordCellMerge.Continue || cell.VerticalMerge == WordCellMerge.Continue) {
                        return gridColumns;
                    }
                    if (cell.HorizontalMerge == WordCellMerge.Restart || cell.VerticalMerge == WordCellMerge.Restart) {
                        return gridColumns;
                    }
                    if (cell.WidthType == null || cell.Width == null || GetWidthCss(cell.WidthType?.ToOpenXml(), cell.Width) == null) {
                        return gridColumns;
                    }

                    columns.Add((cell.WidthType?.ToOpenXml(), cell.Width.Value));
                }

                if (columns.Count > 0) {
                    return columns;
                }

                return gridColumns;
            }

            List<(TableWidthUnitValues? Type, int Width)> GetGridColumnWidths(WordTable table) {
                var gridWidths = table.GridColumnWidth;
                return gridWidths.Count > 0
                    ? gridWidths.Select(width => ((TableWidthUnitValues?)TableWidthUnitValues.Dxa, width)).ToList()
                    : new List<(TableWidthUnitValues? Type, int Width)>();
            }

            static string FormatCssNumber(double value) {
                return Math.Round(value, 2).ToString("0.##", CultureInfo.InvariantCulture);
            }

            string? GetTextAlignCss(JustificationValues? justification) {
                if (justification == null) {
                    return null;
                }

                if (justification == JustificationValues.Center) {
                    return "center";
                }

                if (justification == JustificationValues.Right) {
                    return "right";
                }

                if (justification == JustificationValues.Left) {
                    return "left";
                }

                if (justification == JustificationValues.Both) {
                    return "justify";
                }

                return null;
            }

            JustificationValues? GetCellAlignment(WordTableCell cell) {
                JustificationValues? align = null;
                foreach (var p in cell.Paragraphs) {
                    if (p.ParagraphAlignment == null) {
                        continue;
                    }
                    if (align == null) {
                        align = p.ParagraphAlignment?.ToOpenXml();
                    } else if (align != p.ParagraphAlignment?.ToOpenXml()) {
                        return null;
                    }
                }
                return align;
            }

            string? BuildBorderCss(BorderValues? style, string? colorHex, UInt32Value? size) {
                if (style == null) {
                    return null;
                }

                string cssStyle = "solid";
                if (style == BorderValues.Dashed) {
                    cssStyle = "dashed";
                } else if (style == BorderValues.Dotted) {
                    cssStyle = "dotted";
                } else if (style == BorderValues.Double) {
                    cssStyle = "double";
                }

                string? normalizedColor = NormalizeSixDigitHexColor(colorHex);
                string color = normalizedColor != null ? $"#{normalizedColor}" : "black";
                double widthPt = size != null ? size.Value / 8.0 : 1.0;
                double widthPx = widthPt * 96 / 72;
                string width = $"{Math.Round(widthPx)}px";
                return $"{width} {cssStyle} {color}";
            }

            List<string> GetBorderCss(WordTableCell cell) {
                List<string> styles = new();
                var b = cell.Borders;
                if (b == null) {
                    return styles;
                }

                var left = BuildBorderCss(b.LeftStyle?.ToOpenXml(), b.LeftColorHex, b.LeftSize);
                var right = BuildBorderCss(b.RightStyle?.ToOpenXml(), b.RightColorHex, b.RightSize);
                var top = BuildBorderCss(b.TopStyle?.ToOpenXml(), b.TopColorHex, b.TopSize);
                var bottom = BuildBorderCss(b.BottomStyle?.ToOpenXml(), b.BottomColorHex, b.BottomSize);

                if (left == null && right == null && top == null && bottom == null) {
                    return styles;
                }

                if (left == top && top == right && right == bottom && left != null) {
                    styles.Add($"border:{left}");
                } else {
                    if (left != null) {
                        styles.Add($"border-left:{left}");
                    }
                    if (right != null) {
                        styles.Add($"border-right:{right}");
                    }
                    if (top != null) {
                        styles.Add($"border-top:{top}");
                    }
                    if (bottom != null) {
                        styles.Add($"border-bottom:{bottom}");
                    }
                }

                return styles;
            }

            List<string> GetParagraphBorderCss(WordParagraph p) {
                List<string> styles = new();
                var b = p.Borders;
                if (b == null) return styles;

                var left = BuildBorderCss(b.LeftStyle?.ToOpenXml(), b.LeftColorHex, b.LeftSize);
                var right = BuildBorderCss(b.RightStyle?.ToOpenXml(), b.RightColorHex, b.RightSize);
                var top = BuildBorderCss(b.TopStyle?.ToOpenXml(), b.TopColorHex, b.TopSize);
                var bottom = BuildBorderCss(b.BottomStyle?.ToOpenXml(), b.BottomColorHex, b.BottomSize);

                if (left == null && right == null && top == null && bottom == null) {
                    return styles;
                }
                if (left == top && top == right && right == bottom && left != null) {
                    styles.Add($"border:{left}");
                } else {
                    if (left != null) styles.Add($"border-left:{left}");
                    if (right != null) styles.Add($"border-right:{right}");
                    if (top != null) styles.Add($"border-top:{top}");
                    if (bottom != null) styles.Add($"border-bottom:{bottom}");
                }
                return styles;
            }

            bool CellHasBorder(WordTableCell cell) {
                var b = cell.Borders;
                return b != null && (b.LeftStyle != null || b.RightStyle != null || b.TopStyle != null || b.BottomStyle != null);
            }

            bool TableHasBorder(WordTable table) {
                return table.Rows.Any(r => r.Cells.Any(CellHasBorder));
            }
    }
}
