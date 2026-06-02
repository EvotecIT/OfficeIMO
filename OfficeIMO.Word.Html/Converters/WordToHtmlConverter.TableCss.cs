using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {

            string? GetWidthCss(TableWidthUnitValues? type, int? width) {
                if (type == null || width == null) {
                    return null;
                }

                if (type == TableWidthUnitValues.Pct) {
                    return $"{width.Value / 50}%";
                }

                if (type == TableWidthUnitValues.Dxa) {
                    double points = width.Value / 20.0;
                    double pixels = points * 96 / 72;
                    return $"{Math.Round(pixels)}px";
                }

                return null;
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
                        align = p.ParagraphAlignment;
                    } else if (align != p.ParagraphAlignment) {
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

                string color = !string.IsNullOrEmpty(colorHex) ? $"#{colorHex}" : "black";
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

                var left = BuildBorderCss(b.LeftStyle, b.LeftColorHex, b.LeftSize);
                var right = BuildBorderCss(b.RightStyle, b.RightColorHex, b.RightSize);
                var top = BuildBorderCss(b.TopStyle, b.TopColorHex, b.TopSize);
                var bottom = BuildBorderCss(b.BottomStyle, b.BottomColorHex, b.BottomSize);

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

                var left = BuildBorderCss(b.LeftStyle, b.LeftColorHex, b.LeftSize);
                var right = BuildBorderCss(b.RightStyle, b.RightColorHex, b.RightSize);
                var top = BuildBorderCss(b.TopStyle, b.TopColorHex, b.TopSize);
                var bottom = BuildBorderCss(b.BottomStyle, b.BottomColorHex, b.BottomSize);

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
