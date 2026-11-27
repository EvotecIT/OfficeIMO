using AngleSharp.Dom;
using System.Globalization;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        private static IEnumerable<WordElement> ExpandTableCellBlockContent(WordTableCell cell) =>
            ExpandTableCellBlockContent(
                cell.Document,
                cell.Node(),
                int.MaxValue,
                CancellationToken.None);

        private static IEnumerable<WordElement> ExpandTableCellBlockContent(
            WordDocument document,
            OfficeXmlNode container,
            long maxElements,
            CancellationToken cancellationToken) {
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maxElements);
            var pending = new Stack<OfficeXmlNode>();
            PushChildrenInDocumentOrder(container, pending);
            long visited = 0;
            while (pending.Count > 0) {
                cancellationToken.ThrowIfCancellationRequested();
                if (++visited > maxElements) {
                    throw new InvalidDataException(
                        $"The Word table-cell content exceeds the {maxElements}-element HTML conversion limit.");
                }
                OfficeXmlNode child = pending.Pop();
                // A w:p or w:tbl held untyped, inside unknown markup, has no model object; its children are read like a wrapper's.
                if (child.IsW("p") && child.IsModelElement()) {
                    foreach (WordParagraph converted in document.ConvertParagraphNodeToWordParagraphs(child)) {
                        yield return converted;
                    }
                } else if (child.IsW("tbl") && child.IsModelElement()) {
                    yield return WordConverterNodes.CreateTable(document, child);
                } else if (!child.IsW("tcPr")) {
                    // Leaf elements hold no children, so only wrappers (content controls, custom XML,
                    // alternate content) add anything here.
                    PushChildrenInDocumentOrder(child, pending);
                }
            }
        }

        private static void PushChildrenInDocumentOrder(
            OfficeXmlNode container,
            Stack<OfficeXmlNode> pending) {
            for (OfficeXmlNode child = container.LastChild; !child.IsNone; child = child.PreviousSibling) {
                pending.Push(child);
            }
        }

        private static bool IsMarkedListItemTable(WordTable table) {
            OfficeXmlNode content = table.Node().Parent;
            OfficeXmlNode control = content.Parent;
            if (!content.IsW("sdtContent") || control.ContentControlLevel() != WordContentControlLevel.Block) {
                return false;
            }
            return string.Equals(
                control.WChild("sdtPr").ChildVal("tag"),
                HtmlWordRoundTripMarkers.ListItemTableTag,
                StringComparison.Ordinal);
        }

        private static readonly HashSet<string> KnownMonospaceFonts = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
            "Consolas", "Courier", "Courier New", "Lucida Console", "Lucida Sans Typewriter", "Menlo", "Monaco",
            "SF Mono", "SFMono-Regular", "DejaVu Sans Mono", "Liberation Mono", "Cousine", "Inconsolata",
            "Source Code Pro", "Fira Code", "Fira Mono", "Cascadia Mono", "Cascadia Code", "JetBrains Mono",
            "Roboto Mono", "Ubuntu Mono", "Noto Sans Mono", "Andale Mono", "monospace"
        };

        // A monospace run is code unless monospace is simply the document's own default font.
        private static bool IsInlineCodeRun(WordParagraph run, string? documentFont) =>
            IsMonospaceFontFamily(run.FontFamily)
            && !string.Equals(run.FontFamily, documentFont, StringComparison.OrdinalIgnoreCase);

        /// <summary>Recognises monospace families by name, independent of the fonts installed on this machine.</summary>
        internal static bool IsMonospaceFontFamily(string? family) {
            if (string.IsNullOrWhiteSpace(family)) return false;
            string name = family!.Trim().Trim('"', '\'');
            return KnownMonospaceFonts.Contains(name)
                || name.IndexOf(" Mono", StringComparison.OrdinalIgnoreCase) >= 0
                || name.EndsWith("Mono", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsFollowedByContinuingListParagraph(
            IReadOnlyList<WordElement> elements,
            int tableIndex) {
            int? precedingNumberId = null;
            for (int index = tableIndex - 1; index >= 0; index--) {
                if (elements[index] is WordParagraph paragraph && WordDocumentTraversal.GetListInfo(paragraph) != null) {
                    precedingNumberId = paragraph._listNumberId;
                    break;
                }
            }
            if (!precedingNumberId.HasValue) return false;

            for (int index = tableIndex + 1; index < elements.Count; index++) {
                if (elements[index] is not WordParagraph paragraph || paragraph.IsEmpty) continue;
                return WordDocumentTraversal.GetListInfo(paragraph) != null &&
                       paragraph._listNumberId == precedingNumberId;
            }
            return false;
        }


            string? GetWidthCss(WordTableWidth? width) {
                if (width?.PercentValue is double percent) return $"{FormatCssNumber(percent)}%";
                if (width?.LengthValue is { } length) return $"{FormatCssNumber(length.ToPixels())}px";
                if (width?.Kind == WordTableWidthKind.Auto) return "auto";
                return null;
            }

            // A fixed Word width is a share of the page's text column. Expressing it as a share of
            // the HTML body keeps the table's proportions instead of pinning it to a narrow pixel width.
            string? GetTopLevelTableWidthCss(WordTable table) {
                if (table.Width?.LengthValue is not { } length || length <= OfficeLength.Zero) {
                    return GetWidthCss(table.Width);
                }

                WordSection? section = table.Document.Sections.FirstOrDefault();
                double pageWidth = section?.PageSetup.Size.Width?.ToPoints() ?? 612D;
                double textWidth = pageWidth - (section?.PageSetup.Margins.Left?.ToPoints() ?? 72D) - (section?.PageSetup.Margins.Right?.ToPoints() ?? 72D);
                if (textWidth <= 0D) return GetWidthCss(table.Width);
                double percent = Math.Min(100D, length.ToPoints() * 100D / textWidth);
                return $"{FormatCssNumber(percent)}%";
            }

            string? GetTableCellSpacingCss(WordTable table) {
                var cellSpacing = table.StyleDetails?.CellSpacing;
                if (cellSpacing == null || cellSpacing.Value <= global::OfficeIMO.OfficeLength.Zero) {
                    return null;
                }

                return FormatTwips(checked((int)cellSpacing.Value.ToTwips()));
            }

            void AppendColumnGroup(IDocument htmlDoc, IElement tableElement, WordTable table) {
                var columns = GetColumnWidths(table);
                if (columns.Count == 0) {
                    return;
                }

                var colGroup = CreateOutputElement(htmlDoc, "colgroup");
                foreach (var width in columns) {
                    var col = CreateOutputElement(htmlDoc, "col");
                    var widthCss = GetWidthCss(width);
                    if (!string.IsNullOrEmpty(widthCss)) {
                        SetOutputAttribute(col, "style", $"width:{widthCss}", "TableColumn:style");
                    }
                    colGroup.AppendChild(col);
                }
                tableElement.AppendChild(colGroup);
            }

            List<WordTableWidth?> GetColumnWidths(WordTable table) {
                if (table.Rows.Count == 0) {
                    return new List<WordTableWidth?>();
                }

                var gridColumns = GetGridColumnWidths(table);
                var firstRow = table.Rows[0];
                var columns = new List<WordTableWidth?>();
                foreach (var cell in firstRow.Cells) {
                    if (cell.HorizontalMerge == WordCellMerge.Continue || cell.VerticalMerge == WordCellMerge.Continue) {
                        return gridColumns;
                    }
                    if (cell.HorizontalMerge == WordCellMerge.Restart || cell.VerticalMerge == WordCellMerge.Restart) {
                        return gridColumns;
                    }
                    if (cell.Width == null || GetWidthCss(cell.Width) == null) {
                        return gridColumns;
                    }

                    columns.Add(cell.Width);
                }

                if (columns.Count > 0) {
                    return columns;
                }

                return gridColumns;
            }

            List<WordTableWidth?> GetGridColumnWidths(WordTable table) {
                var gridWidths = table.GridColumnWidth;
                return gridWidths.Count > 0
                    ? gridWidths.Select(width => (WordTableWidth?)WordTableWidth.Length(width)).ToList()
                    : new List<WordTableWidth?>();
            }

            static string FormatCssNumber(double value) {
                return Math.Round(value, 2).ToString("0.##", CultureInfo.InvariantCulture);
            }

            string? GetTextAlignCss(WordParagraphAlignment? justification) {
                if (justification == null) {
                    return null;
                }

                if (justification == WordParagraphAlignment.Center) {
                    return "center";
                }

                if (justification == WordParagraphAlignment.Right) {
                    return "right";
                }

                if (justification == WordParagraphAlignment.Left) {
                    return "left";
                }

                if (justification == WordParagraphAlignment.Both) {
                    return "justify";
                }

                return null;
            }

            WordParagraphAlignment? GetCellAlignment(WordTableCell cell) {
                WordParagraphAlignment? align = null;
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

            string? BuildBorderCss(string? style, string? colorHex, uint? size) {
                if (style == null) {
                    return null;
                }

                if (style == "nil") {
                    // Word nil wins a collapsed-border conflict against a visible neighbour.
                    return "hidden";
                }
                if (style == "none") {
                    // Word none yields to the opposing border in a collapsed conflict.
                    return "none";
                }

                string cssStyle = "solid";
                if (style == "dashed") {
                    cssStyle = "dashed";
                } else if (style == "dotted") {
                    cssStyle = "dotted";
                } else if (style == "double") {
                    cssStyle = "double";
                }

                string? normalizedColor = NormalizeSixDigitHexColor(colorHex);
                string color = normalizedColor != null ? $"#{normalizedColor}" : "black";
                double widthPt = size != null ? size.Value / 8.0 : 1.0;
                string width = widthPt.ToString("0.###", CultureInfo.InvariantCulture) + "pt";
                return $"{width} {cssStyle} {color}";
            }

            List<string> GetBorderCss(WordTableCell cell) {
                List<string> styles = new();
                string? ReadBorder(WordTableBorderSide side) {
                    OfficeXmlNode border = WordConverterNodes.ResolveTableCellBorderNode(cell, side);
                    return BuildBorderCss(border.WVal(), WordConverterNodes.ResolveBorderColorHex(border, cell.Document), border.WUInt32("sz"));
                }
                var left = ReadBorder(WordTableBorderSide.Left);
                var right = ReadBorder(WordTableBorderSide.Right);
                var top = ReadBorder(WordTableBorderSide.Top);
                var bottom = ReadBorder(WordTableBorderSide.Bottom);

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

                // Preserve native eighth-point quantization in the CSS point width.
                var left = BuildBorderCss(b.LeftStyle?.ToOpenXmlToken(), b.LeftColorValue, WordBorderLengths.ToSize(b.LeftSize));
                var right = BuildBorderCss(b.RightStyle?.ToOpenXmlToken(), b.RightColorValue, WordBorderLengths.ToSize(b.RightSize));
                var top = BuildBorderCss(b.TopStyle?.ToOpenXmlToken(), b.TopColorValue, WordBorderLengths.ToSize(b.TopSize));
                var bottom = BuildBorderCss(b.BottomStyle?.ToOpenXmlToken(), b.BottomColorValue, WordBorderLengths.ToSize(b.BottomSize));

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

    }
}
