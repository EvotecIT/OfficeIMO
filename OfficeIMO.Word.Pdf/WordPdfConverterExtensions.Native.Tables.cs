using System.Collections.Generic;
using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static void RenderNativeTable(INativePdfFlow pdf, WordTable table, Func<WordParagraph, (int Level, string Marker)?> getMarker, Dictionary<long, int> footnoteNumbersById, PdfSaveOptions? options) {
            RecordNativeBodyTableDiagnostics(table, options, "body table");

            TableLayout layout = TableLayoutCache.GetLayout(table);
            var rows = new List<PdfCore.PdfTableCell[]>();
            var cellFills = new Dictionary<(int Row, int Column), PdfCore.PdfColor>();
            var cellBorders = new Dictionary<(int Row, int Column), PdfCore.PdfCellBorder>();
            var cellPaddings = new Dictionary<(int Row, int Column), PdfCore.PdfCellPadding>();
            var cellAlignments = new Dictionary<(int Row, int Column), PdfCore.PdfColumnAlign>();
            var cellVerticalAlignments = new Dictionary<(int Row, int Column), PdfCore.PdfCellVerticalAlign>();
            var horizontalAlignments = CreateNativeTableHorizontalAlignments(layout);
            var verticalAlignments = CreateNativeTableVerticalAlignments(layout);
            for (int rowIndex = 0; rowIndex < layout.Rows.Count; rowIndex++) {
                IReadOnlyList<WordTableCell> row = layout.Rows[rowIndex];
                var nativeCells = new List<PdfCore.PdfTableCell>();
                int logicalColumnIndex = 0;
                for (int columnIndex = 0; columnIndex < row.Count; columnIndex++) {
                    WordTableCell cell = row[columnIndex];
                    if (IsNativeHorizontalMergeContinuation(cell)) {
                        continue;
                    }

                    int columnSpan = GetNativeCellColumnSpan(cell);
                    if (IsNativeVerticalMergeContinuation(cell)) {
                        logicalColumnIndex += columnSpan;
                        continue;
                    }

                    IReadOnlyList<PdfCore.TextRun> cellRuns = CreateNativeCellRuns(cell, footnoteNumbersById);
                    IReadOnlyList<PdfCore.PdfTableCellCheckBox> checkBoxes = CreateNativeTableCellCheckBoxes(cell);
                    IReadOnlyList<PdfCore.PdfTableCellFormField> formFields = CreateNativeTableCellFormFields(cell);
                    IReadOnlyList<PdfCore.PdfTableCellImage> images = CreateNativeTableCellImages(cell);
                    (string? LinkUri, string? LinkContents) link = GetNativeCellLink(cell);
                    int rowSpan = GetNativeCellRowSpan(cell);
                    nativeCells.Add(new PdfCore.PdfTableCell(
                        cellRuns,
                        columnSpan,
                        link.LinkUri,
                        link.LinkContents,
                        rowSpan,
                        checkBoxes.Count == 0 ? null : checkBoxes,
                        formFields.Count == 0 ? null : formFields,
                        images.Count == 0 ? null : images));

                    PdfCore.PdfColor? fill = ParseNativeColor(cell.ShadingFillColorHex);
                    if (fill.HasValue) {
                        cellFills[(rowIndex, logicalColumnIndex)] = fill.Value;
                    }

                    PdfCore.PdfCellBorder? border = CreateNativeTableCellBorder(cell.Borders);
                    if (border != null) {
                        cellBorders[(rowIndex, logicalColumnIndex)] = border;
                    }

                    PdfCore.PdfCellPadding? padding = CreateNativeTableCellPadding(cell);
                    if (padding != null) {
                        cellPaddings[(rowIndex, logicalColumnIndex)] = padding;
                    }

                    PdfCore.PdfColumnAlign cellAlignment = GetNativeCellHorizontalAlignment(cell);
                    if (cellAlignment != PdfCore.PdfColumnAlign.Left) {
                        cellAlignments[(rowIndex, logicalColumnIndex)] = cellAlignment;
                    }

                    PdfCore.PdfCellVerticalAlign cellVerticalAlignment = MapNativeCellVerticalAlign(cell.VerticalAlignment);
                    if (cellVerticalAlignment != PdfCore.PdfCellVerticalAlign.Top) {
                        cellVerticalAlignments[(rowIndex, logicalColumnIndex)] = cellVerticalAlignment;
                    }

                    logicalColumnIndex += columnSpan;
                }

                rows.Add(nativeCells.ToArray());
            }

            if (rows.Count == 0) {
                return;
            }

            PdfCore.PdfTableStyle style = CreateNativeTableStyle(table, rows.Count, options);
            if (cellFills.Count > 0) {
                style.CellFills = cellFills;
            }

            if (cellBorders.Count > 0) {
                style.CellBorders = cellBorders;
            }

            if (cellPaddings.Count > 0) {
                style.CellPaddings = cellPaddings;
            }

            if (cellAlignments.Count > 0) {
                style.CellAlignments = cellAlignments;
            }

            if (cellVerticalAlignments.Count > 0) {
                style.CellVerticalAlignments = cellVerticalAlignments;
            }

            style.ColumnWidthPoints = CreateNativeColumnWidthPoints(layout, style);

            if (horizontalAlignments != null) {
                style.Alignments = horizontalAlignments;
            }

            if (verticalAlignments != null) {
                style.VerticalAlignments = verticalAlignments;
            }

            pdf.Table(rows, MapNativeTableAlignment(table.Alignment), style);
        }

        private static List<double?>? CreateNativeColumnWidthPoints(TableLayout layout, PdfCore.PdfTableStyle style) {
            if (style.AutoFitColumns || layout.ColumnWidths.Length == 0 || !layout.ColumnWidths.All(width => width > 0)) {
                return null;
            }

            var widths = layout.ColumnWidths.Select(width => (double)width).ToList();
            double totalWidth = widths.Sum();
            if (style.MaxWidth.HasValue && totalWidth > style.MaxWidth.Value + 0.001D) {
                double scale = style.MaxWidth.Value / totalWidth;
                for (int i = 0; i < widths.Count; i++) {
                    widths[i] *= scale;
                }
            }

            return widths.Select(width => (double?)width).ToList();
        }

        private static PdfCore.PdfTableStyle CreateNativeTableStyle(WordTable table, int rowCount, PdfSaveOptions? options) {
            PdfCore.PdfTableStyle style = ResolveNativeWordTableStyle(table) ?? new PdfCore.PdfTableStyle {
                RowStripeFill = null
            };
            style.FontSize ??= 10D;
            style.LineHeight ??= 1.15D;

            int repeatedHeaderRowCount = GetNativeTableRepeatedHeaderRowCount(table, rowCount);
            style.HeaderRowCount = GetNativeTableVisualHeaderRowCount(table, rowCount, repeatedHeaderRowCount);
            style.RepeatHeaderRowCount = repeatedHeaderRowCount;
            if (options?.DefaultTableBorders == true && style.BorderColor == null) {
                style.BorderColor = PdfCore.PdfColor.LightGray;
            }

            ApplyNativeTableBorders(table, style);
            ApplyNativeTableDefaultCellMargins(table, style);
            ApplyNativeTableLayoutOptions(table, style);
            ApplyNativeTableRowOptions(table, style);
            return style;
        }

        private static void ApplyNativeTableLayoutOptions(WordTable table, PdfCore.PdfTableStyle style) {
            W.TableProperties? properties = table._tableProperties;
            if (IsNativeTableAutoFitToContents(properties)) {
                style.AutoFitColumns = true;
            }

            double? maxWidth = GetNativeTablePreferredWidth(properties?.TableWidth);
            if (maxWidth.HasValue) {
                style.MaxWidth = maxWidth.Value;
            }

            double? leftIndent = GetNativeTableLeftIndent(properties?.TableIndentation);
            if (leftIndent.HasValue) {
                style.LeftIndent = leftIndent.Value;
            }

            double? cellSpacing = GetNativeTableCellSpacing(properties?.TableCellSpacing);
            if (cellSpacing.HasValue) {
                style.CellSpacing = cellSpacing.Value;
            }
        }

        private static bool IsNativeTableAutoFitToContents(W.TableProperties? properties) =>
            properties?.TableLayout?.Type?.Value == W.TableLayoutValues.Autofit &&
            properties.TableWidth?.Type?.Value == W.TableWidthUnitValues.Auto;

        private static double? GetNativeTablePreferredWidth(W.TableWidth? width) {
            if (width?.Type?.Value != W.TableWidthUnitValues.Dxa) {
                return null;
            }

            return ConvertNativeTwipsToPoints(width.Width?.Value);
        }

        private static double? GetNativeTableLeftIndent(W.TableIndentation? indentation) {
            if (indentation?.Type?.Value != W.TableWidthUnitValues.Dxa || indentation.Width == null) {
                return null;
            }

            return ConvertNativeTwipsToPoints(indentation.Width.Value);
        }

        private static double? GetNativeTableCellSpacing(W.TableCellSpacing? spacing) {
            if (spacing?.Type?.Value != W.TableWidthUnitValues.Dxa) {
                return null;
            }

            return ConvertNativeTwipsToPoints(spacing.Width?.Value);
        }

        private static void ApplyNativeTableBorders(WordTable table, PdfCore.PdfTableStyle style) {
            (PdfCore.PdfColor Color, double Width)? border = GetNativeUniformTableBorder(table._tableProperties?.TableBorders);
            if (border == null) {
                return;
            }

            style.BorderColor = border.Value.Color;
            style.BorderWidth = border.Value.Width;
        }

        private static (PdfCore.PdfColor Color, double Width)? GetNativeUniformTableBorder(W.TableBorders? borders) {
            if (borders == null) {
                return null;
            }

            W.BorderType?[] allBorders = {
                borders.TopBorder,
                borders.BottomBorder,
                borders.LeftBorder,
                borders.RightBorder,
                borders.InsideHorizontalBorder,
                borders.InsideVerticalBorder
            };

            if (allBorders.Any(border => border == null || !HasNativeBorder(border.Val?.Value))) {
                return null;
            }

            W.BorderValues style = allBorders[0]!.Val!.Value;
            if (allBorders.Any(border => border!.Val?.Value != style)) {
                return null;
            }

            uint size = allBorders[0]!.Size?.Value ?? 4U;
            if (allBorders.Any(border => (border!.Size?.Value ?? 4U) != size)) {
                return null;
            }

            string? color = NormalizeNativeBorderColor(allBorders[0]!.Color?.Value);
            if (allBorders.Any(border => !string.Equals(color, NormalizeNativeBorderColor(border!.Color?.Value), StringComparison.OrdinalIgnoreCase))) {
                return null;
            }

            return (ParseNativeColor(color) ?? PdfCore.PdfColor.Black, size / 8D);
        }

        private static void ApplyNativeTableDefaultCellMargins(WordTable table, PdfCore.PdfTableStyle style) {
            W.TableCellMarginDefault? margins = table._tableProperties?.TableCellMarginDefault;
            if (margins == null) {
                style.CellPaddingTop = 3D;
                style.CellPaddingBottom = 3D;
                return;
            }

            double? top = ConvertNativeTwipsToPoints(margins.TopMargin?.Width?.Value);
            double? bottom = ConvertNativeTwipsToPoints(margins.BottomMargin?.Width?.Value);
            double? left = margins.TableCellLeftMargin?.Width == null
                ? null
                : ConvertNativeTwipsToPoints(margins.TableCellLeftMargin.Width.Value);
            double? right = margins.TableCellRightMargin?.Width == null
                ? null
                : ConvertNativeTwipsToPoints(margins.TableCellRightMargin.Width.Value);

            style.CellPaddingTop = top ?? 3D;
            style.CellPaddingBottom = bottom ?? 3D;

            if (left.HasValue) {
                style.CellPaddingLeft = left.Value;
            }

            if (right.HasValue) {
                style.CellPaddingRight = right.Value;
            }
        }

        private static PdfCore.PdfCellPadding? CreateNativeTableCellPadding(WordTableCell cell) {
            double? top = cell.MarginTopWidth.HasValue ? ConvertNativeTwipsToPoints(cell.MarginTopWidth.Value) : null;
            double? bottom = cell.MarginBottomWidth.HasValue ? ConvertNativeTwipsToPoints(cell.MarginBottomWidth.Value) : null;
            double? left = cell.MarginLeftWidth.HasValue ? ConvertNativeTwipsToPoints(cell.MarginLeftWidth.Value) : null;
            double? right = cell.MarginRightWidth.HasValue ? ConvertNativeTwipsToPoints(cell.MarginRightWidth.Value) : null;
            if (!top.HasValue && !bottom.HasValue && !left.HasValue && !right.HasValue) {
                return null;
            }

            return new PdfCore.PdfCellPadding {
                Top = top,
                Bottom = bottom,
                Left = left,
                Right = right
            };
        }

        private static double? ConvertNativeTwipsToPoints(string? value) {
            if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int twips) || twips < 0) {
                return null;
            }

            return twips / 20D;
        }

        private static double? ConvertNativeTwipsToPoints(int twips) {
            return twips < 0 ? null : twips / 20D;
        }

        private static double ConvertNativeEmusToPoints(long emus) {
            return emus <= 0 ? 0D : emus / 12700D;
        }

        private static void ApplyNativeTableRowOptions(WordTable table, PdfCore.PdfTableStyle style) {
            style.AllowRowBreakAcrossPages = table.AllowRowToBreakAcrossPages;
            List<bool?>? rowBreakPolicies = GetNativeTableRowBreakPolicies(table);
            if (rowBreakPolicies != null) {
                style.RowAllowBreakAcrossPages = rowBreakPolicies;
            }

            List<double?>? rowHeights = GetNativeTableRowHeights(table);
            if (rowHeights == null) {
                return;
            }

            double? uniformHeight = GetNativeUniformTableRowHeight(rowHeights);
            if (uniformHeight.HasValue) {
                style.MinRowHeight = uniformHeight.Value;
            } else {
                style.RowMinHeights = rowHeights;
            }
        }

        private static List<bool?>? GetNativeTableRowBreakPolicies(WordTable table) {
            var policies = new List<bool?>(table.Rows.Count);
            bool? firstPolicy = null;
            bool hasMixedPolicies = false;
            foreach (WordTableRow row in table.Rows) {
                bool policy = row.AllowRowToBreakAcrossPages;
                policies.Add(policy);
                if (!firstPolicy.HasValue) {
                    firstPolicy = policy;
                    continue;
                }

                hasMixedPolicies |= firstPolicy.Value != policy;
            }

            return hasMixedPolicies ? policies : null;
        }

        private static List<double?>? GetNativeTableRowHeights(WordTable table) {
            var heights = new List<double?>(table.Rows.Count);
            bool hasHeight = false;
            foreach (WordTableRow row in table.Rows) {
                double? height = row.Height.HasValue && row.Height.Value > 0
                    ? ConvertNativeTwipsToPoints(row.Height.Value)
                    : null;
                heights.Add(height);
                hasHeight |= height.HasValue;
            }

            return hasHeight ? heights : null;
        }

        private static double? GetNativeUniformTableRowHeight(IReadOnlyList<double?> rowHeights) {
            double? height = null;
            foreach (double? rowHeight in rowHeights) {
                if (!rowHeight.HasValue) {
                    return null;
                }

                if (!height.HasValue) {
                    height = rowHeight.Value;
                    continue;
                }

                if (System.Math.Abs(height.Value - rowHeight.Value) > 0.001D) {
                    return null;
                }
            }

            return height;
        }

        private static PdfCore.PdfTableStyle? ResolveNativeWordTableStyle(WordTable table) {
            WordTableStyle? wordStyle = table.Style;
            if (!wordStyle.HasValue) {
                return null;
            }

            return PdfCore.TableStyles.TryFromWordTableStyle(wordStyle.Value.ToString(), out PdfCore.PdfTableStyle? style)
                ? style
                : null;
        }

        private static int GetNativeTableVisualHeaderRowCount(WordTable table, int rowCount, int repeatedHeaderRowCount) {
            if (rowCount == 0) {
                return 0;
            }

            int headerRowCount = repeatedHeaderRowCount;
            if (table.ConditionalFormattingFirstRow == true || headerRowCount > 0) {
                headerRowCount = System.Math.Max(headerRowCount, 1);
            }

            return System.Math.Min(headerRowCount, rowCount);
        }

        private static int GetNativeTableRepeatedHeaderRowCount(WordTable table, int rowCount) {
            if (rowCount == 0 || table.Rows.Count == 0) {
                return 0;
            }

            int repeatedHeaderRowCount = 0;
            foreach (WordTableRow row in table.Rows) {
                if (!row.RepeatHeaderRowAtTheTopOfEachPage) {
                    break;
                }

                repeatedHeaderRowCount++;
                if (repeatedHeaderRowCount == rowCount) {
                    break;
                }
            }

            return repeatedHeaderRowCount;
        }

        private static PdfCore.PdfAlign MapNativeTableAlignment(W.TableRowAlignmentValues? alignment) {
            if (alignment == W.TableRowAlignmentValues.Center) {
                return PdfCore.PdfAlign.Center;
            }

            if (alignment == W.TableRowAlignmentValues.Right) {
                return PdfCore.PdfAlign.Right;
            }

            return PdfCore.PdfAlign.Left;
        }

        private static List<PdfCore.PdfColumnAlign>? CreateNativeTableHorizontalAlignments(TableLayout layout) {
            int columnCount = GetNativeTableColumnCount(layout);
            if (columnCount == 0) {
                return null;
            }

            var alignments = new List<PdfCore.PdfColumnAlign>(columnCount);
            bool hasExplicitAlignment = false;
            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                PdfCore.PdfColumnAlign? columnAlignment = null;
                bool conflict = false;
                foreach ((WordTableCell Cell, int Column, int ColumnSpan) cell in EnumerateNativeTableCells(layout)) {
                    if (columnIndex < cell.Column || columnIndex >= cell.Column + cell.ColumnSpan) {
                        continue;
                    }

                    PdfCore.PdfColumnAlign alignment = GetNativeCellHorizontalAlignment(cell.Cell);
                    if (columnAlignment == null) {
                        columnAlignment = alignment;
                    } else if (columnAlignment.Value != alignment) {
                        conflict = true;
                        break;
                    }
                }

                PdfCore.PdfColumnAlign resolved = conflict ? PdfCore.PdfColumnAlign.Left : columnAlignment ?? PdfCore.PdfColumnAlign.Left;
                if (resolved != PdfCore.PdfColumnAlign.Left) {
                    hasExplicitAlignment = true;
                }

                alignments.Add(resolved);
            }

            return hasExplicitAlignment ? alignments : null;
        }

        private static PdfCore.PdfColumnAlign GetNativeCellHorizontalAlignment(WordTableCell cell) {
            PdfCore.PdfColumnAlign? alignment = null;
            foreach (WordParagraph paragraph in cell.Paragraphs) {
                string text = GetNativeCellParagraphText(paragraph);
                if (string.IsNullOrWhiteSpace(text)) {
                    continue;
                }

                PdfCore.PdfColumnAlign paragraphAlignment = MapNativeColumnAlign(paragraph.ParagraphAlignment);
                if (alignment == null) {
                    alignment = paragraphAlignment;
                } else if (alignment.Value != paragraphAlignment) {
                    return PdfCore.PdfColumnAlign.Left;
                }
            }

            return alignment ?? PdfCore.PdfColumnAlign.Left;
        }

    }
}
