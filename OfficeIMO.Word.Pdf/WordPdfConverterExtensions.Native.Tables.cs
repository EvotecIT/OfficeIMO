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
        private const int NativeOfficeImoScaffoldCellWidthTwips = 2400;

        private static void RenderNativeTable(INativePdfFlow pdf, WordTable table, Func<WordParagraph, (int Level, string Marker)?> getMarker, Dictionary<long, int> footnoteNumbersById, PdfSaveOptions? options, double? contentWidth, NativeDocumentDefaults nativeDefaults) {
            RecordNativeBodyTableDiagnostics(table, options, "body table");

            TableLayout layout = TableLayoutCache.GetLayout(table);
            bool hasExplicitDefaultTableStyle = options?.PdfOptions?.HasExplicitDefaultTableStyle == true;
            NativeTableStyleDefaults tableStyleDefaults = GetNativeTableStyleDefaults(
                table,
                nativeDefaults,
                ignoreFallbackTableStyle: hasExplicitDefaultTableStyle);
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

                    NativeCellText cellText = CreateNativeCellText(cell, footnoteNumbersById, nativeDefaults, tableStyleDefaults);
                    IReadOnlyList<PdfCore.PdfTableCellCheckBox> checkBoxes = CreateNativeTableCellCheckBoxes(cell);
                    IReadOnlyList<PdfCore.PdfTableCellFormField> formFields = CreateNativeTableCellFormFields(cell);
                    IReadOnlyList<PdfCore.PdfTableCellImage> images = CreateNativeTableCellImages(cell);
                    (string? LinkUri, string? LinkContents) link = GetNativeCellLink(cell);
                    int rowSpan = GetNativeCellRowSpan(cell);
                    nativeCells.Add(new PdfCore.PdfTableCell(
                        cellText.Runs,
                        cellText.Paragraphs,
                        columnSpan,
                        link.LinkUri,
                        link.LinkContents,
                        rowSpan,
                        checkBoxes.Count == 0 ? null : checkBoxes,
                        formFields.Count == 0 ? null : formFields,
                        images.Count == 0 ? null : images,
                        noWrap: !cell.WrapText));

                    PdfCore.PdfColor? fill = ParseNativeColor(cell.ShadingFillColorHex) ?? tableStyleDefaults.CellFill;
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

            PdfCore.PdfTableStyle style = CreateNativeTableStyle(table, rows.Count, options, contentWidth, nativeDefaults, tableStyleDefaults);
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

            ApplyNativeColumnWidths(table, layout, style, contentWidth);

            if (horizontalAlignments != null) {
                style.Alignments = horizontalAlignments;
            }

            if (verticalAlignments != null) {
                style.VerticalAlignments = verticalAlignments;
            }

            pdf.Table(rows, MapNativeTableAlignment(table.Alignment), style);
        }

        private static void ApplyNativeColumnWidths(WordTable table, TableLayout layout, PdfCore.PdfTableStyle style, double? contentWidth) {
            if (style.AutoFitColumns) {
                return;
            }

            style.ColumnWidthPoints = CreateNativeColumnWidthPoints(layout, style);
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

        private static PdfCore.PdfTableStyle CreateNativeTableStyle(WordTable table, int rowCount, PdfSaveOptions? options) =>
            CreateNativeTableStyle(table, rowCount, options, null);

        private static PdfCore.PdfTableStyle CreateNativeTableStyle(WordTable table, int rowCount, PdfSaveOptions? options, double? contentWidth) =>
            CreateNativeTableStyle(table, rowCount, options, contentWidth, NativeDocumentDefaults.WordDefault);

        private static PdfCore.PdfTableStyle CreateNativeTableStyle(WordTable table, int rowCount, PdfSaveOptions? options, double? contentWidth, NativeDocumentDefaults nativeDefaults) {
            bool hasExplicitDefaultTableStyle = options?.PdfOptions?.HasExplicitDefaultTableStyle == true;
            NativeTableStyleDefaults tableStyleDefaults = GetNativeTableStyleDefaults(
                table,
                nativeDefaults,
                ignoreFallbackTableStyle: hasExplicitDefaultTableStyle);
            return CreateNativeTableStyle(table, rowCount, options, contentWidth, nativeDefaults, tableStyleDefaults);
        }

        private static PdfCore.PdfTableStyle CreateNativeTableStyle(WordTable table, int rowCount, PdfSaveOptions? options, double? contentWidth, NativeDocumentDefaults nativeDefaults, NativeTableStyleDefaults tableStyleDefaults) {
            bool hasExplicitDefaultTableStyle = options?.PdfOptions?.HasExplicitDefaultTableStyle == true;
            PdfCore.PdfTableStyle? wordStyle = ResolveNativeWordTableStyle(table, hasExplicitDefaultTableStyle);
            bool usesConfiguredDefaultStyle = wordStyle == null && hasExplicitDefaultTableStyle;
            PdfCore.PdfTableStyle style = wordStyle ?? CreateNativeDefaultTableStyle(options);
            if (!usesConfiguredDefaultStyle) {
                style.FontSize ??= nativeDefaults.FontSize;
                double? tableParagraphLineHeight = ShouldApplyNativeTableStyleParagraphLineHeight(table)
                    ? ResolveNativeTableStyleParagraphLineHeight(tableStyleDefaults, style.FontSize ?? nativeDefaults.FontSize)
                    : null;
                style.LineHeight ??= tableParagraphLineHeight ?? nativeDefaults.ParagraphLineHeight;
            }

            int repeatedHeaderRowCount = GetNativeTableRepeatedHeaderRowCount(table, rowCount);
            style.HeaderRowCount = GetNativeTableVisualHeaderRowCount(table, rowCount, repeatedHeaderRowCount);
            style.RepeatHeaderRowCount = repeatedHeaderRowCount;
            if (repeatedHeaderRowCount > 0) {
                style.PageContinuationSpacingBefore = Math.Max(style.PageContinuationSpacingBefore, NativeTablePageContinuationSpacingBefore);
            }

            if (options?.DefaultTableBorders == true && style.BorderColor == null) {
                style.BorderColor = PdfCore.PdfColor.LightGray;
            }

            ApplyNativeTableBorders(table, style, tableStyleDefaults);
            ApplyNativeTableDefaultCellMargins(
                table,
                style,
                usesConfiguredDefaultStyle,
                ShouldApplyNativeTableStyleCellPadding(table) ? tableStyleDefaults : NativeTableStyleDefaults.Empty);
            ApplyNativeTableLayoutOptions(table, style, contentWidth);
            ApplyNativeTableRowOptions(table, style);
            return style;
        }

        private static double? ResolveNativeTableStyleParagraphLineHeight(NativeTableStyleDefaults tableStyleDefaults, double fontSize) {
            if (tableStyleDefaults.ParagraphLineSpacingPoints.HasValue && fontSize > 0D) {
                return ResolveNativeLineSpacingHeight(
                    tableStyleDefaults.ParagraphLineSpacingPoints.Value,
                    tableStyleDefaults.ParagraphLineSpacingRule,
                    fontSize,
                    NativeWordTableSingleLineHeight);
            }

            return tableStyleDefaults.ParagraphLineHeight;
        }

        private static PdfCore.PdfTableStyle CreateNativeDefaultTableStyle(PdfSaveOptions? options) {
            PdfCore.PdfTableStyle? configuredStyle = options?.PdfOptions?.HasExplicitDefaultTableStyle == true
                ? options.PdfOptions.DefaultTableStyle
                : null;
            if (configuredStyle != null) {
                return configuredStyle.Clone();
            }

            return new PdfCore.PdfTableStyle {
                RowStripeFill = null
            };
        }

        private static void ApplyNativeTableLayoutOptions(WordTable table, PdfCore.PdfTableStyle style, double? contentWidth) {
            W.TableProperties? properties = table._tableProperties;
            if (IsNativeTableAutoFitLayout(properties) &&
                (IsNativeExplicitAutoFitTableLayout(properties) || !HasNativeTableAuthoredFixedCellWidths(table))) {
                style.AutoFitColumns = true;
            }

            double? maxWidth = GetNativeTablePreferredWidth(properties?.TableWidth, contentWidth);
            if (maxWidth.HasValue) {
                style.MaxWidth = maxWidth.Value;
                style.PreserveWidth = true;
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

        private static bool HasNativeTableAuthoredFixedCellWidths(WordTable table) {
            foreach (WordTableRow row in table.Rows) {
                foreach (WordTableCell cell in row.Cells) {
                    if (cell.Width.GetValueOrDefault() > 0 &&
                        cell.WidthType == W.TableWidthUnitValues.Dxa &&
                        cell.Width.GetValueOrDefault() != NativeOfficeImoScaffoldCellWidthTwips) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool IsNativeExplicitAutoFitTableLayout(W.TableProperties? properties) =>
            properties?.TableLayout?.Type?.Value == W.TableLayoutValues.Autofit;

        private static bool IsNativeTableAutoFitToContents(W.TableProperties? properties) =>
            IsNativeTableAutoFitLayout(properties) &&
            properties?.TableWidth?.Type?.Value == W.TableWidthUnitValues.Auto;

        private static bool IsNativeTableAutoFitLayout(W.TableProperties? properties) {
            if (properties?.TableLayout?.Type?.Value == W.TableLayoutValues.Autofit) {
                return true;
            }

            if (properties?.TableLayout?.Type?.Value == W.TableLayoutValues.Fixed) {
                return false;
            }

            return properties?.TableWidth?.Type?.Value == W.TableWidthUnitValues.Auto;
        }

        private static double? GetNativeTablePreferredWidth(W.TableWidth? width, double? contentWidth) {
            if (width?.Type?.Value == W.TableWidthUnitValues.Pct) {
                double? percent = GetNativeTablePreferredWidthPercent(width);
                if (!percent.HasValue || !contentWidth.HasValue || contentWidth.Value <= 0D) {
                    return null;
                }

                return contentWidth.Value * percent.Value;
            }

            if (width?.Type?.Value != W.TableWidthUnitValues.Dxa) {
                return null;
            }

            return ConvertNativeTwipsToPoints(width.Width?.Value);
        }

        private static double? GetNativeTablePreferredWidthPercent(W.TableWidth width) {
            string? rawWidth = width.Width?.Value;
            if (string.IsNullOrWhiteSpace(rawWidth)) {
                return null;
            }

            string valueText = rawWidth!.Trim();
            if (valueText.EndsWith("%", StringComparison.Ordinal)) {
                string percentText = valueText.Substring(0, valueText.Length - 1);
                if (!double.TryParse(percentText, NumberStyles.Float, CultureInfo.InvariantCulture, out double percent) ||
                    percent <= 0D ||
                    double.IsNaN(percent) ||
                    double.IsInfinity(percent)) {
                    return null;
                }

                return percent / 100D;
            }

            if (!int.TryParse(valueText, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value) || value <= 0) {
                return null;
            }

            return value / 5000D;
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

        private static void ApplyNativeTableBorders(WordTable table, PdfCore.PdfTableStyle style, NativeTableStyleDefaults tableStyleDefaults) {
            (PdfCore.PdfColor Color, double Width)? border = GetNativeUniformTableBorder(table._tableProperties?.TableBorders) ?? tableStyleDefaults.TableBorder;
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

        private static void ApplyNativeTableDefaultCellMargins(WordTable table, PdfCore.PdfTableStyle style, bool preserveConfiguredFallbackPadding, NativeTableStyleDefaults tableStyleDefaults) {
            W.TableCellMarginDefault? margins = table._tableProperties?.TableCellMarginDefault;
            if (margins == null) {
                if (tableStyleDefaults.CellPadding != null) {
                    ApplyNativeResolvedTableCellPadding(style, tableStyleDefaults.CellPadding);
                }

                if (!preserveConfiguredFallbackPadding) {
                    style.CellPaddingTop ??= 3D;
                    style.CellPaddingBottom ??= 3D;
                }

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

            if (top.HasValue) {
                style.CellPaddingTop = top.Value;
            } else if (!preserveConfiguredFallbackPadding) {
                style.CellPaddingTop = 3D;
            }

            if (bottom.HasValue) {
                style.CellPaddingBottom = bottom.Value;
            } else if (!preserveConfiguredFallbackPadding) {
                style.CellPaddingBottom = 3D;
            }

            if (left.HasValue) {
                style.CellPaddingLeft = left.Value;
            }

            if (right.HasValue) {
                style.CellPaddingRight = right.Value;
            }
        }

        private static void ApplyNativeResolvedTableCellPadding(PdfCore.PdfTableStyle style, PdfCore.PdfCellPadding padding) {
            if (padding.Top.HasValue) {
                style.CellPaddingTop = padding.Top.Value;
            }

            if (padding.Bottom.HasValue) {
                style.CellPaddingBottom = padding.Bottom.Value;
            }

            if (padding.Left.HasValue) {
                style.CellPaddingLeft = padding.Left.Value;
            }

            if (padding.Right.HasValue) {
                style.CellPaddingRight = padding.Right.Value;
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

        private static PdfCore.PdfTableStyle? ResolveNativeWordTableStyle(WordTable table, bool preferConfiguredDefaultStyle) {
            string? wordStyle = GetNativeTableStyleId(table);
            if (string.IsNullOrWhiteSpace(wordStyle)) {
                return null;
            }

            if (preferConfiguredDefaultStyle && IsNativeFallbackTableStyleId(wordStyle)) {
                return null;
            }

            return PdfCore.TableStyles.TryFromWordTableStyle(wordStyle!, out PdfCore.PdfTableStyle? style)
                ? style
                : null;
        }

        private static bool ShouldApplyNativeTableStyleParagraphLineHeight(WordTable table) {
            string? styleId = GetNativeTableStyleId(table);
            if (string.IsNullOrWhiteSpace(styleId)) {
                return false;
            }

            if (!PdfCore.TableStyles.TryGetCanonicalWordStyleName(styleId!, out string? canonicalStyleName)) {
                return true;
            }

            return string.Equals(canonicalStyleName, "TableGrid", StringComparison.OrdinalIgnoreCase);
        }

        private static bool ShouldApplyNativeTableStyleCellPadding(WordTable table) {
            string? styleId = GetNativeTableStyleId(table);
            if (string.IsNullOrWhiteSpace(styleId)) {
                return false;
            }

            if (!PdfCore.TableStyles.TryGetCanonicalWordStyleName(styleId!, out string? canonicalStyleName)) {
                return true;
            }

            return string.Equals(canonicalStyleName, "TableGrid", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(canonicalStyleName, "TableNormal", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(canonicalStyleName, "PlainTable1", StringComparison.OrdinalIgnoreCase);
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

                PdfCore.PdfColumnAlign paragraphAlignment = ResolveNativeColumnAlign(paragraph);
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
