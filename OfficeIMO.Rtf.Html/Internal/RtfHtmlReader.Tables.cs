using System.Globalization;

namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private void StartTable() {
            _table = _document.AddTable(0, 1);
            _row = null;
            _cell = null;
            _rowSpans.Clear();
            _paragraph = null;
        }

        private void StartRow() {
            StartRow(null, HtmlStyleDeclaration.Empty);
        }

        private void StartRow(HtmlToken? token, HtmlStyleDeclaration style) {
            if (_table == null) {
                StartTable();
            }

            _row = _table!.AddRow();
            _row.RepeatHeader = _tableHead > 0;
            _tableColumnIndex = 0;
            _cell = null;
            _cellTextAlignment = null;
            _paragraph = null;
            ApplyRowStyle(style);
            if (token != null) {
                ApplyRowAttributes(token);
            }
        }

        private void StartCell(HtmlToken token, HtmlStyleDeclaration style, bool isHeader) {
            if (_row == null) {
                StartRow();
            }

            AddPendingRowSpanContinuations();

            int columnStart = _tableColumnIndex;
            int columnSpan = ReadSpan(token, "colspan");
            int rowSpan = ReadSpan(token, "rowspan");
            _cell = AddCellAtColumn(columnStart);
            if (columnSpan > 1) {
                _cell.HorizontalMerge = RtfTableCellMerge.First;
            }

            if (rowSpan > 1) {
                _cell.VerticalMerge = RtfTableCellMerge.First;
            }

            for (int offset = 1; offset < columnSpan; offset++) {
                RtfTableCell continuation = AddCellAtColumn(columnStart + offset);
                continuation.HorizontalMerge = RtfTableCellMerge.Continue;
                if (rowSpan > 1) {
                    continuation.VerticalMerge = RtfTableCellMerge.First;
                }
            }

            if (rowSpan > 1) {
                TrackRowSpan(columnStart, columnSpan, rowSpan);
            }

            _tableColumnIndex += columnSpan;
            _cellTextAlignment = isHeader ? RtfTextAlignment.Center : null;
            ApplyCellStyle(style);
            ApplyCellAttributes(token);
            if (isHeader) {
                _bold++;
            }

            _paragraph = null;
        }

        private void EndRow() {
            AddPendingRowSpanContinuations(includeTrailingColumns: true);
            _tableColumnIndex = 0;
            _cellTextAlignment = null;
            _paragraph = null;
        }

        private RtfTableCell AddCellAtColumn(int zeroBasedColumn) {
            return _row!.AddCell((zeroBasedColumn + 1) * 2400);
        }

        private void AddPendingRowSpanContinuations(bool includeTrailingColumns = false) {
            while (_tableColumnIndex < _rowSpans.Count) {
                RowSpanState state = _rowSpans[_tableColumnIndex];
                if (state.RemainingRows <= 0) {
                    if (!includeTrailingColumns) {
                        break;
                    }

                    _tableColumnIndex++;
                    continue;
                }

                RtfTableCell continuation = AddCellAtColumn(_tableColumnIndex);
                continuation.VerticalMerge = RtfTableCellMerge.Continue;
                if (state.ColumnSpan > 1) {
                    continuation.HorizontalMerge = state.Offset == 0
                        ? RtfTableCellMerge.First
                        : RtfTableCellMerge.Continue;
                }

                state.RemainingRows--;
                _tableColumnIndex++;
            }
        }

        private void ApplyRowStyle(HtmlStyleDeclaration style) {
            if (_row == null) {
                return;
            }

            if (style.BackgroundColor != null) {
                _row.BackgroundColorIndex = GetOrAddColorIndex(style.BackgroundColor);
            }

            if (style.ShadingForegroundColor != null) {
                _row.ShadingForegroundColorIndex = GetOrAddColorIndex(style.ShadingForegroundColor);
            }

            if (style.ShadingPatternValue.HasValue) {
                _row.ShadingPatternValue = style.ShadingPatternValue.Value;
            }

            if (style.ShadingPatternPercent.HasValue) {
                _row.ShadingPatternPercent = style.ShadingPatternPercent.Value;
            }

            if (style.ShadingPattern.HasValue) {
                _row.ShadingPattern = style.ShadingPattern.Value;
            }

            if (style.TextAlignment.HasValue && TryMapTableAlignment(style.TextAlignment.Value, out RtfTableAlignment alignment)) {
                _row.Alignment = alignment;
            }

            if (style.Direction.HasValue) {
                _row.Direction = MapTableRowDirection(style.Direction.Value);
            }

            if (style.TableWidth.HasValue) {
                _row.PreferredWidth = style.TableWidth.Value;
                _row.PreferredWidthUnit = style.TableWidthUnit;
            }

            if (style.TableHeightTwips.HasValue) {
                _row.HeightTwips = style.TableHeightTwips.Value;
            }

            if (style.PaddingTopTwips.HasValue ||
                style.PaddingLeftTwips.HasValue ||
                style.PaddingBottomTwips.HasValue ||
                style.PaddingRightTwips.HasValue) {
                _row.SetPadding(style.PaddingTopTwips, style.PaddingLeftTwips, style.PaddingBottomTwips, style.PaddingRightTwips);
            }
        }

        private void ApplyRowAttributes(HtmlToken token) {
            if (_row == null) {
                return;
            }

            string? align = GetAttribute(token, "align");
            if (!string.IsNullOrWhiteSpace(align) && TryParseTableAlign(align!, out RtfTableAlignment alignment)) {
                _row.Alignment = alignment;
            }

            string? background = GetAttribute(token, "bgcolor");
            if (!string.IsNullOrWhiteSpace(background) && HtmlStyleDeclarationParser.TryParseColor(background!, out RtfColor? color)) {
                _row.BackgroundColorIndex = GetOrAddColorIndex(color!);
            }

            string? width = GetAttribute(token, "width");
            if (!string.IsNullOrWhiteSpace(width) && HtmlStyleDeclarationParser.TryParseTableWidth(width!, out int tableWidth, out RtfTableWidthUnit widthUnit)) {
                _row.PreferredWidth = tableWidth;
                _row.PreferredWidthUnit = widthUnit;
            }

            string? height = GetAttribute(token, "height");
            if (!string.IsNullOrWhiteSpace(height) && HtmlStyleDeclarationParser.TryParseTwips(height!, out int heightTwips)) {
                _row.HeightTwips = heightTwips;
            }
        }

        private void TrackRowSpan(int columnStart, int columnSpan, int rowSpan) {
            while (_rowSpans.Count < columnStart + columnSpan) {
                _rowSpans.Add(new RowSpanState());
            }

            for (int offset = 0; offset < columnSpan; offset++) {
                RowSpanState state = _rowSpans[columnStart + offset];
                state.RemainingRows = Math.Max(state.RemainingRows, rowSpan - 1);
                state.ColumnSpan = columnSpan;
                state.Offset = offset;
            }
        }

        private void ApplyCellStyle(HtmlStyleDeclaration style) {
            if (_cell == null) {
                return;
            }

            if (style.TextAlignment.HasValue) {
                _cellTextAlignment = style.TextAlignment.Value;
            }

            if (style.BackgroundColor != null) {
                _cell.BackgroundColorIndex = GetOrAddColorIndex(style.BackgroundColor);
            }

            if (style.ShadingForegroundColor != null) {
                _cell.ShadingForegroundColorIndex = GetOrAddColorIndex(style.ShadingForegroundColor);
            }

            if (style.ShadingPatternPercent.HasValue) {
                _cell.ShadingPatternPercent = style.ShadingPatternPercent.Value;
            }

            if (style.ShadingPattern.HasValue) {
                _cell.ShadingPattern = style.ShadingPattern.Value;
            }

            if (style.TableCellVerticalAlignment.HasValue) {
                _cell.VerticalAlignment = style.TableCellVerticalAlignment.Value;
            }

            if (style.TableCellTextFlow.HasValue) {
                _cell.TextFlow = style.TableCellTextFlow.Value;
            }

            if (style.TableWidth.HasValue) {
                _cell.PreferredWidth = style.TableWidth.Value;
                _cell.PreferredWidthUnit = style.TableWidthUnit;
            }

            if (style.NoWrap.HasValue) {
                _cell.NoWrap = style.NoWrap.Value;
            }

            if (style.HideCellMark.HasValue) {
                _cell.HideCellMark = style.HideCellMark.Value;
            }

            if (style.FitText.HasValue) {
                _cell.FitText = style.FitText.Value;
            }

            if (style.PaddingTopTwips.HasValue ||
                style.PaddingLeftTwips.HasValue ||
                style.PaddingBottomTwips.HasValue ||
                style.PaddingRightTwips.HasValue) {
                _cell.SetPadding(style.PaddingTopTwips, style.PaddingLeftTwips, style.PaddingBottomTwips, style.PaddingRightTwips);
            }

            ApplyCellBorder(_cell.TopBorder, style.TopBorder);
            ApplyCellBorder(_cell.LeftBorder, style.LeftBorder);
            ApplyCellBorder(_cell.BottomBorder, style.BottomBorder);
            ApplyCellBorder(_cell.RightBorder, style.RightBorder);
        }

        private void ApplyCellAttributes(HtmlToken token) {
            if (_cell == null) {
                return;
            }

            string? align = GetAttribute(token, "align");
            if (!string.IsNullOrWhiteSpace(align) && TryParseTextAlign(align!, out RtfTextAlignment textAlignment)) {
                _cellTextAlignment = textAlignment;
            }

            string? verticalAlign = GetAttribute(token, "valign");
            if (!string.IsNullOrWhiteSpace(verticalAlign) && TryParseCellVerticalAlign(verticalAlign!, out RtfTableCellVerticalAlignment cellAlignment)) {
                _cell.VerticalAlignment = cellAlignment;
            }

            string? background = GetAttribute(token, "bgcolor");
            if (!string.IsNullOrWhiteSpace(background) && HtmlStyleDeclarationParser.TryParseColor(background!, out RtfColor? color)) {
                _cell.BackgroundColorIndex = GetOrAddColorIndex(color!);
            }

            string? width = GetAttribute(token, "width");
            if (!string.IsNullOrWhiteSpace(width) && HtmlStyleDeclarationParser.TryParseTableWidth(width!, out int tableWidth, out RtfTableWidthUnit widthUnit)) {
                _cell.PreferredWidth = tableWidth;
                _cell.PreferredWidthUnit = widthUnit;
            }

            if (token.Attributes.ContainsKey("nowrap")) {
                _cell.NoWrap = true;
            }
        }

        private void ApplyCellBorder(RtfTableCellBorder target, HtmlBorderDeclaration? source) {
            if (source == null) {
                return;
            }

            target.Style = source.Style;
            target.Width = source.Width;
            target.ColorIndex = source.Color == null ? null : GetOrAddColorIndex(source.Color);
        }

        private static int ReadSpan(HtmlToken token, string name) {
            string? value = GetAttribute(token, name);
            return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int span) && span > 1
                ? Math.Min(span, 256)
                : 1;
        }

        private static bool TryParseTextAlign(string value, out RtfTextAlignment alignment) {
            switch (value.Trim().ToLowerInvariant()) {
                case "center":
                case "middle":
                    alignment = RtfTextAlignment.Center;
                    return true;
                case "right":
                    alignment = RtfTextAlignment.Right;
                    return true;
                case "justify":
                    alignment = RtfTextAlignment.Justify;
                    return true;
                case "left":
                    alignment = RtfTextAlignment.Left;
                    return true;
                default:
                    alignment = RtfTextAlignment.Left;
                    return false;
            }
        }

        private static bool TryParseTableAlign(string value, out RtfTableAlignment alignment) {
            switch (value.Trim().ToLowerInvariant()) {
                case "center":
                case "middle":
                    alignment = RtfTableAlignment.Center;
                    return true;
                case "right":
                    alignment = RtfTableAlignment.Right;
                    return true;
                case "left":
                    alignment = RtfTableAlignment.Left;
                    return true;
                default:
                    alignment = RtfTableAlignment.Left;
                    return false;
            }
        }

        private static bool TryMapTableAlignment(RtfTextAlignment textAlignment, out RtfTableAlignment alignment) {
            switch (textAlignment) {
                case RtfTextAlignment.Center:
                    alignment = RtfTableAlignment.Center;
                    return true;
                case RtfTextAlignment.Right:
                    alignment = RtfTableAlignment.Right;
                    return true;
                case RtfTextAlignment.Left:
                    alignment = RtfTableAlignment.Left;
                    return true;
                default:
                    alignment = RtfTableAlignment.Left;
                    return false;
            }
        }

        private static RtfTableRowDirection MapTableRowDirection(RtfTextDirection direction) =>
            direction == RtfTextDirection.RightToLeft
                ? RtfTableRowDirection.RightToLeft
                : RtfTableRowDirection.LeftToRight;

        private static bool TryParseCellVerticalAlign(string value, out RtfTableCellVerticalAlignment alignment) {
            switch (value.Trim().ToLowerInvariant()) {
                case "middle":
                case "center":
                    alignment = RtfTableCellVerticalAlignment.Center;
                    return true;
                case "bottom":
                case "baseline":
                    alignment = RtfTableCellVerticalAlignment.Bottom;
                    return true;
                case "top":
                    alignment = RtfTableCellVerticalAlignment.Top;
                    return true;
                default:
                    alignment = RtfTableCellVerticalAlignment.Top;
                    return false;
            }
        }

        private sealed class RowSpanState {
            internal int RemainingRows { get; set; }

            internal int ColumnSpan { get; set; } = 1;

            internal int Offset { get; set; }
        }
    }
}
