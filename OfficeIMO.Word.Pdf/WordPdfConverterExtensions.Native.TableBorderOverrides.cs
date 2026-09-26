using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static bool HasNativeDirectCellBorder(WordTableCellBorder borders) =>
            borders.TopStyle.HasValue || borders.RightStyle.HasValue ||
            borders.BottomStyle.HasValue || borders.LeftStyle.HasValue ||
            borders.TopLeftToBottomRightStyle.HasValue || borders.TopRightToBottomLeftStyle.HasValue;

        private static bool HasNativeConditionalHiddenBorders(WordTable table, NativeTableStyleDefaults defaults) =>
            (table.ConditionalFormattingFirstRow == true && HasHiddenBorder(defaults.FirstRowStyle.CellBorders)) ||
            (table.ConditionalFormattingLastRow == true && HasHiddenBorder(defaults.LastRowStyle.CellBorders)) ||
            (table.ConditionalFormattingFirstColumn == true && HasHiddenBorder(defaults.FirstColumnStyle.CellBorders)) ||
            (table.ConditionalFormattingLastColumn == true && HasHiddenBorder(defaults.LastColumnStyle.CellBorders)) ||
            (table.ConditionalFormattingNoHorizontalBand != true && HasHiddenBorder(defaults.Band1HorizontalStyle.CellBorders)) ||
            (table.ConditionalFormattingNoVerticalBand != true && HasHiddenBorder(defaults.Band1VerticalStyle.CellBorders));

        private static bool HasHiddenBorder(W.TableCellBorders? borders) {
            if (borders == null) {
                return false;
            }

            foreach (var child in borders.ChildElements) {
                if (child is W.BorderType border &&
                    (border.Val?.Value == W.BorderValues.Nil || border.Val?.Value == W.BorderValues.None)) {
                    return true;
                }
            }

            return false;
        }

        private static void ApplyNativeDirectCellBorders(
            PdfCore.PdfTableStyle style,
            TableLayout layout,
            Dictionary<(int Row, int Column), WordTableCellBorder> directBorders) {
            if (directBorders.Count == 0) {
                return;
            }

            var effective = style.CellBorders == null
                ? new Dictionary<(int Row, int Column), PdfCore.PdfCellBorder>()
                : new Dictionary<(int Row, int Column), PdfCore.PdfCellBorder>(style.CellBorders);

            // The PDF table grid is drawn before cell borders. Convert it to cell borders
            // so an explicit Word nil or none side can suppress its own inherited line.
            // Only nil suppresses the opposing side in a collapsed-border conflict.
            MaterializeNativeTableBorderGrid(style, layout, effective);

            foreach (var entry in directBorders) {
                WordTableCellBorder source = entry.Value;
                PdfCore.PdfCellBorder border = effective.TryGetValue(entry.Key, out PdfCore.PdfCellBorder? inherited)
                    ? inherited.Clone()
                    : new PdfCore.PdfCellBorder {
                        Color = null,
                        Width = 0D,
                        Top = false,
                        Right = false,
                        Bottom = false,
                        Left = false
                    };

                if (source.TopStyle.HasValue) {
                    border.Top = HasNativeBorder(source.TopStyle);
                    border.TopBorder = CreateNativeCellBorderSide(source.TopStyle, source.TopColorHex, source.TopSize) ?? NativeHiddenCellBorderSide();
                }
                if (source.RightStyle.HasValue) {
                    border.Right = HasNativeBorder(source.RightStyle);
                    border.RightBorder = CreateNativeCellBorderSide(source.RightStyle, source.RightColorHex, source.RightSize) ?? NativeHiddenCellBorderSide();
                }
                if (source.BottomStyle.HasValue) {
                    border.Bottom = HasNativeBorder(source.BottomStyle);
                    border.BottomBorder = CreateNativeCellBorderSide(source.BottomStyle, source.BottomColorHex, source.BottomSize) ?? NativeHiddenCellBorderSide();
                }
                if (source.LeftStyle.HasValue) {
                    border.Left = HasNativeBorder(source.LeftStyle);
                    border.LeftBorder = CreateNativeCellBorderSide(source.LeftStyle, source.LeftColorHex, source.LeftSize) ?? NativeHiddenCellBorderSide();
                }
                if (source.TopLeftToBottomRightStyle.HasValue) {
                    border.DiagonalDown = HasNativeBorder(source.TopLeftToBottomRightStyle);
                    border.DiagonalDownBorder = CreateNativeCellBorderSide(source.TopLeftToBottomRightStyle, source.TopLeftToBottomRightColorHex, source.TopLeftToBottomRightSize);
                }
                if (source.TopRightToBottomLeftStyle.HasValue) {
                    border.DiagonalUp = HasNativeBorder(source.TopRightToBottomLeftStyle);
                    border.DiagonalUpBorder = CreateNativeCellBorderSide(source.TopRightToBottomLeftStyle, source.TopRightToBottomLeftColorHex, source.TopRightToBottomLeftSize);
                }
                effective[entry.Key] = border;
            }

            style.CellBorders = effective;
        }

        private static PdfCore.PdfCellBorderSide NativeHiddenCellBorderSide() =>
            new PdfCore.PdfCellBorderSide { Color = null, Width = 0D };

        private static bool IsNativeHiddenCellBorderSide(PdfCore.PdfCellBorderSide? side) =>
            side != null && !side.Color.HasValue && side.Width <= 0D;

        private static void ReconcileNativeHiddenSharedBorders(
            WordTable table,
            TableLayout layout,
            NativeTableStyleDefaults tableStyleDefaults,
            int headerRowCount,
            Dictionary<(int Row, int Column), PdfCore.PdfCellBorder> borders,
            Dictionary<(int Row, int Column), WordTableCellBorder> directBorders) {
            int columnCount = GetNativeTableColumnCount(layout);
            var occupied = new Dictionary<(int Row, int Column), (int Row, int Column)>();
            var spans = new Dictionary<(int Row, int Column), (int Rows, int Columns)>();
            for (int rowIndex = 0; rowIndex < layout.Rows.Count; rowIndex++) {
                int logicalColumn = GetNativeTableRowStartColumn(layout, rowIndex);
                foreach (WordTableCell cell in layout.Rows[rowIndex]) {
                    if (IsNativeHorizontalMergeContinuation(cell)) {
                        continue;
                    }

                    int columnSpan = GetNativeCellColumnSpan(cell);
                    if (!IsNativeVerticalMergeContinuation(cell)) {
                        var key = (rowIndex, logicalColumn);
                        int rowSpan = GetNativeCellRowSpan(cell);
                        spans[key] = (rowSpan, columnSpan);
                        for (int row = rowIndex; row < rowIndex + rowSpan && row < layout.Rows.Count; row++) {
                            for (int column = logicalColumn; column < logicalColumn + columnSpan; column++) {
                                occupied[(row, column)] = key;
                            }
                        }
                    }
                    logicalColumn += columnSpan;
                }
            }

            foreach (var entry in spans) {
                if (!borders.TryGetValue(entry.Key, out PdfCore.PdfCellBorder? border)) {
                    continue;
                }
                int row = entry.Key.Item1;
                int column = entry.Key.Item2;
                if (IsNativeHiddenCellBorderSide(border.RightBorder) && IsNativeNilBorderSide(table, layout, tableStyleDefaults, headerRowCount, columnCount, entry.Key, entry.Value.Columns, NativeBorderEdge.Right, directBorders)) {
                    for (int offset = 0; offset < entry.Value.Rows; offset++) {
                        HideNativeNeighborBorder(occupied, spans, borders, entry.Key, (row + offset, column + entry.Value.Columns), NativeBorderEdge.Left);
                    }
                }
                if (IsNativeHiddenCellBorderSide(border.LeftBorder) && IsNativeNilBorderSide(table, layout, tableStyleDefaults, headerRowCount, columnCount, entry.Key, entry.Value.Columns, NativeBorderEdge.Left, directBorders)) {
                    for (int offset = 0; offset < entry.Value.Rows; offset++) {
                        HideNativeNeighborBorder(occupied, spans, borders, entry.Key, (row + offset, column - 1), NativeBorderEdge.Right);
                    }
                }
                if (IsNativeHiddenCellBorderSide(border.BottomBorder) && IsNativeNilBorderSide(table, layout, tableStyleDefaults, headerRowCount, columnCount, entry.Key, entry.Value.Columns, NativeBorderEdge.Bottom, directBorders)) {
                    for (int offset = 0; offset < entry.Value.Columns; offset++) {
                        HideNativeNeighborBorder(occupied, spans, borders, entry.Key, (row + entry.Value.Rows, column + offset), NativeBorderEdge.Top);
                    }
                }
                if (IsNativeHiddenCellBorderSide(border.TopBorder) && IsNativeNilBorderSide(table, layout, tableStyleDefaults, headerRowCount, columnCount, entry.Key, entry.Value.Columns, NativeBorderEdge.Top, directBorders)) {
                    for (int offset = 0; offset < entry.Value.Columns; offset++) {
                        HideNativeNeighborBorder(occupied, spans, borders, entry.Key, (row - 1, column + offset), NativeBorderEdge.Bottom);
                    }
                }
            }
        }

        private enum NativeBorderEdge { Top, Right, Bottom, Left }

        private static bool IsNativeNilBorderSide(
            WordTable table,
            TableLayout layout,
            NativeTableStyleDefaults defaults,
            int headerRowCount,
            int columnCount,
            (int Row, int Column) key,
            int columnSpan,
            NativeBorderEdge edge,
            Dictionary<(int Row, int Column), WordTableCellBorder> directBorders) {
            if (directBorders.TryGetValue(key, out WordTableCellBorder? direct)) {
                WordBorderStyle? style = edge switch {
                    NativeBorderEdge.Top => direct.TopStyle,
                    NativeBorderEdge.Right => direct.RightStyle,
                    NativeBorderEdge.Bottom => direct.BottomStyle,
                    _ => direct.LeftStyle
                };
                if (style.HasValue) {
                    return style.Value == WordBorderStyle.Nil;
                }
            }

            W.BorderValues? conditional = null;
            if (table.ConditionalFormattingFirstRow == true && key.Row == 0) {
                UpdateNativeConditionalBorderValue(ref conditional, defaults.FirstRowStyle.CellBorders, edge);
            }
            if (table.ConditionalFormattingLastRow == true && layout.Rows.Count > headerRowCount && key.Row == layout.Rows.Count - 1) {
                UpdateNativeConditionalBorderValue(ref conditional, defaults.LastRowStyle.CellBorders, edge);
            }
            int footerStartRow = table.ConditionalFormattingLastRow == true && layout.Rows.Count > headerRowCount
                ? layout.Rows.Count - 1 : layout.Rows.Count;
            if (key.Row >= headerRowCount && key.Row < footerStartRow) {
                if (table.ConditionalFormattingNoHorizontalBand != true && (key.Row - headerRowCount) % 2 == 1) {
                    UpdateNativeConditionalBorderValue(ref conditional, defaults.Band1HorizontalStyle.CellBorders, edge);
                }
                if (table.ConditionalFormattingNoVerticalBand != true && key.Column % 2 == 1) {
                    UpdateNativeConditionalBorderValue(ref conditional, defaults.Band1VerticalStyle.CellBorders, edge);
                }
            }
            if (table.ConditionalFormattingFirstColumn == true && key.Column == 0) {
                UpdateNativeConditionalBorderValue(ref conditional, defaults.FirstColumnStyle.CellBorders, edge);
            }
            if (table.ConditionalFormattingLastColumn == true && key.Column + columnSpan >= columnCount) {
                UpdateNativeConditionalBorderValue(ref conditional, defaults.LastColumnStyle.CellBorders, edge);
            }
            // Word's nil wins a collapsed shared edge; none hides only its own side.
            return conditional != W.BorderValues.None;
        }

        private static void UpdateNativeConditionalBorderValue(ref W.BorderValues? value, W.TableCellBorders? borders, NativeBorderEdge edge) {
            W.BorderType? border = edge switch {
                NativeBorderEdge.Top => borders?.GetFirstChild<W.TopBorder>(),
                NativeBorderEdge.Right => (W.BorderType?)borders?.GetFirstChild<W.RightBorder>() ?? borders?.GetFirstChild<W.EndBorder>(),
                NativeBorderEdge.Bottom => borders?.GetFirstChild<W.BottomBorder>(),
                _ => (W.BorderType?)borders?.GetFirstChild<W.LeftBorder>() ?? borders?.GetFirstChild<W.StartBorder>()
            };
            if (HasNativeConditionalBorderOverride(border)) {
                value = border!.Val!.Value;
            }
        }

        private static void HideNativeNeighborBorder(
            Dictionary<(int Row, int Column), (int Row, int Column)> occupied,
            Dictionary<(int Row, int Column), (int Rows, int Columns)> spans,
            Dictionary<(int Row, int Column), PdfCore.PdfCellBorder> borders,
            (int Row, int Column) source,
            (int Row, int Column) neighborPosition,
            NativeBorderEdge edge) {
            if (!occupied.TryGetValue(neighborPosition, out var neighborKey) || neighborKey == source ||
                !spans.TryGetValue(neighborKey, out var neighborSpan) ||
                !borders.TryGetValue(neighborKey, out PdfCore.PdfCellBorder? neighbor)) {
                return;
            }

            int rowSegment = neighborPosition.Row - neighborKey.Row;
            int columnSegment = neighborPosition.Column - neighborKey.Column;

            switch (edge) {
                case NativeBorderEdge.Top:
                    if (neighborSpan.Columns > 1) {
                        neighbor.HiddenTopColumnSegments ??= new HashSet<int>();
                        neighbor.HiddenTopColumnSegments.Add(columnSegment);
                        break;
                    }
                    neighbor.Top = false;
                    neighbor.TopBorder = NativeHiddenCellBorderSide();
                    break;
                case NativeBorderEdge.Right:
                    if (neighborSpan.Rows > 1) {
                        neighbor.HiddenRightRowSegments ??= new HashSet<int>();
                        neighbor.HiddenRightRowSegments.Add(rowSegment);
                        break;
                    }
                    neighbor.Right = false;
                    neighbor.RightBorder = NativeHiddenCellBorderSide();
                    break;
                case NativeBorderEdge.Bottom:
                    if (neighborSpan.Columns > 1) {
                        neighbor.HiddenBottomColumnSegments ??= new HashSet<int>();
                        neighbor.HiddenBottomColumnSegments.Add(columnSegment);
                        break;
                    }
                    neighbor.Bottom = false;
                    neighbor.BottomBorder = NativeHiddenCellBorderSide();
                    break;
                default:
                    if (neighborSpan.Rows > 1) {
                        neighbor.HiddenLeftRowSegments ??= new HashSet<int>();
                        neighbor.HiddenLeftRowSegments.Add(rowSegment);
                        break;
                    }
                    neighbor.Left = false;
                    neighbor.LeftBorder = NativeHiddenCellBorderSide();
                    break;
            }
        }

        private static void MaterializeNativeTableBorderGrid(PdfCore.PdfTableStyle style, TableLayout layout) {
            var effective = style.CellBorders == null
                ? new Dictionary<(int Row, int Column), PdfCore.PdfCellBorder>()
                : new Dictionary<(int Row, int Column), PdfCore.PdfCellBorder>(style.CellBorders);
            MaterializeNativeTableBorderGrid(style, layout, effective);
            style.CellBorders = effective;
        }

        private static void MaterializeNativeTableBorderGrid(
            PdfCore.PdfTableStyle style,
            TableLayout layout,
            Dictionary<(int Row, int Column), PdfCore.PdfCellBorder> effective) {
            if (!style.BorderColor.HasValue || style.BorderWidth <= 0D) {
                return;
            }

            for (int rowIndex = 0; rowIndex < layout.Rows.Count; rowIndex++) {
                int logicalColumn = GetNativeTableRowStartColumn(layout, rowIndex);
                foreach (WordTableCell cell in layout.Rows[rowIndex]) {
                    if (IsNativeHorizontalMergeContinuation(cell)) {
                        continue;
                    }

                    int span = GetNativeCellColumnSpan(cell);
                    if (!IsNativeVerticalMergeContinuation(cell)) {
                        var key = (rowIndex, logicalColumn);
                        var inherited = new PdfCore.PdfCellBorder {
                            Color = style.BorderColor,
                            Width = style.BorderWidth
                        };
                        if (effective.TryGetValue(key, out PdfCore.PdfCellBorder? existing)) {
                            // A configured cell border is a full cell override. Conditional
                            // Word borders use side overrides and only replace those sides.
                            inherited = (existing.Color.HasValue && existing.Width > 0D) ||
                                (!existing.Top && !existing.Right && !existing.Bottom && !existing.Left &&
                                 !existing.DiagonalUp && !existing.DiagonalDown)
                                ? existing.Clone()
                                : MergeNativeCellBorder(inherited, existing);
                        }
                        effective[key] = inherited;
                    }
                    logicalColumn += span;
                }
            }
            style.BorderColor = null;
            style.BorderWidth = 0D;
        }
    }
}
