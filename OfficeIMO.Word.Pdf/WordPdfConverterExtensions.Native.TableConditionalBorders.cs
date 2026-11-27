using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {

    public static partial class WordPdfConverterExtensions {
        private static void ApplyNativeTableConditionalBorders(WordTable table, TableLayout layout, NativeTableStyleDefaults tableStyleDefaults, PdfCore.PdfTableStyle style) {
            if (layout.Rows.Count == 0) {
                return;
            }

            Dictionary<(int Row, int Column), PdfCore.PdfCellBorder>? cellBorders = style.CellBorders == null
                ? null
                : new Dictionary<(int Row, int Column), PdfCore.PdfCellBorder>(style.CellBorders);
            cellBorders ??= new Dictionary<(int Row, int Column), PdfCore.PdfCellBorder>();
            bool changed = false;

            int footerStartRowIndex = table.ConditionalFormattingLastRow == true && layout.Rows.Count > style.HeaderRowCount
                ? layout.Rows.Count - 1 : layout.Rows.Count;
            changed |= ApplyNativeTableConditionalBandBorders(table, layout, tableStyleDefaults, style.HeaderRowCount, footerStartRowIndex, cellBorders);
            if (table.ConditionalFormattingFirstRow == true) {
                changed |= ApplyNativeTableConditionalRowBorder(layout, 0, tableStyleDefaults.FirstRowStyle, cellBorders);
            }

            if (table.ConditionalFormattingLastRow == true && layout.Rows.Count > style.HeaderRowCount) {
                changed |= ApplyNativeTableConditionalRowBorder(layout, layout.Rows.Count - 1, tableStyleDefaults.LastRowStyle, cellBorders);
            }

            int columnCount = GetNativeTableColumnCount(layout);
            if (columnCount > 0) {
                changed |= ApplyNativeTableConditionalColumnBorders(table, layout, tableStyleDefaults, columnCount, cellBorders);
            }

            if (changed) {
                style.CellBorders = cellBorders;
            }
        }

        private static bool ApplyNativeTableConditionalRowBorder(TableLayout layout, int rowIndex, NativeTableConditionalStyleDefaults conditionalStyle, Dictionary<(int Row, int Column), PdfCore.PdfCellBorder> cellBorders) {
            PdfCore.PdfCellBorder? border = CreateNativeTableConditionalCellBorder(conditionalStyle.CellBorders);
            if (border == null) {
                return false;
            }

            IReadOnlyList<WordTableCell> row = layout.Rows[rowIndex];
            bool changed = false;
            int logicalColumnIndex = GetNativeTableRowStartColumn(layout, rowIndex);
            for (int cellIndex = 0; cellIndex < row.Count; cellIndex++) {
                WordTableCell cell = row[cellIndex];
                if (IsNativeHorizontalMergeContinuation(cell)) {
                    continue;
                }

                int columnSpan = GetNativeCellColumnSpan(cell);
                if (IsNativeVerticalMergeContinuation(cell)) {
                    logicalColumnIndex += columnSpan;
                    continue;
                }

                (int Row, int Column) key = (rowIndex, logicalColumnIndex);
                cellBorders[key] = MergeNativeCellBorder(cellBorders.TryGetValue(key, out PdfCore.PdfCellBorder? existing) ? existing : null, border);
                changed = true;
                logicalColumnIndex += columnSpan;
            }

            return changed;
        }

        private static bool ApplyNativeTableConditionalBandBorders(WordTable table, TableLayout layout, NativeTableStyleDefaults tableStyleDefaults, int headerRowCount, int footerStartRowIndex, Dictionary<(int Row, int Column), PdfCore.PdfCellBorder> cellBorders) {
            bool changed = false;
            for (int rowIndex = headerRowCount; rowIndex < footerStartRowIndex; rowIndex++) {
                IReadOnlyList<WordTableCell> row = layout.Rows[rowIndex];
                int logicalColumnIndex = GetNativeTableRowStartColumn(layout, rowIndex);
                for (int cellIndex = 0; cellIndex < row.Count; cellIndex++) {
                    WordTableCell cell = row[cellIndex];
                    if (IsNativeHorizontalMergeContinuation(cell)) {
                        continue;
                    }

                    int columnSpan = GetNativeCellColumnSpan(cell);
                    if (IsNativeVerticalMergeContinuation(cell)) {
                        logicalColumnIndex += columnSpan;
                        continue;
                    }

                    (int Row, int Column) key = (rowIndex, logicalColumnIndex);
                    changed |= ApplyNativeTableConditionalCellBorder(cellBorders, key, CreateNativeTableConditionalCellBorder(GetNativeHorizontalBandStyle(table, tableStyleDefaults, rowIndex, headerRowCount, footerStartRowIndex).CellBorders));
                    changed |= ApplyNativeTableConditionalCellBorder(cellBorders, key, CreateNativeTableConditionalCellBorder(GetNativeVerticalBandStyle(table, tableStyleDefaults, rowIndex, logicalColumnIndex, headerRowCount, footerStartRowIndex).CellBorders));

                    logicalColumnIndex += columnSpan;
                }
            }

            return changed;
        }

        private static bool ApplyNativeTableConditionalColumnBorders(WordTable table, TableLayout layout, NativeTableStyleDefaults tableStyleDefaults, int columnCount, Dictionary<(int Row, int Column), PdfCore.PdfCellBorder> cellBorders) {
            PdfCore.PdfCellBorder? firstColumnBorder = table.ConditionalFormattingFirstColumn == true
                ? CreateNativeTableConditionalCellBorder(tableStyleDefaults.FirstColumnStyle.CellBorders)
                : null;
            PdfCore.PdfCellBorder? lastColumnBorder = table.ConditionalFormattingLastColumn == true
                ? CreateNativeTableConditionalCellBorder(tableStyleDefaults.LastColumnStyle.CellBorders)
                : null;
            if (firstColumnBorder == null && lastColumnBorder == null) {
                return false;
            }

            bool changed = false;
            for (int rowIndex = 0; rowIndex < layout.Rows.Count; rowIndex++) {
                IReadOnlyList<WordTableCell> row = layout.Rows[rowIndex];
                int logicalColumnIndex = GetNativeTableRowStartColumn(layout, rowIndex);
                for (int cellIndex = 0; cellIndex < row.Count; cellIndex++) {
                    WordTableCell cell = row[cellIndex];
                    if (IsNativeHorizontalMergeContinuation(cell)) {
                        continue;
                    }

                    int columnSpan = GetNativeCellColumnSpan(cell);
                    if (IsNativeVerticalMergeContinuation(cell)) {
                        logicalColumnIndex += columnSpan;
                        continue;
                    }

                    (int Row, int Column) key = (rowIndex, logicalColumnIndex);
                    if (logicalColumnIndex == 0) {
                        changed |= ApplyNativeTableConditionalCellBorder(cellBorders, key, firstColumnBorder);
                    }

                    if (logicalColumnIndex + columnSpan >= columnCount) {
                        changed |= ApplyNativeTableConditionalCellBorder(cellBorders, key, lastColumnBorder);
                    }

                    logicalColumnIndex += columnSpan;
                }
            }

            return changed;
        }

        private static bool ApplyNativeTableConditionalCellBorder(Dictionary<(int Row, int Column), PdfCore.PdfCellBorder> cellBorders, (int Row, int Column) key, PdfCore.PdfCellBorder? border) {
            if (border == null) {
                return false;
            }

            cellBorders[key] = MergeNativeCellBorder(cellBorders.TryGetValue(key, out PdfCore.PdfCellBorder? existing) ? existing : null, border);
            return true;
        }

        private static PdfCore.PdfCellBorder? CreateNativeTableConditionalCellBorder(OfficeXmlNode borders) {
            if (borders.IsNone) {
                return null;
            }

            OfficeXmlNode top = borders.WChild("top");
            OfficeXmlNode right = FirstNativeBorderSide(borders, "right", "end");
            OfficeXmlNode bottom = borders.WChild("bottom");
            OfficeXmlNode left = FirstNativeBorderSide(borders, "left", "start");
            bool hasTop = HasNativeConditionalBorderOverride(top);
            bool hasRight = HasNativeConditionalBorderOverride(right);
            bool hasBottom = HasNativeConditionalBorderOverride(bottom);
            bool hasLeft = HasNativeConditionalBorderOverride(left);
            if (!hasTop && !hasRight && !hasBottom && !hasLeft) {
                return null;
            }

            return new PdfCore.PdfCellBorder {
                Color = null,
                Width = global::OfficeIMO.OfficeLength.Points(0D),
                TopBorder = CreateNativeConditionalBorderSide(top),
                RightBorder = CreateNativeConditionalBorderSide(right),
                BottomBorder = CreateNativeConditionalBorderSide(bottom),
                LeftBorder = CreateNativeConditionalBorderSide(left),
                Top = hasTop,
                Right = hasRight,
                Bottom = hasBottom,
                Left = hasLeft
            };
        }

        /// <summary>The first border side present under one of its two names (physical, then logical).</summary>
        private static OfficeXmlNode FirstNativeBorderSide(OfficeXmlNode borders, string localName, string alternateLocalName) {
            OfficeXmlNode side = borders.WChild(localName);
            return side.Exists ? side : borders.WChild(alternateLocalName);
        }

        private static bool HasNativeConditionalBorderOverride(OfficeXmlNode border) =>
            HasNativeBorder(border.WVal()) ||
            border.WVal() == "nil" ||
            border.WVal() == "none";

        private static PdfCore.PdfCellBorderSide? CreateNativeConditionalBorderSide(OfficeXmlNode border) {
            if (border.WVal() == "nil" || border.WVal() == "none") {
                return new PdfCore.PdfCellBorderSide { Color = null, Width = global::OfficeIMO.OfficeLength.Points(0D) };
            }

            return CreateNativeCellBorderSide(border);
        }

        private static PdfCore.PdfCellBorder MergeNativeCellBorder(PdfCore.PdfCellBorder? existing, PdfCore.PdfCellBorder overlay) {
            PdfCore.PdfCellBorder result = existing?.Clone() ?? new PdfCore.PdfCellBorder {
                Color = null,
                Width = global::OfficeIMO.OfficeLength.Points(0D),
                Top = false,
                Right = false,
                Bottom = false,
                Left = false,
                DiagonalUp = false,
                DiagonalDown = false
            };

            if (overlay.Top) {
                result.Top = true;
                result.TopBorder = overlay.TopBorder;
            }

            if (overlay.Right) {
                result.Right = true;
                result.RightBorder = overlay.RightBorder;
            }

            if (overlay.Bottom) {
                result.Bottom = true;
                result.BottomBorder = overlay.BottomBorder;
            }

            if (overlay.Left) {
                result.Left = true;
                result.LeftBorder = overlay.LeftBorder;
            }

            return result;
        }
    }
}
