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
            // so an explicit Word nil/none side can actually suppress the inherited line.
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
                    border.TopBorder = CreateNativeCellBorderSide(source.TopStyle, source.TopColorHex, source.TopSize);
                }
                if (source.RightStyle.HasValue) {
                    border.Right = HasNativeBorder(source.RightStyle);
                    border.RightBorder = CreateNativeCellBorderSide(source.RightStyle, source.RightColorHex, source.RightSize);
                }
                if (source.BottomStyle.HasValue) {
                    border.Bottom = HasNativeBorder(source.BottomStyle);
                    border.BottomBorder = CreateNativeCellBorderSide(source.BottomStyle, source.BottomColorHex, source.BottomSize);
                }
                if (source.LeftStyle.HasValue) {
                    border.Left = HasNativeBorder(source.LeftStyle);
                    border.LeftBorder = CreateNativeCellBorderSide(source.LeftStyle, source.LeftColorHex, source.LeftSize);
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
