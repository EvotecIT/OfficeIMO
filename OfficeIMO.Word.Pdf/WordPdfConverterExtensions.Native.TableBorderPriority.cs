using System.Collections.Generic;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static bool HasNativeConditionalCellBorders(WordTable table, NativeTableStyleDefaults defaults) =>
            (table.ConditionalFormattingFirstRow == true && defaults.FirstRowStyle.CellBorders?.HasChildren == true) ||
            (table.ConditionalFormattingLastRow == true && defaults.LastRowStyle.CellBorders?.HasChildren == true) ||
            (table.ConditionalFormattingFirstColumn == true && defaults.FirstColumnStyle.CellBorders?.HasChildren == true) ||
            (table.ConditionalFormattingLastColumn == true && defaults.LastColumnStyle.CellBorders?.HasChildren == true) ||
            (table.ConditionalFormattingNoHorizontalBand != true && (defaults.Band1HorizontalStyle.CellBorders?.HasChildren == true || defaults.Band2HorizontalStyle.CellBorders?.HasChildren == true)) ||
            (table.ConditionalFormattingNoVerticalBand != true && (defaults.Band1VerticalStyle.CellBorders?.HasChildren == true || defaults.Band2VerticalStyle.CellBorders?.HasChildren == true));

        /// <summary>Word cell borders take precedence over adjacent borders inherited from the table.</summary>
        private static void ReconcileNativeCellTableBorders(WordTable table, TableLayout layout,
            NativeTableStyleDefaults defaults, int headerRows, int columns,
            Dictionary<(int Row, int Column), (int Rows, int Columns)> spans,
            Dictionary<(int Row, int Column), (int Row, int Column)> occupied,
            Dictionary<(int Row, int Column), PdfCore.PdfCellBorder> borders,
            Dictionary<(int Row, int Column), WordTableCellBorder> direct) {
            foreach (var entry in spans) {
                if (!borders.TryGetValue(entry.Key, out PdfCore.PdfCellBorder? border)) continue;
                border.TopPriority = GetPriority(entry.Key, entry.Value.Columns, NativeBorderEdge.Top);
                Reconcile(NativeBorderEdge.Right, border.Right, border.RightBorder);
                border.BottomPriority = GetPriority(entry.Key, entry.Value.Columns, NativeBorderEdge.Bottom);
                Reconcile(NativeBorderEdge.Left, border.Left, border.LeftBorder);

                void Reconcile(NativeBorderEdge edge, bool enabled, PdfCore.PdfCellBorderSide? side) {
                    if (!enabled || side?.Color == null || side.Width <= 0D ||
                        !HasCellOverride(entry.Key, entry.Value.Columns, edge)) return;
                    bool horizontal = edge == NativeBorderEdge.Top || edge == NativeBorderEdge.Bottom;
                    int segments = horizontal ? entry.Value.Columns : entry.Value.Rows;
                    NativeBorderEdge opposing = edge switch {
                        NativeBorderEdge.Top => NativeBorderEdge.Bottom,
                        NativeBorderEdge.Right => NativeBorderEdge.Left,
                        NativeBorderEdge.Bottom => NativeBorderEdge.Top,
                        _ => NativeBorderEdge.Right
                    };
                    for (int segment = 0; segment < segments; segment++) {
                        (int Row, int Column) position = edge switch {
                            NativeBorderEdge.Top => (entry.Key.Row - 1, entry.Key.Column + segment),
                            NativeBorderEdge.Right => (entry.Key.Row + segment, entry.Key.Column + entry.Value.Columns),
                            NativeBorderEdge.Bottom => (entry.Key.Row + entry.Value.Rows, entry.Key.Column + segment),
                            _ => (entry.Key.Row + segment, entry.Key.Column - 1)
                        };
                        if (!occupied.TryGetValue(position, out var neighbor) || neighbor == entry.Key ||
                            !spans.TryGetValue(neighbor, out var neighborSpan) || HasCellOverride(neighbor, neighborSpan.Columns, opposing)) continue;
                        HideNativeNeighborBorder(occupied, spans, borders, entry.Key, position, opposing);
                    }
                }
            }

            PdfCore.PdfCellBorderPriority GetPriority((int Row, int Column) key, int span, NativeBorderEdge edge) {
                WordBorderStyle? value = null;
                if (direct.TryGetValue(key, out WordTableCellBorder? source))
                    value = edge == NativeBorderEdge.Top ? source.TopStyle : source.BottomStyle;
                if (value.HasValue) return value == WordBorderStyle.Nil ? PdfCore.PdfCellBorderPriority.Nil : PdfCore.PdfCellBorderPriority.Cell;
                var conditional = GetNativeConditionalCellBorderValue(table, layout, defaults, headerRows, columns, key, span, edge);
                return conditional == DocumentFormat.OpenXml.Wordprocessing.BorderValues.Nil ? PdfCore.PdfCellBorderPriority.Nil
                    : conditional.HasValue ? PdfCore.PdfCellBorderPriority.Cell : PdfCore.PdfCellBorderPriority.Table;
            }

            bool HasCellOverride((int Row, int Column) key, int span, NativeBorderEdge edge) {
                if (direct.TryGetValue(key, out WordTableCellBorder? source)) {
                    WordBorderStyle? value = edge switch {
                        NativeBorderEdge.Top => source.TopStyle,
                        NativeBorderEdge.Right => source.RightStyle,
                        NativeBorderEdge.Bottom => source.BottomStyle,
                        _ => source.LeftStyle
                    };
                    if (value.HasValue) return true;
                }
                return GetNativeConditionalCellBorderValue(table, layout, defaults, headerRows, columns, key, span, edge).HasValue;
            }
        }
    }
}
