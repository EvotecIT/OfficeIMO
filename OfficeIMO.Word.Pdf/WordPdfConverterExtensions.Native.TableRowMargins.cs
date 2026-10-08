using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        /// <summary>Resolves authored row heights into physical Word cell boxes, including effective vertical cell margins.</summary>
        private static void ApplyNativeUnspacedTableRowMargins(WordTable table, TableLayout layout, PdfCore.PdfTableStyle style) {
            // Spaced cells use their independent border-frame geometry instead.
            if (style.CellSpacing > 0D) return;
            var minimums = style.RowMinHeights == null ? new List<double?>() : new List<double?>(style.RowMinHeights);
            var fixedHeights = style.FixedRowHeights == null ? new List<double?>() : new List<double?>(style.FixedRowHeights);
            while (minimums.Count < table.Rows.Count) minimums.Add(null);
            while (fixedHeights.Count < table.Rows.Count) fixedHeights.Add(null);
            bool hasMinimum = false, hasFixed = false;
            for (int row = 0; row < table.Rows.Count; row++) {
                W.TableRowHeight? height = table.Rows[row]._tableRow.TableRowProperties?.GetFirstChild<W.TableRowHeight>();
                if (height?.Val?.Value is not > 0 || height.HeightType?.Value == W.HeightRuleValues.Auto) continue;
                double points = height.Val.Value / 20D;
                double bottom = GetNativeTableRowMargin(layout, style, row, top: false);
                if (height.HeightType?.Value == W.HeightRuleValues.Exact) {
                    fixedHeights[row] = points + bottom;
                    hasFixed = true;
                } else {
                    minimums[row] = points + bottom + GetNativeTableRowMargin(layout, style, row, top: true);
                    hasMinimum = true;
                }
            }
            if (hasFixed) style.FixedRowHeights = fixedHeights;
            if (hasMinimum) style.RowMinHeights = minimums;
        }
    }
}
