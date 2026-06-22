using System;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Applies reusable row layout and row-wide cell style options to one worksheet row.
        /// </summary>
        /// <param name="rowIndex">The 1-based row index to update.</param>
        /// <param name="options">The layout options to apply.</param>
        public void SetRowLayout(int rowIndex, ExcelRowLayoutOptions options) {
            if (rowIndex <= 0 || rowIndex > A1.MaxRows) {
                throw new ArgumentOutOfRangeException(nameof(rowIndex), "Row index must be between 1 and the Excel row limit.");
            }

            if (options == null) {
                throw new ArgumentNullException(nameof(options));
            }

            if (options.ClearHeight) {
                SetRowHeight(rowIndex, 0);
            }

            if (options.Height.HasValue) {
                SetRowHeight(rowIndex, options.Height.Value);
            }

            if (options.Hidden.HasValue) {
                SetRowHidden(rowIndex, options.Hidden.Value);
            }

            if (HasRowCellStyleOptions(options)) {
                var (firstColumn, lastColumn) = ResolveRowLayoutColumnBounds(options);
                ApplyRowCellStyle(rowIndex, firstColumn, lastColumn, options);
            }

            if (options.AutoFit) {
                AutoFitRow(rowIndex);
            }
        }

        /// <summary>
        /// Applies the same reusable row layout and row-wide cell style options to multiple worksheet rows.
        /// </summary>
        /// <param name="rowIndexes">The 1-based row indexes to update.</param>
        /// <param name="options">The layout options to apply.</param>
        public void SetRowsLayout(IEnumerable<int> rowIndexes, ExcelRowLayoutOptions options) {
            if (rowIndexes == null) {
                throw new ArgumentNullException(nameof(rowIndexes));
            }

            foreach (int rowIndex in rowIndexes) {
                SetRowLayout(rowIndex, options);
            }
        }

        private static bool HasRowCellStyleOptions(ExcelRowLayoutOptions options) {
            return options.Bold.HasValue
                || options.Italic.HasValue
                || options.Underline.HasValue
                || options.WrapText.HasValue
                || !string.IsNullOrWhiteSpace(options.FontName)
                || !string.IsNullOrWhiteSpace(options.BackgroundColor);
        }

        private (int FirstColumn, int LastColumn) ResolveRowLayoutColumnBounds(ExcelRowLayoutOptions options) {
            int firstColumn;
            int lastColumn;

            if (options.FirstColumn.HasValue || options.LastColumn.HasValue) {
                firstColumn = options.FirstColumn ?? options.LastColumn!.Value;
                lastColumn = options.LastColumn ?? firstColumn;
            } else if (A1.TryParseRange(GetUsedRangeA1(), out _, out int usedFirstColumn, out _, out int usedLastColumn)) {
                firstColumn = usedFirstColumn;
                lastColumn = usedLastColumn;
            } else {
                firstColumn = 1;
                lastColumn = 1;
            }

            if (firstColumn <= 0) {
                throw new ArgumentOutOfRangeException(nameof(options), "FirstColumn must be 1 or greater.");
            }

            if (lastColumn < firstColumn) {
                throw new ArgumentException("LastColumn must be greater than or equal to FirstColumn.", nameof(options));
            }

            if (lastColumn > A1.MaxColumns) {
                throw new ArgumentOutOfRangeException(nameof(options), "LastColumn exceeds the maximum Excel column.");
            }

            return (firstColumn, lastColumn);
        }

        private void ApplyRowCellStyle(int rowIndex, int firstColumn, int lastColumn, ExcelRowLayoutOptions options) {
            for (int columnIndex = firstColumn; columnIndex <= lastColumn; columnIndex++) {
                if (options.Bold.HasValue) {
                    CellBold(rowIndex, columnIndex, options.Bold.Value);
                }

                if (options.Italic.HasValue) {
                    CellItalic(rowIndex, columnIndex, options.Italic.Value);
                }

                if (options.Underline.HasValue) {
                    CellUnderline(rowIndex, columnIndex, options.Underline.Value);
                }

                if (options.WrapText.HasValue) {
                    CellWrapText(rowIndex, columnIndex, options.WrapText.Value);
                }

                if (!string.IsNullOrWhiteSpace(options.FontName)) {
                    CellFontName(rowIndex, columnIndex, options.FontName!);
                }

                if (!string.IsNullOrWhiteSpace(options.BackgroundColor)) {
                    CellBackground(rowIndex, columnIndex, options.BackgroundColor!);
                }
            }
        }
    }
}
