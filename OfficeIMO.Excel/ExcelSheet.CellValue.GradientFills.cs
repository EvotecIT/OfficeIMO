using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Applies a two-color linear gradient background to a single cell. Native theme roles remain linked to the workbook theme.
        /// </summary>
        /// <param name="row">The 1-based row index of the cell to fill.</param>
        /// <param name="column">The 1-based column index of the cell to fill.</param>
        /// <param name="fromColor">The gradient start color as a literal or theme role.</param>
        /// <param name="toColor">The gradient end color as a literal or theme role.</param>
        /// <param name="degree">The linear gradient angle in degrees.</param>
        internal void CellGradientBackground(int row, int column, OfficeColor fromColor, OfficeColor toColor, double degree = 0) {
            WriteLockConditional(() => {
                var cell = GetCell(row, column);
                ApplyGradientBackground(cell, fromColor, toColor, degree);
            });
        }

        /// <summary>
        /// Applies a two-color linear gradient fill to every cell in the range.
        /// </summary>
        /// <param name="a1Range">The A1 range to fill.</param>
        /// <param name="fromColor">The gradient start color as a literal or theme role.</param>
        /// <param name="toColor">The gradient end color as a literal or theme role.</param>
        /// <param name="degree">The linear gradient angle in degrees.</param>
        internal void FillRangeGradient(string a1Range, OfficeColor fromColor, OfficeColor toColor, double degree = 0) {
            var (r1, c1, r2, c2) = ExcelA1.ParseRange(a1Range);

            if (!_excelDocument.IsMaterializingDeferredDataSetImport) {
                MaterializeDeferredDataSetImportIfNeeded();
            }

            WriteLock(() => FillRangeGradientCore(r1, c1, r2, c2, fromColor, toColor, degree));
        }

        private void ApplyGradientBackground(Cell cell, OfficeColor fromColor, OfficeColor toColor, double degree) {
            var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
            var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<DocumentFormat.OpenXml.Packaging.WorkbookStylesPart>();
            var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            EnsureDefaultStylePrimitives(stylesheet);

            uint fillId = GetOrCreateFill(stylesheet, CreateGradientFill(fromColor, toColor, degree));
            ApplyCellFormatOverride(stylesheet, cell, format => {
                format.FillId = fillId;
                format.ApplyFill = true;
            });

            SaveStylesheet(stylesPart);
        }

        private void FillRangeGradientCore(int firstRow, int firstColumn, int lastRow, int lastColumn, OfficeColor fromColor, OfficeColor toColor, double degree) {
            var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
            var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<DocumentFormat.OpenXml.Packaging.WorkbookStylesPart>();
            var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            EnsureDefaultStylePrimitives(stylesheet);

            uint fillId = GetOrCreateFill(stylesheet, CreateGradientFill(fromColor, toColor, degree));
            var styleIndexes = new Dictionary<uint, uint>();

            for (int row = firstRow; row <= lastRow; row++) {
                for (int column = firstColumn; column <= lastColumn; column++) {
                    Cell cell = GetCell(row, column);
                    uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
                    cell.StyleIndex = GetOrAddCellFormatOverride(styleIndexes, stylesheet, baseStyleIndex, format => {
                        format.FillId = fillId;
                        format.ApplyFill = true;
                    });
                }
            }

            SaveStylesheet(stylesPart);
        }

        private Fill CreateGradientFill(OfficeColor fromColor, OfficeColor toColor, double degree) {
            return new Fill(new GradientFill(
                new GradientStop(ExcelColor.ToSpreadsheetColor(fromColor, _excelDocument.ColorTheme)) { Position = 0D },
                new GradientStop(ExcelColor.ToSpreadsheetColor(toColor, _excelDocument.ColorTheme)) { Position = 1D }) {
                Type = GradientValues.Linear,
                Degree = degree
            });
        }
    }
}
