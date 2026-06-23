using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Applies a two-color linear gradient background to a single cell. Accepts #RRGGBB or #AARRGGBB.
        /// </summary>
        /// <param name="row">The 1-based row index of the cell to fill.</param>
        /// <param name="column">The 1-based column index of the cell to fill.</param>
        /// <param name="fromHexColor">The gradient start color expressed as an ARGB or RGB hex string.</param>
        /// <param name="toHexColor">The gradient end color expressed as an ARGB or RGB hex string.</param>
        /// <param name="degree">The linear gradient angle in degrees.</param>
        public void CellGradientBackground(int row, int column, string fromHexColor, string toHexColor, double degree = 0) {
            if (string.IsNullOrWhiteSpace(fromHexColor) || string.IsNullOrWhiteSpace(toHexColor)) {
                return;
            }

            WriteLockConditional(() => {
                var cell = GetCell(row, column);
                ApplyGradientBackground(cell, fromHexColor, toHexColor, degree);
            });
        }

        /// <summary>
        /// Applies a two-color linear gradient fill to every cell in the range.
        /// </summary>
        /// <param name="a1Range">The A1 range to fill.</param>
        /// <param name="fromHexColor">The gradient start color expressed as an ARGB or RGB hex string.</param>
        /// <param name="toHexColor">The gradient end color expressed as an ARGB or RGB hex string.</param>
        /// <param name="degree">The linear gradient angle in degrees.</param>
        public void FillRangeGradient(string a1Range, string fromHexColor, string toHexColor, double degree = 0) {
            if (string.IsNullOrWhiteSpace(fromHexColor) || string.IsNullOrWhiteSpace(toHexColor)) {
                return;
            }

            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);

            if (!_excelDocument.IsMaterializingDeferredDataSetImport) {
                MaterializeDeferredDataSetImportIfNeeded();
            }

            WriteLock(() => FillRangeGradientCore(r1, c1, r2, c2, fromHexColor, toHexColor, degree));
        }

        private void ApplyGradientBackground(Cell cell, string fromHexColor, string toHexColor, double degree) {
            var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
            var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<DocumentFormat.OpenXml.Packaging.WorkbookStylesPart>();
            var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            EnsureDefaultStylePrimitives(stylesheet);

            uint fillId = GetOrCreateFill(stylesheet, CreateGradientFill(fromHexColor, toHexColor, degree));
            ApplyCellFormatOverride(stylesheet, cell, format => {
                format.FillId = fillId;
                format.ApplyFill = true;
            });

            stylesPart.Stylesheet.Save();
        }

        private void FillRangeGradientCore(int firstRow, int firstColumn, int lastRow, int lastColumn, string fromHexColor, string toHexColor, double degree) {
            var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
            var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<DocumentFormat.OpenXml.Packaging.WorkbookStylesPart>();
            var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            EnsureDefaultStylePrimitives(stylesheet);

            uint fillId = GetOrCreateFill(stylesheet, CreateGradientFill(fromHexColor, toHexColor, degree));
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

            stylesPart.Stylesheet.Save();
        }

        private static Fill CreateGradientFill(string fromHexColor, string toHexColor, double degree) {
            return new Fill(new GradientFill(
                new GradientStop(new Color { Rgb = NormalizeHexColor(fromHexColor) }) { Position = 0D },
                new GradientStop(new Color { Rgb = NormalizeHexColor(toHexColor) }) { Position = 1D }) {
                Type = GradientValues.Linear,
                Degree = degree
            });
        }
    }
}
