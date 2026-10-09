using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Removes the previous value, formula and inline text before assigning
        /// tabular replacement data, without changing the cell's style.
        /// </summary>
        private void SetPreparedTabularCellValue(Cell cell, CellValue value, DocumentFormat.OpenXml.EnumValue<CellValues> type) {
            ClearTabularReplacementValueMetadata(cell);
            cell.CellValue = value;
            cell.DataType = type;
        }

        private void SetMissingTabularCellValue(Cell cell) {
            ClearTabularReplacementValueMetadata(cell);
            cell.CellValue = null;
            cell.DataType = null;
            // An imported missing value still occupies its tabular coordinate.
            // The explicit default format keeps native blank-cell writers from
            // treating it as an incidental unstyled stub left by ClearRange.
            cell.StyleIndex ??= GetMissingTabularCellStyleIndex();
        }

        private uint GetMissingTabularCellStyleIndex() {
            WorkbookPart workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
            WorkbookStylesPart stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
            Stylesheet stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            if (stylesheet.CellFormats == null || !stylesheet.CellFormats.Elements<CellFormat>().Any()) {
                // The explicit native presence marker is also a style reference
                // in XLSX, so its default cell format must exist before saving.
                EnsureDefaultStylePrimitives(stylesheet);
                SaveStylesheet(stylesPart);
            }
            return 0U;
        }

        /// <summary>
        /// Removes stale formula and inline text metadata while retaining an
        /// already validated replacement value and the existing style.
        /// </summary>
        private void ClearTabularReplacementValueMetadata(Cell cell) {
            ClearCellValueMetadata(cell);
            cell.CellFormula = null;
            cell.InlineString = null;
        }
    }
}
