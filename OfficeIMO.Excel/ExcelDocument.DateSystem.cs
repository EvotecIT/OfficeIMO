using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Gets or sets the workbook date system used for serialized date values.
        /// </summary>
        public ExcelDateSystem DateSystem {
            get {
                if (_deferredWorkbook is { } deferred) return deferred.Date1904 ? ExcelDateSystem.NineteenFour : ExcelDateSystem.NineteenHundred;
                WorkbookProperties? properties = WorkbookRoot.GetFirstChild<WorkbookProperties>();
                return properties?.Date1904?.Value == true
                    ? ExcelDateSystem.NineteenFour
                    : ExcelDateSystem.NineteenHundred;
            }
            set {
                if (value != ExcelDateSystem.NineteenHundred && value != ExcelDateSystem.NineteenFour) {
                    throw new ArgumentOutOfRangeException(nameof(value), value, "Unsupported Excel date system.");
                }

                Workbook workbook = WorkbookRoot;
                WorkbookProperties? properties = workbook.GetFirstChild<WorkbookProperties>();
                if (properties == null) {
                    properties = new WorkbookProperties();
                    OpenXmlWorkbookElementOrder.InsertInOrder(workbook, properties);
                }

                ExcelDateSystem current = properties.Date1904?.Value == true
                    ? ExcelDateSystem.NineteenFour
                    : ExcelDateSystem.NineteenHundred;
                if (current == value) {
                    return;
                }

                ConvertExistingDateSerials(current, value);
                bool use1904 = value == ExcelDateSystem.NineteenFour;
                properties.Date1904 = use1904 ? true : null;
                RefreshDeferredDirectDataSetDateSystem(value);
                WorkbookRoot.Save();
                MarkPackageDirty();
            }
        }

        private void ConvertExistingDateSerials(ExcelDateSystem current, ExcelDateSystem target) {
            if (current == target || _spreadSheetDocument == null) {
                return;
            }

            StylesCache styles = StylesCache.Build(_spreadSheetDocument);
            if (!styles.HasDateStyles) {
                return;
            }

            long offset = target == ExcelDateSystem.NineteenFour
                ? -(long)ExcelDateSystemConverter.Date1904OffsetDays
                : (long)ExcelDateSystemConverter.Date1904OffsetDays;

            foreach (WorksheetPart worksheetPart in WorkbookPartRoot.WorksheetParts) {
                Worksheet? worksheet = worksheetPart.Worksheet;
                if (worksheet == null) {
                    continue;
                }

                bool changed = false;
                foreach (Cell cell in worksheet.Descendants<Cell>()) {
                    if (cell.CellFormula != null || !IsNumericDateCell(cell, styles)) {
                        continue;
                    }

                    string? text = cell.CellValue?.Text;
                    if (!OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double serial)) {
                        continue;
                    }

                    cell.CellValue = new CellValue(InvariantNumberText.Get(ExcelDateSystemConverter.ShiftSerialDays(serial, offset)));
                    changed = true;
                }

                if (changed) {
                    worksheet.Save();
                }
            }
        }

        private static bool IsNumericDateCell(Cell cell, StylesCache styles) {
            if (cell.StyleIndex?.Value is not uint styleIndex || !styles.IsDateLike(styleIndex)) {
                return false;
            }

            CellValues? dataType = cell.DataType?.Value;
            return dataType == null || dataType == CellValues.Number;
        }
    }
}
