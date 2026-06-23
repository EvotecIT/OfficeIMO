using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Loads a legacy binary `.xls` workbook and projects supported content into a normal OfficeIMO Excel document.
        /// The resulting document saves as Open XML `.xlsx`; native `.xls` save is intentionally out of scope.
        /// </summary>
        public static ExcelDocument LoadLegacyXls(string path, LegacyXlsImportOptions? options = null) {
            return LegacyXlsWorkbook.Load(path, options).ToExcelDocument();
        }

        /// <summary>
        /// Loads a legacy binary `.xls` workbook and returns both the projected OfficeIMO document and the import report.
        /// The resulting document saves as Open XML `.xlsx`; native `.xls` save is intentionally out of scope.
        /// </summary>
        public static LegacyXlsLoadResult LoadLegacyXlsWithReport(string path, LegacyXlsImportOptions? options = null) {
            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(path, options);
            return new LegacyXlsLoadResult(workbook.ToExcelDocument(), workbook);
        }

        /// <summary>
        /// Loads a legacy binary `.xls` workbook stream and projects supported content into a normal OfficeIMO Excel document.
        /// The resulting document saves as Open XML `.xlsx`; native `.xls` save is intentionally out of scope.
        /// </summary>
        public static ExcelDocument LoadLegacyXls(Stream stream, LegacyXlsImportOptions? options = null) {
            return LegacyXlsWorkbook.Load(stream, options).ToExcelDocument();
        }

        /// <summary>
        /// Loads a legacy binary `.xls` workbook stream and returns both the projected OfficeIMO document and the import report.
        /// The resulting document saves as Open XML `.xlsx`; native `.xls` save is intentionally out of scope.
        /// </summary>
        public static LegacyXlsLoadResult LoadLegacyXlsWithReport(Stream stream, LegacyXlsImportOptions? options = null) {
            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(stream, options);
            return new LegacyXlsLoadResult(workbook.ToExcelDocument(), workbook);
        }
    }
}
