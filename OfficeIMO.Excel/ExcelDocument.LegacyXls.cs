using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Loads a legacy binary `.xls` workbook and projects supported content into a normal OfficeIMO Excel document.
        /// The resulting document saves as Open XML `.xlsx`; native `.xls` save is intentionally out of scope.
        /// </summary>
        public static ExcelDocument LoadLegacyXls(string path, LegacyXlsImportOptions? options = null) {
            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(path, options);
            ExcelDocument document = workbook.ToExcelDocument();
            document.MarkLoadedFromLegacyXls(path, workbook);
            return document;
        }

        /// <summary>
        /// Loads a legacy binary `.xls` workbook and returns both the projected OfficeIMO document and the import report.
        /// The resulting document saves as Open XML `.xlsx`; native `.xls` save is intentionally out of scope.
        /// </summary>
        public static LegacyXlsLoadResult LoadLegacyXlsWithReport(string path, LegacyXlsImportOptions? options = null) {
            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(path, options);
            ExcelDocument document = workbook.ToExcelDocument();
            document.MarkLoadedFromLegacyXls(path, workbook);
            return new LegacyXlsLoadResult(document, workbook);
        }

        /// <summary>
        /// Loads a legacy binary `.xls` workbook stream and projects supported content into a normal OfficeIMO Excel document.
        /// The resulting document saves as Open XML `.xlsx`; native `.xls` save is intentionally out of scope.
        /// </summary>
        public static ExcelDocument LoadLegacyXls(Stream stream, LegacyXlsImportOptions? options = null) {
            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(stream, options);
            ExcelDocument document = workbook.ToExcelDocument();
            document.MarkLoadedFromLegacyXls(sourcePath: null, workbook);
            return document;
        }

        /// <summary>
        /// Loads a legacy binary `.xls` workbook stream and returns both the projected OfficeIMO document and the import report.
        /// The resulting document saves as Open XML `.xlsx`; native `.xls` save is intentionally out of scope.
        /// </summary>
        public static LegacyXlsLoadResult LoadLegacyXlsWithReport(Stream stream, LegacyXlsImportOptions? options = null) {
            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(stream, options);
            ExcelDocument document = workbook.ToExcelDocument();
            document.MarkLoadedFromLegacyXls(sourcePath: null, workbook);
            return new LegacyXlsLoadResult(document, workbook);
        }
    }
}
