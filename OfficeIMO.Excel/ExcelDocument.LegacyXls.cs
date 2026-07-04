using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Loads a legacy binary `.xls` workbook and projects supported content into a normal OfficeIMO Excel document.
        /// The resulting document can save as Open XML `.xlsx` or as a supported native `.xls` subset.
        /// </summary>
        public static ExcelDocument LoadLegacyXls(string path, LegacyXlsImportOptions? options = null) {
            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(path, options);
            return ProjectLoadedLegacyXlsWorkbook(workbook, path);
        }

        /// <summary>
        /// Loads a legacy binary `.xls` workbook and returns both the projected OfficeIMO document and the import report.
        /// The resulting document can save as Open XML `.xlsx` or as a supported native `.xls` subset.
        /// </summary>
        public static LegacyXlsLoadResult LoadLegacyXlsWithReport(string path, LegacyXlsImportOptions? options = null) {
            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(path, options);
            return CreateLegacyXlsLoadResult(workbook, path);
        }

        /// <summary>
        /// Loads a legacy binary `.xls` workbook stream and projects supported content into a normal OfficeIMO Excel document.
        /// The resulting document can save as Open XML `.xlsx` or as a supported native `.xls` subset.
        /// </summary>
        public static ExcelDocument LoadLegacyXls(Stream stream, LegacyXlsImportOptions? options = null) {
            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(stream, options);
            return ProjectLoadedLegacyXlsWorkbook(workbook, sourcePath: null);
        }

        /// <summary>
        /// Loads a legacy binary `.xls` workbook stream and returns both the projected OfficeIMO document and the import report.
        /// The resulting document can save as Open XML `.xlsx` or as a supported native `.xls` subset.
        /// </summary>
        public static LegacyXlsLoadResult LoadLegacyXlsWithReport(Stream stream, LegacyXlsImportOptions? options = null) {
            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(stream, options);
            return CreateLegacyXlsLoadResult(workbook, sourcePath: null);
        }

        private static LegacyXlsLoadResult CreateLegacyXlsLoadResult(LegacyXlsWorkbook workbook, string? sourcePath) {
            try {
                return new LegacyXlsLoadResult(ProjectLoadedLegacyXlsWorkbook(workbook, sourcePath), workbook);
            } catch (InvalidDataException exception) {
                return new LegacyXlsLoadResult(document: null, workbook, exception);
            }
        }
    }
}
