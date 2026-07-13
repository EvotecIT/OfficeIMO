using System.Data;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    /// <summary>
    /// Converts structured logical PDF tables into Excel worksheets.
    /// </summary>
    public static class PdfExcelTableConverterExtensions {
        /// <summary>
        /// Extracts logical PDF tables into a new Excel workbook written to <paramref name="workbookPath"/>.
        /// </summary>
        /// <param name="document">Logical PDF document to import.</param>
        /// <param name="workbookPath">Destination workbook path.</param>
        /// <param name="options">Optional import settings.</param>
        /// <returns>Metadata for every imported table.</returns>
        public static IReadOnlyList<PdfExcelTableImportResult> SaveAsExcelFromPdfTables(
            this PdfCore.PdfLogicalDocument document,
            string workbookPath,
            PdfExcelTableImportOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (string.IsNullOrWhiteSpace(workbookPath)) throw new ArgumentException("Workbook path cannot be empty.", nameof(workbookPath));

            using ExcelDocument workbook = ExcelDocument.Create(workbookPath);
            IReadOnlyList<PdfExcelTableImportResult> results = ImportTables(document, workbook, options ?? new PdfExcelTableImportOptions());
            workbook.Save();
            return results;
        }

        /// <summary>
        /// Extracts logical PDF tables into a new Excel workbook written to <paramref name="workbookStream"/>.
        /// </summary>
        /// <param name="document">Logical PDF document to import.</param>
        /// <param name="workbookStream">Writable destination stream for the workbook package.</param>
        /// <param name="options">Optional import settings.</param>
        /// <returns>Metadata for every imported table.</returns>
        public static IReadOnlyList<PdfExcelTableImportResult> SaveAsExcelFromPdfTables(
            this PdfCore.PdfLogicalDocument document,
            Stream workbookStream,
            PdfExcelTableImportOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (workbookStream == null) throw new ArgumentNullException(nameof(workbookStream));

            using ExcelDocument workbook = ExcelDocument.Create(workbookStream);
            IReadOnlyList<PdfExcelTableImportResult> results = ImportTables(document, workbook, options ?? new PdfExcelTableImportOptions());
            workbook.Save();
            return results;
        }

        /// <summary>
        /// Extracts logical PDF tables into Excel workbook bytes.
        /// </summary>
        /// <param name="document">Logical PDF document to import.</param>
        /// <param name="options">Optional import settings.</param>
        /// <returns>Workbook package bytes.</returns>
        public static byte[] ToExcelBytesFromPdfTables(
            this PdfCore.PdfLogicalDocument document,
            PdfExcelTableImportOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));

            using var stream = new MemoryStream();
            document.SaveAsExcelFromPdfTables(stream, options);
            return stream.ToArray();
        }

        /// <summary>
        /// Loads a PDF file, extracts logical tables, and writes them to a new Excel workbook.
        /// </summary>
        /// <param name="pdfPath">Source PDF path.</param>
        /// <param name="workbookPath">Destination workbook path.</param>
        /// <param name="options">Optional import settings.</param>
        /// <returns>Metadata for every imported table.</returns>
        public static IReadOnlyList<PdfExcelTableImportResult> SaveAsExcelFromPdfTables(
            string pdfPath,
            string workbookPath,
            PdfExcelTableImportOptions? options = null) {
            if (string.IsNullOrWhiteSpace(pdfPath)) throw new ArgumentException("PDF path cannot be empty.", nameof(pdfPath));
            if (string.IsNullOrWhiteSpace(workbookPath)) throw new ArgumentException("Workbook path cannot be empty.", nameof(workbookPath));

            options ??= new PdfExcelTableImportOptions();
            PdfCore.PdfLogicalDocument document = LoadPdf(pdfPath, options);
            return document.SaveAsExcelFromPdfTables(workbookPath, options);
        }

        /// <summary>
        /// Loads PDF bytes, extracts logical tables, and writes them to a new Excel workbook stream.
        /// </summary>
        /// <param name="pdfBytes">Source PDF bytes.</param>
        /// <param name="workbookStream">Writable destination stream for the workbook package.</param>
        /// <param name="options">Optional import settings.</param>
        /// <returns>Metadata for every imported table.</returns>
        public static IReadOnlyList<PdfExcelTableImportResult> SaveAsExcelFromPdfTables(
            byte[] pdfBytes,
            Stream workbookStream,
            PdfExcelTableImportOptions? options = null) {
            if (pdfBytes == null) throw new ArgumentNullException(nameof(pdfBytes));
            if (workbookStream == null) throw new ArgumentNullException(nameof(workbookStream));

            options ??= new PdfExcelTableImportOptions();
            PdfCore.PdfLogicalDocument document = LoadPdf(pdfBytes, options);
            return document.SaveAsExcelFromPdfTables(workbookStream, options);
        }

        /// <summary>
        /// Loads a PDF stream, extracts logical tables, and writes them to a new Excel workbook stream.
        /// </summary>
        /// <param name="pdfStream">Readable source PDF stream.</param>
        /// <param name="workbookStream">Writable destination stream for the workbook package.</param>
        /// <param name="options">Optional import settings.</param>
        /// <returns>Metadata for every imported table.</returns>
        public static IReadOnlyList<PdfExcelTableImportResult> SaveAsExcelFromPdfTables(
            Stream pdfStream,
            Stream workbookStream,
            PdfExcelTableImportOptions? options = null) {
            if (pdfStream == null) throw new ArgumentNullException(nameof(pdfStream));
            if (workbookStream == null) throw new ArgumentNullException(nameof(workbookStream));

            options ??= new PdfExcelTableImportOptions();
            PdfCore.PdfLogicalDocument document = LoadPdf(pdfStream, options);
            return document.SaveAsExcelFromPdfTables(workbookStream, options);
        }

        private static IReadOnlyList<PdfExcelTableImportResult> ImportTables(
            PdfCore.PdfLogicalDocument document,
            ExcelDocument workbook,
            PdfExcelTableImportOptions options) {
            IReadOnlyList<PdfCore.PdfLogicalTableExtraction> tables = PdfCore.PdfLogicalTableAnalysis.ExtractTables(document, options.MaxRows);
            if (tables.Count == 0) {
                AddEmptyWorkbookSheet(workbook, options);
                return Array.Empty<PdfExcelTableImportResult>();
            }

            var results = new List<PdfExcelTableImportResult>(tables.Count);
            for (int i = 0; i < tables.Count; i++) {
                PdfCore.PdfLogicalTableExtraction extraction = tables[i];
                PdfCore.PdfLogicalTableData data = extraction.Data;
                string requestedTableName = BuildTableName(options.TableNamePrefix, extraction, i);
                DataTable dataTable = ToDataTable(requestedTableName, data, options);
                ExcelSheet sheet = workbook.AddWorksheet(BuildSheetName(options.SheetNamePrefix, extraction, i), SheetNameValidationMode.Sanitize);
                string range = sheet.InsertDataTableAsTable(
                    dataTable,
                    tableName: requestedTableName,
                    style: options.TableStyle,
                    includeAutoFilter: options.IncludeAutoFilter);

                if (options.AutoFitColumns) {
                    sheet.AutoFitColumns();
                }

                string actualTableName = FindActualTableName(workbook, sheet.Name, range, requestedTableName);
                results.Add(new PdfExcelTableImportResult(
                    extraction.PageIndex,
                    extraction.PageNumber,
                    extraction.TableIndex,
                    extraction.DetectionKind,
                    sheet.Name,
                    actualTableName,
                    range,
                    data.Columns.Count,
                    data.Rows.Count,
                    data.TotalRowCount,
                    data.Truncated));
            }

            return results.AsReadOnly();
        }

        private static PdfCore.PdfLogicalDocument LoadPdf(string path, PdfExcelTableImportOptions options) {
            PdfCore.PdfPageRange[] ranges = GetPageRanges(options);
            return ranges.Length == 0
                ? PdfCore.PdfLogicalDocument.Load(path, options.LayoutOptions)
                : PdfCore.PdfLogicalDocument.LoadPageRanges(path, options.LayoutOptions, ranges);
        }

        private static PdfCore.PdfLogicalDocument LoadPdf(byte[] pdfBytes, PdfExcelTableImportOptions options) {
            PdfCore.PdfPageRange[] ranges = GetPageRanges(options);
            return ranges.Length == 0
                ? PdfCore.PdfLogicalDocument.Load(pdfBytes, options.LayoutOptions)
                : PdfCore.PdfLogicalDocument.LoadPageRanges(pdfBytes, options.LayoutOptions, ranges);
        }

        private static PdfCore.PdfLogicalDocument LoadPdf(Stream stream, PdfExcelTableImportOptions options) {
            PdfCore.PdfPageRange[] ranges = GetPageRanges(options);
            return ranges.Length == 0
                ? PdfCore.PdfLogicalDocument.Load(stream, options.LayoutOptions)
                : PdfCore.PdfLogicalDocument.LoadPageRanges(stream, options.LayoutOptions, ranges);
        }

        private static PdfCore.PdfPageRange[] GetPageRanges(PdfExcelTableImportOptions options) {
            return options.PageRanges == null || options.PageRanges.Count == 0
                ? Array.Empty<PdfCore.PdfPageRange>()
                : options.PageRanges.ToArray();
        }

        private static void AddEmptyWorkbookSheet(ExcelDocument workbook, PdfExcelTableImportOptions options) {
            ExcelSheet sheet = workbook.AddWorksheet(options.EmptyWorkbookSheetName, SheetNameValidationMode.Sanitize);
            sheet.CellValue(1, 1, "No PDF tables detected.");
        }

        private static DataTable ToDataTable(string tableName, PdfCore.PdfLogicalTableData data, PdfExcelTableImportOptions options) {
            var table = new DataTable(tableName) {
                Locale = CultureInfo.InvariantCulture
            };

            bool[] numericColumns = options.ConvertNumericColumns
                ? PdfCore.PdfLogicalTableAnalysis.DetectParsableNumericColumns(data, options.NumericCulture)
                : new bool[data.Columns.Count];
            var usedColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < data.Columns.Count; i++) {
                Type columnType = numericColumns[i] ? typeof(decimal) : typeof(string);
                table.Columns.Add(GetUniqueColumnName(data.Columns[i], i, usedColumns), columnType);
            }

            table.BeginLoadData();
            try {
                for (int rowIndex = 0; rowIndex < data.Rows.Count; rowIndex++) {
                    DataRow row = table.NewRow();
                    IReadOnlyList<string> sourceRow = data.Rows[rowIndex];
                    for (int columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++) {
                        string value = columnIndex < sourceRow.Count ? sourceRow[columnIndex] : string.Empty;
                        if (numericColumns[columnIndex]) {
                            row[columnIndex] = PdfCore.PdfLogicalTableAnalysis.TryParseNumericValue(value, options.NumericCulture, out decimal number)
                                ? number
                                : DBNull.Value;
                        } else {
                            row[columnIndex] = value;
                        }
                    }

                    table.Rows.Add(row);
                }
            } finally {
                table.EndLoadData();
            }

            return table;
        }

        private static string GetUniqueColumnName(string? value, int index, ISet<string> usedColumns) {
            string baseName = string.IsNullOrWhiteSpace(value)
                ? "Column" + (index + 1).ToString(CultureInfo.InvariantCulture)
                : value!.Trim();
            string candidate = baseName;
            int suffix = 2;
            while (!usedColumns.Add(candidate)) {
                candidate = baseName + " " + suffix.ToString(CultureInfo.InvariantCulture);
                suffix++;
            }

            return candidate;
        }

        private static string BuildSheetName(string? prefix, PdfCore.PdfLogicalTableExtraction extraction, int importIndex) {
            string normalizedPrefix = string.IsNullOrWhiteSpace(prefix) ? "PDF" : prefix!.Trim();
            return normalizedPrefix
                + " P" + extraction.PageNumber.ToString(CultureInfo.InvariantCulture)
                + " T" + (extraction.TableIndex + 1).ToString(CultureInfo.InvariantCulture)
                + " #" + (importIndex + 1).ToString(CultureInfo.InvariantCulture);
        }

        private static string BuildTableName(string? prefix, PdfCore.PdfLogicalTableExtraction extraction, int importIndex) {
            string normalizedPrefix = NormalizeIdentifierPrefix(prefix, "PdfTable");
            return normalizedPrefix
                + "_P" + extraction.PageNumber.ToString(CultureInfo.InvariantCulture)
                + "_T" + (extraction.TableIndex + 1).ToString(CultureInfo.InvariantCulture)
                + "_" + (importIndex + 1).ToString(CultureInfo.InvariantCulture);
        }

        private static string NormalizeIdentifierPrefix(string? prefix, string fallback) {
            string source = string.IsNullOrWhiteSpace(prefix) ? fallback : prefix!.Trim();
            var chars = new char[source.Length + 1];
            int count = 0;
            for (int i = 0; i < source.Length; i++) {
                char ch = source[i];
                chars[count++] = char.IsLetterOrDigit(ch) || ch == '_' ? ch : '_';
            }

            string normalized = new string(chars, 0, count).Trim('_');
            if (normalized.Length == 0) {
                normalized = fallback;
            }

            if (!char.IsLetter(normalized[0]) && normalized[0] != '_') {
                normalized = "_" + normalized;
            }

            return normalized;
        }

        private static string FindActualTableName(ExcelDocument workbook, string sheetName, string range, string fallback) {
            ExcelTableInfo? table = workbook.GetTables()
                .LastOrDefault(candidate =>
                    string.Equals(candidate.SheetName, sheetName, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(candidate.Range, range, StringComparison.OrdinalIgnoreCase));
            return table?.Name ?? fallback;
        }
    }
}
