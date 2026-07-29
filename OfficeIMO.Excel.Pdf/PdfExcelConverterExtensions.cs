using System.Data;
using System.Threading;
using System.Threading.Tasks;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    /// <summary>
    /// Converts structured logical PDF tables into Excel worksheets.
    /// </summary>
    public static class PdfExcelConverterExtensions {
        /// <summary>Imports logical PDF tables from an opened PDF into a new editable Excel document.</summary>
        public static ExcelDocument ToExcelDocument(
            this PdfCore.PdfDocument document,
            PdfExcelImportOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return document.Read.Logical().ToExcelDocument(options);
        }

        /// <summary>Imports logical PDF tables from an opened PDF into an editable Excel document plus an explicit table-scope report.</summary>
        public static PdfExcelImportResult ToExcelDocumentResult(
            this PdfCore.PdfDocument document,
            PdfExcelImportOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return document.Read.Logical().ToExcelDocumentResult(options);
        }

        /// <summary>Imports logical PDF tables from an opened PDF into a new Excel workbook.</summary>
        public static PdfExcelImportReport SaveAsExcel(
            this PdfCore.PdfDocument document,
            string workbookPath,
            PdfExcelImportOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return document.Read.Logical().SaveAsExcel(workbookPath, options);
        }

        /// <summary>Imports logical PDF tables from an opened PDF into a caller-owned workbook stream.</summary>
        public static PdfExcelImportReport SaveAsExcel(
            this PdfCore.PdfDocument document,
            Stream workbookStream,
            PdfExcelImportOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return document.Read.Logical().SaveAsExcel(workbookStream, options);
        }

        /// <summary>Imports logical PDF tables from an opened PDF and asynchronously saves a new Excel workbook.</summary>
        public static Task<PdfExcelImportReport> SaveAsExcelAsync(
            this PdfCore.PdfDocument document,
            string workbookPath,
            PdfExcelImportOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return document.Read.Logical().SaveAsExcelAsync(workbookPath, options, cancellationToken);
        }

        /// <summary>Imports logical PDF tables from an opened PDF and asynchronously saves to a caller-owned workbook stream.</summary>
        public static Task<PdfExcelImportReport> SaveAsExcelAsync(
            this PdfCore.PdfDocument document,
            Stream workbookStream,
            PdfExcelImportOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return document.Read.Logical().SaveAsExcelAsync(workbookStream, options, cancellationToken);
        }

        /// <summary>Imports logical PDF tables into a new Excel workbook at <paramref name="workbookPath"/>.</summary>
        public static PdfExcelImportReport SaveAsExcel(
            this PdfCore.PdfLogicalDocument document,
            string workbookPath,
            PdfExcelImportOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (string.IsNullOrWhiteSpace(workbookPath)) throw new ArgumentException("Workbook path cannot be empty.", nameof(workbookPath));

            PdfExcelImportResult result = document.ToExcelDocumentResult(options);
            using (result.Value) {
                result.Value.Save(workbookPath);
            }
            return result.Report;
        }

        /// <summary>Imports logical PDF tables into an Excel workbook written to a caller-owned stream.</summary>
        public static PdfExcelImportReport SaveAsExcel(
            this PdfCore.PdfLogicalDocument document,
            Stream workbookStream,
            PdfExcelImportOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (workbookStream == null) throw new ArgumentNullException(nameof(workbookStream));
            if (!workbookStream.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(workbookStream));

            PdfExcelImportResult result = document.ToExcelDocumentResult(options);
            using (result.Value) {
                result.Value.Save(workbookStream);
            }
            return result.Report;
        }

        /// <summary>Imports logical PDF tables into a new editable Excel document.</summary>
        public static ExcelDocument ToExcelDocument(
            this PdfCore.PdfLogicalDocument document,
            PdfExcelImportOptions? options = null) => document.ToExcelDocumentResult(options).Value;

        /// <summary>Imports logical PDF tables into an editable Excel document plus an explicit table-scope report.</summary>
        public static PdfExcelImportResult ToExcelDocumentResult(
            this PdfCore.PdfLogicalDocument document,
            PdfExcelImportOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            ExcelDocument workbook = ExcelDocument.Create();
            IReadOnlyList<PdfExcelImportEntry> entries = ImportTables(document, workbook, options ?? new PdfExcelImportOptions());
            PdfCore.PdfTableExtractionScopeReport sourceScope = PdfCore.PdfLogicalTableAnalysis.AnalyzeExtractionScope(document);
            return new PdfExcelImportResult(workbook, new PdfExcelImportReport(entries, sourceScope));
        }

        /// <summary>Asynchronously imports logical PDF tables into an Excel workbook written to a file.</summary>
        public static async Task<PdfExcelImportReport> SaveAsExcelAsync(
            this PdfCore.PdfLogicalDocument document,
            string workbookPath,
            PdfExcelImportOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (string.IsNullOrWhiteSpace(workbookPath)) throw new ArgumentException("Workbook path cannot be empty.", nameof(workbookPath));
            cancellationToken.ThrowIfCancellationRequested();
            PdfExcelImportResult result = document.ToExcelDocumentResult(options);
            using (result.Value) {
                await result.Value.SaveAsync(workbookPath, cancellationToken: cancellationToken).ConfigureAwait(false);
            }
            return result.Report;
        }

        /// <summary>Asynchronously imports logical PDF tables into an Excel workbook written to a caller-owned stream.</summary>
        public static async Task<PdfExcelImportReport> SaveAsExcelAsync(
            this PdfCore.PdfLogicalDocument document,
            Stream workbookStream,
            PdfExcelImportOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (workbookStream == null) throw new ArgumentNullException(nameof(workbookStream));
            if (!workbookStream.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(workbookStream));
            cancellationToken.ThrowIfCancellationRequested();
            PdfExcelImportResult result = document.ToExcelDocumentResult(options);
            using (result.Value) {
                await result.Value.SaveAsync(workbookStream, cancellationToken).ConfigureAwait(false);
            }
            return result.Report;
        }

        private static IReadOnlyList<PdfExcelImportEntry> ImportTables(
            PdfCore.PdfLogicalDocument document,
            ExcelDocument workbook,
            PdfExcelImportOptions options) {
            IReadOnlyList<PdfCore.PdfLogicalTableExtraction> tables = PdfCore.PdfLogicalTableAnalysis.ExtractTables(document, options.MaxRows);
            if (tables.Count == 0) {
                AddEmptyWorkbookSheet(workbook, options);
                return Array.Empty<PdfExcelImportEntry>();
            }

            var results = new List<PdfExcelImportEntry>(tables.Count);
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
                results.Add(new PdfExcelImportEntry(
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

        private static void AddEmptyWorkbookSheet(ExcelDocument workbook, PdfExcelImportOptions options) {
            ExcelSheet sheet = workbook.AddWorksheet(options.EmptyWorkbookSheetName, SheetNameValidationMode.Sanitize);
            sheet.CellValue(1, 1, "No PDF tables detected.");
        }

        private static DataTable ToDataTable(string tableName, PdfCore.PdfLogicalTableData data, PdfExcelImportOptions options) {
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
