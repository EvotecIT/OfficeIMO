using System.Data;
using System.Threading;
using System.Threading.Tasks;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    /// <summary>
    /// Converts structured logical PDF tables into Excel worksheets.
    /// </summary>
    public static class PdfExcelTableConverterExtensions {
        /// <summary>Imports logical PDF tables from an opened PDF into a new editable Excel document.</summary>
        public static ExcelDocument ImportTablesToExcelDocument(
            this PdfCore.PdfDocument document,
            PdfExcelTableImportOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return document.Read.Logical().ImportTablesToExcelDocument(options);
        }

        /// <summary>Imports logical PDF tables from an opened PDF into an editable Excel document plus an explicit table-scope report.</summary>
        public static PdfExcelTableImportResult ImportTablesToExcelDocumentResult(
            this PdfCore.PdfDocument document,
            PdfExcelTableImportOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return document.Read.Logical().ImportTablesToExcelDocumentResult(options);
        }

        /// <summary>Imports logical PDF tables from an opened PDF into a new Excel workbook.</summary>
        public static PdfExcelTableImportReport SaveTablesAsExcel(
            this PdfCore.PdfDocument document,
            string workbookPath,
            PdfExcelTableImportOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return document.Read.Logical().SaveTablesAsExcel(workbookPath, options);
        }

        /// <summary>Imports logical PDF tables from an opened PDF into a caller-owned workbook stream.</summary>
        public static PdfExcelTableImportReport SaveTablesAsExcel(
            this PdfCore.PdfDocument document,
            Stream workbookStream,
            PdfExcelTableImportOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return document.Read.Logical().SaveTablesAsExcel(workbookStream, options);
        }

        /// <summary>Imports logical PDF tables from an opened PDF and asynchronously saves a new Excel workbook.</summary>
        public static Task<PdfExcelTableImportReport> SaveTablesAsExcelAsync(
            this PdfCore.PdfDocument document,
            string workbookPath,
            PdfExcelTableImportOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return document.Read.Logical().SaveTablesAsExcelAsync(workbookPath, options, cancellationToken);
        }

        /// <summary>Imports logical PDF tables from an opened PDF and asynchronously saves to a caller-owned workbook stream.</summary>
        public static Task<PdfExcelTableImportReport> SaveTablesAsExcelAsync(
            this PdfCore.PdfDocument document,
            Stream workbookStream,
            PdfExcelTableImportOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return document.Read.Logical().SaveTablesAsExcelAsync(workbookStream, options, cancellationToken);
        }

        /// <summary>Imports logical PDF tables into a new Excel workbook at <paramref name="workbookPath"/>.</summary>
        public static PdfExcelTableImportReport SaveTablesAsExcel(
            this PdfCore.PdfLogicalDocument document,
            string workbookPath,
            PdfExcelTableImportOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (string.IsNullOrWhiteSpace(workbookPath)) throw new ArgumentException("Workbook path cannot be empty.", nameof(workbookPath));

            PdfExcelTableImportResult result = document.ImportTablesToExcelDocumentResult(options);
            using (result.Value) {
                result.Value.Save(workbookPath);
            }
            return result.Report;
        }

        /// <summary>Imports logical PDF tables into an Excel workbook written to a caller-owned stream.</summary>
        public static PdfExcelTableImportReport SaveTablesAsExcel(
            this PdfCore.PdfLogicalDocument document,
            Stream workbookStream,
            PdfExcelTableImportOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (workbookStream == null) throw new ArgumentNullException(nameof(workbookStream));
            if (!workbookStream.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(workbookStream));

            PdfExcelTableImportResult result = document.ImportTablesToExcelDocumentResult(options);
            using (result.Value) {
                result.Value.Save(workbookStream);
            }
            return result.Report;
        }

        /// <summary>Imports logical PDF tables into a new editable Excel document.</summary>
        public static ExcelDocument ImportTablesToExcelDocument(
            this PdfCore.PdfLogicalDocument document,
            PdfExcelTableImportOptions? options = null) => document.ImportTablesToExcelDocumentResult(options).Value;

        /// <summary>Imports logical PDF tables into an editable Excel document plus an explicit table-scope report.</summary>
        public static PdfExcelTableImportResult ImportTablesToExcelDocumentResult(
            this PdfCore.PdfLogicalDocument document,
            PdfExcelTableImportOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            ExcelDocument workbook = ExcelDocument.Create();
            IReadOnlyList<PdfExcelTableImportEntry> entries = ImportTables(document, workbook, options ?? new PdfExcelTableImportOptions());
            PdfCore.PdfTableExtractionScopeReport sourceScope = PdfCore.PdfLogicalTableAnalysis.AnalyzeExtractionScope(document);
            return new PdfExcelTableImportResult(workbook, new PdfExcelTableImportReport(entries, sourceScope));
        }

        /// <summary>Asynchronously imports logical PDF tables into an Excel workbook written to a file.</summary>
        public static async Task<PdfExcelTableImportReport> SaveTablesAsExcelAsync(
            this PdfCore.PdfLogicalDocument document,
            string workbookPath,
            PdfExcelTableImportOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (string.IsNullOrWhiteSpace(workbookPath)) throw new ArgumentException("Workbook path cannot be empty.", nameof(workbookPath));
            cancellationToken.ThrowIfCancellationRequested();
            PdfExcelTableImportResult result = document.ImportTablesToExcelDocumentResult(options);
            using (result.Value) {
                await result.Value.SaveAsync(workbookPath, cancellationToken: cancellationToken).ConfigureAwait(false);
            }
            return result.Report;
        }

        /// <summary>Asynchronously imports logical PDF tables into an Excel workbook written to a caller-owned stream.</summary>
        public static async Task<PdfExcelTableImportReport> SaveTablesAsExcelAsync(
            this PdfCore.PdfLogicalDocument document,
            Stream workbookStream,
            PdfExcelTableImportOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (workbookStream == null) throw new ArgumentNullException(nameof(workbookStream));
            if (!workbookStream.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(workbookStream));
            cancellationToken.ThrowIfCancellationRequested();
            PdfExcelTableImportResult result = document.ImportTablesToExcelDocumentResult(options);
            using (result.Value) {
                await result.Value.SaveAsync(workbookStream, cancellationToken).ConfigureAwait(false);
            }
            return result.Report;
        }

        private static IReadOnlyList<PdfExcelTableImportEntry> ImportTables(
            PdfCore.PdfLogicalDocument document,
            ExcelDocument workbook,
            PdfExcelTableImportOptions options) {
            IReadOnlyList<PdfCore.PdfLogicalTableContinuationGroup> tables = PdfCore.PdfLogicalTableContinuations.Group(
                document,
                options.MaxRows,
                options.MergePageContinuations,
                options.MaximumContinuationSegments,
                options.ContinuationGeometryTolerancePoints);
            if (tables.Count == 0) {
                AddEmptyWorkbookSheet(workbook, options);
                return Array.Empty<PdfExcelTableImportEntry>();
            }

            var results = new List<PdfExcelTableImportEntry>(tables.Count);
            for (int i = 0; i < tables.Count; i++) {
                PdfCore.PdfLogicalTableContinuationGroup group = tables[i];
                PdfCore.PdfLogicalTableExtraction extraction = group.Primary;
                string requestedTableName = BuildTableName(options.TableNamePrefix, extraction, i);
                (DataTable dataTable, IReadOnlyList<PdfExcelTableColumnKind> columnKinds) = ToDataTable(
                    requestedTableName,
                    group.Columns,
                    group.Rows,
                    options);
                ExcelSheet sheet = workbook.AddWorksheet(BuildSheetName(options.SheetNamePrefix, extraction, i), SheetNameValidationMode.Sanitize);
                string range = sheet.InsertDataTableAsTable(
                    dataTable,
                    tableName: requestedTableName,
                    style: options.TableStyle,
                    includeAutoFilter: options.IncludeAutoFilter);
                ApplyTypedColumnFormats(sheet, dataTable, columnKinds);

                if (options.AutoFitColumns) {
                    sheet.AutoFitColumns();
                }

                string actualTableName = FindActualTableName(workbook, sheet.Name, range, requestedTableName);
                results.Add(new PdfExcelTableImportEntry(
                    extraction.PageIndex,
                    extraction.PageNumber,
                    extraction.TableIndex,
                    extraction.DetectionKind,
                    sheet.Name,
                    actualTableName,
                    range,
                    group.Columns.Count,
                    group.Rows.Count,
                    group.TotalRowCount,
                    group.Truncated,
                    group.Segments.Select(static segment => segment.PageNumber).ToArray(),
                    group.Segments.Count,
                    group.SuppressedRepeatedHeaderRows,
                    group.AdditionalHeaderRowCount,
                    columnKinds));
            }

            return results.AsReadOnly();
        }

        private static void ApplyTypedColumnFormats(
            ExcelSheet sheet,
            DataTable table,
            IReadOnlyList<PdfExcelTableColumnKind> columnKinds) {
            for (int columnIndex = 0; columnIndex < columnKinds.Count; columnIndex++) {
                if (columnKinds[columnIndex] == PdfExcelTableColumnKind.Percentage) {
                    sheet.ColumnStyleByHeader(table.Columns[columnIndex].ColumnName).Percent(decimals: 2);
                }
            }
        }

        private static void AddEmptyWorkbookSheet(ExcelDocument workbook, PdfExcelTableImportOptions options) {
            ExcelSheet sheet = workbook.AddWorksheet(options.EmptyWorkbookSheetName, SheetNameValidationMode.Sanitize);
            sheet.CellValue(1, 1, "No PDF tables detected.");
        }

        private static (DataTable Table, IReadOnlyList<PdfExcelTableColumnKind> ColumnKinds) ToDataTable(
            string tableName,
            IReadOnlyList<string> columns,
            IReadOnlyList<IReadOnlyList<string>> rows,
            PdfExcelTableImportOptions options) {
            var table = new DataTable(tableName) {
                Locale = CultureInfo.InvariantCulture
            };

            PdfExcelTableColumnKind[] columnKinds = DetectColumnKinds(columns, rows, options);
            var usedColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < columns.Count; i++) {
                table.Columns.Add(GetUniqueColumnName(columns[i], i, usedColumns), GetColumnType(columnKinds[i]));
            }

            table.BeginLoadData();
            try {
                for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
                    DataRow row = table.NewRow();
                    IReadOnlyList<string> sourceRow = rows[rowIndex];
                    for (int columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++) {
                        string value = columnIndex < sourceRow.Count ? sourceRow[columnIndex] : string.Empty;
                        row[columnIndex] = ConvertValue(value, columnKinds[columnIndex], options.NumericCulture);
                    }

                    table.Rows.Add(row);
                }
            } finally {
                table.EndLoadData();
            }

            return (table, Array.AsReadOnly(columnKinds));
        }

        private static PdfExcelTableColumnKind[] DetectColumnKinds(
            IReadOnlyList<string> columns,
            IReadOnlyList<IReadOnlyList<string>> rows,
            PdfExcelTableImportOptions options) {
            var kinds = new PdfExcelTableColumnKind[columns.Count];
            for (int columnIndex = 0; columnIndex < columns.Count; columnIndex++) {
                List<string> values = rows
                    .Select(row => columnIndex < row.Count ? row[columnIndex].Trim() : string.Empty)
                    .Where(static value => value.Length > 0)
                    .ToList();
                if (values.Count == 0) continue;
                if (options.ConvertBooleanColumns && values.All(static value => TryParseBoolean(value, out _))) {
                    kinds[columnIndex] = PdfExcelTableColumnKind.Boolean;
                } else if (options.ConvertPercentageColumns && values.All(value => TryParsePercentage(value, options.NumericCulture, out _))) {
                    kinds[columnIndex] = PdfExcelTableColumnKind.Percentage;
                } else if (options.ConvertNumericColumns &&
                           values.All(PdfCore.PdfLogicalTableAnalysis.LooksLikeNumericValue) &&
                           values.All(value => PdfCore.PdfLogicalTableAnalysis.TryParseNumericValue(value, options.NumericCulture, out _))) {
                    kinds[columnIndex] = PdfExcelTableColumnKind.Number;
                } else if (options.ConvertDateTimeColumns &&
                           HasDateSignal(columns[columnIndex], values) &&
                           values.All(value => DateTime.TryParse(value, options.NumericCulture, DateTimeStyles.AllowWhiteSpaces, out _))) {
                    kinds[columnIndex] = PdfExcelTableColumnKind.DateTime;
                }
            }

            return kinds;
        }

        private static Type GetColumnType(PdfExcelTableColumnKind kind) => kind switch {
            PdfExcelTableColumnKind.Number => typeof(decimal),
            PdfExcelTableColumnKind.Percentage => typeof(decimal),
            PdfExcelTableColumnKind.Boolean => typeof(bool),
            PdfExcelTableColumnKind.DateTime => typeof(DateTime),
            _ => typeof(string)
        };

        private static object ConvertValue(string value, PdfExcelTableColumnKind kind, CultureInfo culture) {
            if (kind == PdfExcelTableColumnKind.Text) return value;
            if (string.IsNullOrWhiteSpace(value)) return DBNull.Value;
            return kind switch {
                PdfExcelTableColumnKind.Number when PdfCore.PdfLogicalTableAnalysis.TryParseNumericValue(value, culture, out decimal number) => number,
                PdfExcelTableColumnKind.Percentage when TryParsePercentage(value, culture, out decimal percentage) => percentage,
                PdfExcelTableColumnKind.Boolean when TryParseBoolean(value, out bool boolean) => boolean,
                PdfExcelTableColumnKind.DateTime when DateTime.TryParse(value, culture, DateTimeStyles.AllowWhiteSpaces, out DateTime dateTime) => dateTime,
                _ => DBNull.Value
            };
        }

        private static bool TryParseBoolean(string value, out bool result) {
            string normalized = value.Trim();
            if (string.Equals(normalized, "true", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(normalized, "yes", StringComparison.OrdinalIgnoreCase)) {
                result = true;
                return true;
            }
            if (string.Equals(normalized, "false", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(normalized, "no", StringComparison.OrdinalIgnoreCase)) {
                result = false;
                return true;
            }
            result = false;
            return false;
        }

        private static bool TryParsePercentage(string value, CultureInfo culture, out decimal result) {
            string normalized = value.Trim();
            if (!normalized.EndsWith("%", StringComparison.Ordinal)) {
                result = 0m;
                return false;
            }

            if (PdfCore.PdfLogicalTableAnalysis.TryParseNumericValue(normalized.Substring(0, normalized.Length - 1), culture, out decimal number)) {
                result = number / 100m;
                return true;
            }

            result = 0m;
            return false;
        }

        private static bool HasDateSignal(string columnName, IReadOnlyList<string> values) {
            string name = columnName.Trim().ToLowerInvariant();
            string[] hints = { "date", "time", "due", "created", "updated", "modified", "issued", "expiry", "expires", "start", "end" };
            if (hints.Any(name.Contains)) return true;
            return values.All(static value => DateTime.TryParseExact(
                value.Trim(),
                UnambiguousDateTimeFormats,
                CultureInfo.InvariantCulture,
                DateTimeStyles.AllowWhiteSpaces,
                out _));
        }

        private static readonly string[] UnambiguousDateTimeFormats = {
            "yyyy-MM-dd", "yyyy/MM/dd", "yyyy.MM.dd",
            "yyyy-MM-dd HH:mm", "yyyy-MM-dd HH:mm:ss",
            "yyyy/MM/dd HH:mm", "yyyy/MM/dd HH:mm:ss",
            "yyyy-MM-dd'T'HH:mm", "yyyy-MM-dd'T'HH:mm:ss",
            "yyyy-MM-dd'T'HH:mm:ss.FFFFFFFK",
            "dd/MM/yyyy", "MM/dd/yyyy", "dd-MM-yyyy", "MM-dd-yyyy", "dd.MM.yyyy",
            "dd/MM/yyyy HH:mm", "MM/dd/yyyy HH:mm",
            "dd-MM-yyyy HH:mm", "MM-dd-yyyy HH:mm", "dd.MM.yyyy HH:mm"
        };

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
