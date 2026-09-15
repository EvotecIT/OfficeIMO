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
            PdfTablesToExcelOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
            if (document == null) throw new ArgumentNullException(nameof(document));
            return ReadForExcel(document, options, cancellationToken).ImportTablesToExcelDocument(options, cancellationToken);
        }

        /// <summary>Imports logical PDF tables from an opened PDF into an editable Excel document plus an explicit table-scope report.</summary>
        public static PdfExcelTableImportResult ImportTablesToExcelDocumentResult(
            this PdfCore.PdfDocument document,
            PdfTablesToExcelOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
            if (document == null) throw new ArgumentNullException(nameof(document));
            return ReadForExcel(document, options, cancellationToken).ImportTablesToExcelDocumentResult(options, cancellationToken);
        }

        /// <summary>Imports logical PDF tables from an opened PDF into a new Excel workbook.</summary>
        public static OfficeOutputResult<PdfExcelTableImportReport> SaveTablesAsExcel(
            this PdfCore.PdfDocument document,
            string workbookPath,
            PdfTablesToExcelOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
            if (document == null) throw new ArgumentNullException(nameof(document));
            return ReadForExcel(document, options, cancellationToken).SaveTablesAsExcel(workbookPath, options, cancellationToken);
        }

        /// <summary>Imports logical PDF tables from an opened PDF into a caller-owned workbook stream.</summary>
        public static OfficeOutputResult<PdfExcelTableImportReport> SaveTablesAsExcel(
            this PdfCore.PdfDocument document,
            Stream workbookStream,
            PdfTablesToExcelOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
            if (document == null) throw new ArgumentNullException(nameof(document));
            return ReadForExcel(document, options, cancellationToken).SaveTablesAsExcel(workbookStream, options, cancellationToken);
        }

        /// <summary>Imports logical PDF tables from an opened PDF and asynchronously saves a new Excel workbook.</summary>
        public static async Task<OfficeOutputResult<PdfExcelTableImportReport>> SaveTablesAsExcelAsync(
            this PdfCore.PdfDocument document,
            string workbookPath,
            PdfTablesToExcelOptions? options = null,
            CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
            if (document == null) throw new ArgumentNullException(nameof(document));
            PdfTablesToExcelOptions operation = (options ?? new PdfTablesToExcelOptions()).CloneForConversion();
            CancellationToken effectiveCancellationToken = cancellationToken;
            operation.CancellationToken = effectiveCancellationToken;
            return await ReadForExcel(document, operation, effectiveCancellationToken)
                .SaveTablesAsExcelAsync(workbookPath, operation, effectiveCancellationToken)
                .ConfigureAwait(false);
        }

        /// <summary>Imports logical PDF tables from an opened PDF and asynchronously saves to a caller-owned workbook stream.</summary>
        public static async Task<OfficeOutputResult<PdfExcelTableImportReport>> SaveTablesAsExcelAsync(
            this PdfCore.PdfDocument document,
            Stream workbookStream,
            PdfTablesToExcelOptions? options = null,
            CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
            if (document == null) throw new ArgumentNullException(nameof(document));
            PdfTablesToExcelOptions operation = (options ?? new PdfTablesToExcelOptions()).CloneForConversion();
            CancellationToken effectiveCancellationToken = cancellationToken;
            operation.CancellationToken = effectiveCancellationToken;
            return await ReadForExcel(document, operation, effectiveCancellationToken)
                .SaveTablesAsExcelAsync(workbookStream, operation, effectiveCancellationToken)
                .ConfigureAwait(false);
        }

        private static PdfCore.PdfDocumentReadResult ReadForExcel(
            PdfCore.PdfDocument document,
            PdfTablesToExcelOptions? options,
            CancellationToken cancellationToken = default) {
            return document.Read(options?.ReadOptions, cancellationToken);
        }

        /// <summary>Imports logical PDF tables into a new Excel workbook at <paramref name="workbookPath"/>.</summary>
        public static OfficeOutputResult<PdfExcelTableImportReport> SaveTablesAsExcel(
            this PdfCore.PdfDocumentReadResult document,
            string workbookPath,
            PdfTablesToExcelOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (string.IsNullOrWhiteSpace(workbookPath)) throw new ArgumentException("Workbook path cannot be empty.", nameof(workbookPath));

            PdfExcelTableImportResult result = document.ImportTablesToExcelDocumentResult(options, cancellationToken);
            using (result.Value) {
                cancellationToken.ThrowIfCancellationRequested();
                result.Value.Save(workbookPath);
            }
            return OfficeOutputResult<PdfExcelTableImportReport>.FromSuccess(workbookPath, result.Report);
        }

        /// <summary>Imports logical PDF tables into an Excel workbook written to a caller-owned stream.</summary>
        public static OfficeOutputResult<PdfExcelTableImportReport> SaveTablesAsExcel(
            this PdfCore.PdfDocumentReadResult document,
            Stream workbookStream,
            PdfTablesToExcelOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (workbookStream == null) throw new ArgumentNullException(nameof(workbookStream));
            if (!workbookStream.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(workbookStream));

            PdfExcelTableImportResult result = document.ImportTablesToExcelDocumentResult(options, cancellationToken);
            using (result.Value) {
                cancellationToken.ThrowIfCancellationRequested();
                result.Value.Save(workbookStream);
            }
            return OfficeOutputResult<PdfExcelTableImportReport>.FromSuccess(null, result.Report);
        }

        /// <summary>Imports logical PDF tables into a new editable Excel document.</summary>
        public static ExcelDocument ImportTablesToExcelDocument(
            this PdfCore.PdfDocumentReadResult document,
            PdfTablesToExcelOptions? options = null, System.Threading.CancellationToken cancellationToken = default) => document.ImportTablesToExcelDocumentResult(options, cancellationToken).Value;

        /// <summary>Imports logical PDF tables into an editable Excel document plus an explicit table-scope report.</summary>
        public static PdfExcelTableImportResult ImportTablesToExcelDocumentResult(
            this PdfCore.PdfDocumentReadResult document,
            PdfTablesToExcelOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
            if (document == null) throw new ArgumentNullException(nameof(document));
            PdfTablesToExcelOptions operation = (options ?? new PdfTablesToExcelOptions()).CloneForConversion();
            operation.CancellationToken = cancellationToken;
            ExcelDocument workbook = ExcelDocument.Create();
            try {
                IReadOnlyList<PdfExcelTableImportEntry> entries = ImportTables(document, workbook, operation);
                PdfCore.PdfTableExtractionScopeReport sourceScope = PdfCore.PdfLogicalTableAnalysis.AnalyzeExtractionScope(document);
                return new PdfExcelTableImportResult(workbook, new PdfExcelTableImportReport(entries, sourceScope));
            } catch {
                workbook.Dispose();
                throw;
            }
        }

        /// <summary>Asynchronously imports logical PDF tables into an Excel workbook written to a file.</summary>
        public static async Task<OfficeOutputResult<PdfExcelTableImportReport>> SaveTablesAsExcelAsync(
            this PdfCore.PdfDocumentReadResult document,
            string workbookPath,
            PdfTablesToExcelOptions? options = null,
            CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (string.IsNullOrWhiteSpace(workbookPath)) throw new ArgumentException("Workbook path cannot be empty.", nameof(workbookPath));
            PdfTablesToExcelOptions operation = (options ?? new PdfTablesToExcelOptions()).CloneForConversion();
            CancellationToken effectiveCancellationToken = cancellationToken;
            operation.CancellationToken = effectiveCancellationToken;
            operation.CancellationToken.ThrowIfCancellationRequested();
            PdfExcelTableImportResult result = document.ImportTablesToExcelDocumentResult(operation, cancellationToken);
            using (result.Value) {
                await result.Value.SaveAsync(workbookPath, cancellationToken: operation.CancellationToken).ConfigureAwait(false);
            }
            return OfficeOutputResult<PdfExcelTableImportReport>.FromSuccess(workbookPath, result.Report);
        }

        /// <summary>Asynchronously imports logical PDF tables into an Excel workbook written to a caller-owned stream.</summary>
        public static async Task<OfficeOutputResult<PdfExcelTableImportReport>> SaveTablesAsExcelAsync(
            this PdfCore.PdfDocumentReadResult document,
            Stream workbookStream,
            PdfTablesToExcelOptions? options = null,
            CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (workbookStream == null) throw new ArgumentNullException(nameof(workbookStream));
            if (!workbookStream.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(workbookStream));
            PdfTablesToExcelOptions operation = (options ?? new PdfTablesToExcelOptions()).CloneForConversion();
            CancellationToken effectiveCancellationToken = cancellationToken;
            operation.CancellationToken = effectiveCancellationToken;
            operation.CancellationToken.ThrowIfCancellationRequested();
            PdfExcelTableImportResult result = document.ImportTablesToExcelDocumentResult(operation, cancellationToken);
            using (result.Value) {
                await result.Value.SaveAsync(workbookStream, operation.CancellationToken).ConfigureAwait(false);
            }
            return OfficeOutputResult<PdfExcelTableImportReport>.FromSuccess(null, result.Report);
        }



        private static IReadOnlyList<PdfExcelTableImportEntry> ImportTables(
            PdfCore.PdfDocumentReadResult document,
            ExcelDocument workbook,
            PdfTablesToExcelOptions options) {
            options.CancellationToken.ThrowIfCancellationRequested();
            IReadOnlyList<PdfCore.PdfLogicalTableContinuationGroup> tables = PdfCore.PdfLogicalTableContinuations.Group(
                document,
                options.MaxRows,
                options.MergePageContinuations,
                options.SuppressRepeatedBodyHeaderRows,
                options.MaximumContinuationSegments,
                options.ContinuationGeometryTolerancePoints,
                options.CancellationToken);
            if (tables.Count == 0) {
                AddEmptyWorkbookSheet(workbook, options);
                return Array.Empty<PdfExcelTableImportEntry>();
            }

            var results = new List<PdfExcelTableImportEntry>(tables.Count);
            for (int i = 0; i < tables.Count; i++) {
                options.CancellationToken.ThrowIfCancellationRequested();
                PdfCore.PdfLogicalTableContinuationGroup group = tables[i];
                PdfCore.PdfLogicalTableExtraction extraction = group.Primary;
                string requestedTableName = BuildTableName(options.TableNamePrefix, extraction, i);
                (DataTable dataTable, IReadOnlyList<PdfExcelTableColumnKind> columnKinds, IReadOnlyList<string?> currencyTokens,
                    IReadOnlyList<PdfCore.PdfLogicalCurrencyAffixPosition?> currencyAffixPositions,
                    IReadOnlyList<bool?> currencyAffixUsesSpacing) = ToDataTable(
                    requestedTableName,
                    group.Columns,
                    group.Rows,
                    options);
                ExcelSheet sheet = workbook.AddWorksheet(BuildSheetName(options.SheetNamePrefix, extraction, i), ExcelSheetNameValidationMode.Sanitize);
                string range = sheet.InsertDataTableAsTable(
                    dataTable,
                    tableName: requestedTableName,
                    style: options.TableStyle,
                    includeAutoFilter: options.IncludeAutoFilter);
                ApplyTypedColumnFormats(
                    sheet,
                    dataTable,
                    group.Rows,
                    columnKinds,
                    currencyTokens,
                    currencyAffixPositions,
                    currencyAffixUsesSpacing,
                    options.NumericCulture,
                    options.CancellationToken);

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
                    columnKinds,
                    currencyTokens,
                    currencyAffixPositions,
                    currencyAffixUsesSpacing));
            }

            return results.AsReadOnly();
        }

        private static void ApplyTypedColumnFormats(
            ExcelSheet sheet,
            DataTable table,
            IReadOnlyList<IReadOnlyList<string>> sourceRows,
            IReadOnlyList<PdfExcelTableColumnKind> columnKinds,
            IReadOnlyList<string?> currencyTokens,
            IReadOnlyList<PdfCore.PdfLogicalCurrencyAffixPosition?> currencyAffixPositions,
            IReadOnlyList<bool?> currencyAffixUsesSpacing,
            CultureInfo numericCulture,
            CancellationToken cancellationToken) {
            for (int columnIndex = 0; columnIndex < columnKinds.Count; columnIndex++) {
                cancellationToken.ThrowIfCancellationRequested();
                if (columnKinds[columnIndex] == PdfExcelTableColumnKind.Percentage) {
                    sheet.ColumnStyleByHeader(table.Columns[columnIndex].ColumnName).Percent(decimals: 2);
                } else if (columnKinds[columnIndex] == PdfExcelTableColumnKind.Time) {
                    sheet.ColumnStyleByHeader(table.Columns[columnIndex].ColumnName).Time();
                } else if (columnKinds[columnIndex] == PdfExcelTableColumnKind.Currency &&
                           !string.IsNullOrWhiteSpace(currencyTokens[columnIndex]) &&
                           currencyAffixPositions[columnIndex].HasValue &&
                           currencyAffixUsesSpacing[columnIndex].HasValue) {
                    for (int rowIndex = 0; rowIndex < sourceRows.Count; rowIndex++) {
                        cancellationToken.ThrowIfCancellationRequested();
                        string sourceValue = columnIndex < sourceRows[rowIndex].Count
                            ? sourceRows[rowIndex][columnIndex]
                            : string.Empty;
                        if (!PdfCore.PdfLogicalTableValueParser.TryParseCurrency(
                                sourceValue,
                                numericCulture,
                                out decimal parsedValue,
                                out _,
                                out _,
                                out _,
                                out int decimalPlaces)) {
                            continue;
                        }
                        sheet.CellAt(rowIndex + 2, columnIndex + 1).SetNumberFormat(
                            BuildCurrencyNumberFormat(
                                currencyTokens[columnIndex]!,
                                currencyAffixPositions[columnIndex]!.Value,
                                currencyAffixUsesSpacing[columnIndex]!.Value,
                                decimalPlaces,
                                parsedValue,
                                sourceValue,
                                numericCulture));
                    }
                }
            }
        }

        private static string BuildCurrencyNumberFormat(
            string currencyToken,
            PdfCore.PdfLogicalCurrencyAffixPosition affixPosition,
            bool affixUsesSpacing,
            int decimalPlaces,
            decimal parsedValue,
            string sourceValue,
            CultureInfo numericCulture) {
            string escapedToken = currencyToken.Replace("\"", "\"\"");
            string literal = "\"" + escapedToken + "\"";
            string separator = affixUsesSpacing ? " " : string.Empty;
            string numeric = decimalPlaces <= 0
                ? "#,##0"
                : "#,##0." + new string('0', Math.Min(decimalPlaces, 28));
            string positive = affixPosition == PdfCore.PdfLogicalCurrencyAffixPosition.Prefix
                ? literal + separator + numeric
                : numeric + separator + literal;
            string normalized = sourceValue.Trim();
            string negativeSign = numericCulture.NumberFormat.NegativeSign;
            if (string.IsNullOrEmpty(negativeSign)) negativeSign = "-";
            bool usesAccountingNotation = normalized.Length >= 3 &&
                normalized[0] == '(' && normalized[normalized.Length - 1] == ')';
            bool usesNegativeSign = normalized.IndexOf(negativeSign, StringComparison.Ordinal) >= 0;
            if (parsedValue > 0M || parsedValue == 0M && !usesAccountingNotation && !usesNegativeSign) {
                string positiveSign = numericCulture.NumberFormat.PositiveSign;
                return !string.IsNullOrEmpty(positiveSign) &&
                       normalized.IndexOf(positiveSign, StringComparison.Ordinal) >= 0
                    ? BuildSignedCurrencyPattern(
                        normalized,
                        positiveSign,
                        currencyToken,
                        affixPosition,
                        literal,
                        separator,
                        numeric,
                        positive)
                    : positive;
            }

            string negative = usesAccountingNotation
                ? "\"(\"" + positive + "\")\""
                : BuildSignedCurrencyPattern(
                    normalized,
                    negativeSign,
                    currencyToken,
                    affixPosition,
                    literal,
                    separator,
                    numeric,
                    positive);
            if (parsedValue == 0M) {
                return positive + ";" + negative + ";" + negative;
            }

            return string.Equals(negative, positive, StringComparison.Ordinal)
                ? positive
                : positive + ";" + negative;
        }

        private static string BuildSignedCurrencyPattern(
            string normalized,
            string sign,
            string currencyToken,
            PdfCore.PdfLogicalCurrencyAffixPosition affixPosition,
            string literal,
            string separator,
            string numeric,
            string unsignedPattern) {
            string signLiteral = "\"" + sign.Replace("\"", "\"\"") + "\"";
            if (normalized.StartsWith(sign, StringComparison.Ordinal)) return signLiteral + unsignedPattern;
            if (normalized.EndsWith(sign, StringComparison.Ordinal)) return unsignedPattern + signLiteral;
            if (normalized.IndexOf(sign, StringComparison.Ordinal) < 0 ||
                normalized.IndexOf(currencyToken, StringComparison.Ordinal) < 0) {
                return unsignedPattern;
            }
            return affixPosition == PdfCore.PdfLogicalCurrencyAffixPosition.Prefix
                ? literal + separator + signLiteral + numeric
                : numeric + signLiteral + separator + literal;
        }

        private static void AddEmptyWorkbookSheet(ExcelDocument workbook, PdfTablesToExcelOptions options) {
            ExcelSheet sheet = workbook.AddWorksheet(options.EmptyWorkbookSheetName, ExcelSheetNameValidationMode.Sanitize);
            sheet.CellValue(1, 1, "No PDF tables detected.");
        }

        private static (
            DataTable Table,
            IReadOnlyList<PdfExcelTableColumnKind> ColumnKinds,
            IReadOnlyList<string?> CurrencyTokens,
            IReadOnlyList<PdfCore.PdfLogicalCurrencyAffixPosition?> CurrencyAffixPositions,
            IReadOnlyList<bool?> CurrencyAffixUsesSpacing) ToDataTable(
            string tableName,
            IReadOnlyList<string> columns,
            IReadOnlyList<IReadOnlyList<string>> rows,
            PdfTablesToExcelOptions options) {
            var table = new DataTable(tableName) {
                Locale = CultureInfo.InvariantCulture
            };

            (PdfExcelTableColumnKind[] columnKinds, string?[] currencyTokens,
                PdfCore.PdfLogicalCurrencyAffixPosition?[] currencyAffixPositions,
                bool?[] currencyAffixUsesSpacing) = DetectColumnKinds(columns, rows, options);
            var usedColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < columns.Count; i++) {
                options.CancellationToken.ThrowIfCancellationRequested();
                AddTypedColumn(table, GetUniqueColumnName(columns[i], i, usedColumns), columnKinds[i]);
            }

            table.BeginLoadData();
            try {
                for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
                    options.CancellationToken.ThrowIfCancellationRequested();
                    DataRow row = table.NewRow();
                    IReadOnlyList<string> sourceRow = rows[rowIndex];
                    for (int columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++) {
                        string value = columnIndex < sourceRow.Count ? sourceRow[columnIndex] : string.Empty;
                        row[columnIndex] = ConvertValue(value, columnKinds[columnIndex], options);
                    }

                    table.Rows.Add(row);
                }
            } finally {
                table.EndLoadData();
            }

            return (
                table,
                Array.AsReadOnly(columnKinds),
                Array.AsReadOnly(currencyTokens),
                Array.AsReadOnly(currencyAffixPositions),
                Array.AsReadOnly(currencyAffixUsesSpacing));
        }

        private static (
            PdfExcelTableColumnKind[] Kinds,
            string?[] CurrencyTokens,
            PdfCore.PdfLogicalCurrencyAffixPosition?[] CurrencyAffixPositions,
            bool?[] CurrencyAffixUsesSpacing) DetectColumnKinds(
            IReadOnlyList<string> columns,
            IReadOnlyList<IReadOnlyList<string>> rows,
            PdfTablesToExcelOptions options) {
            IReadOnlyList<PdfCore.PdfLogicalTableValueProfile> profiles =
                PdfCore.PdfLogicalTableValueAnalysis.Analyze(
                    columns,
                    rows,
                    new PdfCore.PdfLogicalTableValueAnalysisOptions {
                        NumericCulture = options.NumericCulture,
                        DateTimeCulture = options.DateTimeCulture
                    });
            var kinds = new PdfExcelTableColumnKind[profiles.Count];
            var currencyTokens = new string?[profiles.Count];
            var currencyAffixPositions = new PdfCore.PdfLogicalCurrencyAffixPosition?[profiles.Count];
            var currencyAffixUsesSpacing = new bool?[profiles.Count];
            for (int columnIndex = 0; columnIndex < profiles.Count; columnIndex++) {
                kinds[columnIndex] = profiles[columnIndex].Kind switch {
                    PdfCore.PdfLogicalTableValueKind.Boolean when options.ConvertBooleanColumns => PdfExcelTableColumnKind.Boolean,
                    PdfCore.PdfLogicalTableValueKind.Percentage when options.ConvertPercentageColumns => PdfExcelTableColumnKind.Percentage,
                    PdfCore.PdfLogicalTableValueKind.Time when options.ConvertDateTimeColumns => PdfExcelTableColumnKind.Time,
                    PdfCore.PdfLogicalTableValueKind.Number when options.ConvertNumericColumns => PdfExcelTableColumnKind.Number,
                    PdfCore.PdfLogicalTableValueKind.Currency when options.ConvertNumericColumns => PdfExcelTableColumnKind.Currency,
                    PdfCore.PdfLogicalTableValueKind.DateTime when options.ConvertDateTimeColumns => PdfExcelTableColumnKind.DateTime,
                    _ => PdfExcelTableColumnKind.Text
                };
                currencyTokens[columnIndex] = kinds[columnIndex] == PdfExcelTableColumnKind.Currency
                    ? profiles[columnIndex].CurrencyToken
                    : null;
                currencyAffixPositions[columnIndex] = kinds[columnIndex] == PdfExcelTableColumnKind.Currency
                    ? profiles[columnIndex].CurrencyAffixPosition
                    : null;
                currencyAffixUsesSpacing[columnIndex] = kinds[columnIndex] == PdfExcelTableColumnKind.Currency
                    ? profiles[columnIndex].CurrencyAffixUsesSpacing
                    : null;
            }

            return (kinds, currencyTokens, currencyAffixPositions, currencyAffixUsesSpacing);
        }

        private static void AddTypedColumn(DataTable table, string columnName, PdfExcelTableColumnKind kind) {
            switch (kind) {
                case PdfExcelTableColumnKind.Number:
                case PdfExcelTableColumnKind.Percentage:
                case PdfExcelTableColumnKind.Currency:
                    table.Columns.Add(columnName, typeof(decimal));
                    break;
                case PdfExcelTableColumnKind.Boolean:
                    table.Columns.Add(columnName, typeof(bool));
                    break;
                case PdfExcelTableColumnKind.Time:
                    table.Columns.Add(columnName, typeof(TimeSpan));
                    break;
                case PdfExcelTableColumnKind.DateTime:
                    table.Columns.Add(columnName, typeof(DateTime));
                    break;
                default:
                    table.Columns.Add(columnName, typeof(string));
                    break;
            }
        }

        private static object ConvertValue(
            string value,
            PdfExcelTableColumnKind kind,
            PdfTablesToExcelOptions options) {
            if (kind == PdfExcelTableColumnKind.Text) return value;
            if (string.IsNullOrWhiteSpace(value)) return DBNull.Value;
            return kind switch {
                PdfExcelTableColumnKind.Number when PdfCore.PdfLogicalTableAnalysis.TryParseNumericValue(value, options.NumericCulture, out decimal number) => number,
                PdfExcelTableColumnKind.Currency when PdfCore.PdfLogicalTableValueParser.TryParseCurrency(value, options.NumericCulture, out decimal currency, out _) => currency,
                PdfExcelTableColumnKind.Percentage when PdfCore.PdfLogicalTableValueParser.TryParsePercentage(value, options.NumericCulture, out decimal percentage) => percentage,
                PdfExcelTableColumnKind.Boolean when PdfCore.PdfLogicalTableValueParser.TryParseBoolean(value, out bool boolean) => boolean,
                PdfExcelTableColumnKind.Time when PdfCore.PdfLogicalTableValueParser.TryParseTime(value, options.DateTimeCulture, out TimeSpan time) => time,
                PdfExcelTableColumnKind.DateTime when PdfCore.PdfLogicalTableValueParser.TryParseDateTime(value, options.DateTimeCulture, out DateTime dateTime) => dateTime,
                _ => DBNull.Value
            };
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
