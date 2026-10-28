using System.Data;
using System.Threading;
using System.Threading.Tasks;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Excel {
    /// <summary>
    /// Converts structured logical PDF tables into Excel worksheets.
    /// </summary>
    public static partial class PdfExcelTableConverterExtensions {


        /// <summary>Imports logical PDF tables from an opened PDF into an editable Excel document plus an explicit table-scope report.</summary>
        internal static PdfExcelTableImportResult ConvertToExcelWithEvidence(
            this PdfCore.PdfDocument document,
            ExcelPdfImportOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
            if (document == null) throw new ArgumentNullException(nameof(document));
            return ReadForExcel(document, options, cancellationToken).ConvertToExcelWithEvidence(options, cancellationToken);
        }









        private static PdfCore.PdfDocumentReadResult ReadForExcel(
            PdfCore.PdfDocument document,
            ExcelPdfImportOptions? options,
            CancellationToken cancellationToken = default) {
            return document.Read(options?.ReadOptions, cancellationToken);
        }







        /// <summary>Imports logical PDF tables into an editable Excel document plus an explicit table-scope report.</summary>
        internal static PdfExcelTableImportResult ConvertToExcelWithEvidence(
            this PdfCore.PdfDocumentReadResult document,
            ExcelPdfImportOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
            if (document == null) throw new ArgumentNullException(nameof(document));
            ExcelPdfImportOptions operation = (options ?? new ExcelPdfImportOptions()).Clone();
            operation.CancellationToken = cancellationToken;
            PdfCore.PdfLogicalContentProjection.RequireComplete(document, cancellationToken);
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







        private static IReadOnlyList<PdfExcelTableImportEntry> ImportTables(
            PdfCore.PdfDocumentReadResult document,
            ExcelDocument workbook,
            ExcelPdfImportOptions options) {
            options.CancellationToken.ThrowIfCancellationRequested();
            IReadOnlyList<PdfCore.PdfLogicalTableContinuationGroup> tables = PdfCore.PdfLogicalTableContinuations.Group(
                document,
                options.UseFirstRowAsHeader && options.MaxRows > 0 && options.MaxRows < int.MaxValue ? options.MaxRows + 1 : options.MaxRows,
                options.MergePageContinuations,
                options.SuppressRepeatedBodyHeaderRows,
                options.MaximumContinuationSegments,
                options.ContinuationGeometryTolerancePoints,
                options.CancellationToken);
            if (tables.Count == 0) {
                AddEmptyWorkbookSheet(workbook, options);
                options.CancellationToken.ThrowIfCancellationRequested();
                return Array.Empty<PdfExcelTableImportEntry>();
            }

            var results = new List<PdfExcelTableImportEntry>(tables.Count);
            for (int i = 0; i < tables.Count; i++) {
                options.CancellationToken.ThrowIfCancellationRequested();
                PdfCore.PdfLogicalTableContinuationGroup group = tables[i];
                PdfCore.PdfLogicalTableExtraction extraction = group.Primary;
                bool confirmedHeader = options.UseFirstRowAsHeader && !group.Data.Structure.HasHeaderRow && group.Rows.Count > 0;
                IReadOnlyList<string> columns = confirmedHeader ? group.Rows[0] : group.Columns;
                IReadOnlyList<IReadOnlyList<string>> rows = confirmedHeader ? group.Rows.Skip(1).ToArray() : group.Rows;
                int totalRows = group.TotalRowCount - (confirmedHeader ? 1 : 0);
                if (options.UseFirstRowAsHeader && options.MaxRows > 0) rows = rows.Take(options.MaxRows).ToArray();
                string requestedTableName = BuildTableName(options.TableNamePrefix, extraction, i);
                (DataTable dataTable, IReadOnlyList<ExcelPdfTableColumnKind> columnKinds, IReadOnlyList<string?> currencyTokens,
                    IReadOnlyList<PdfCore.PdfLogicalCurrencyAffixPosition?> currencyAffixPositions,
                    IReadOnlyList<bool?> currencyAffixUsesSpacing) = ToDataTable(
                    requestedTableName,
                    columns,
                    rows,
                    options);
                ExcelSheet sheet = workbook.AddSheet(BuildSheetName(options.SheetNamePrefix, extraction, i), ExcelSheetNameValidationMode.Sanitize, options.CancellationToken);
                string range = sheet.InsertDataTableAsTable(
                    dataTable,
                    tableName: requestedTableName,
                    style: options.TableStyle,
                    includeAutoFilter: options.IncludeAutoFilter,
                    ct: options.CancellationToken);
                ApplyTypedColumnFormats(
                    sheet,
                    dataTable,
                    rows,
                    columnKinds,
                    currencyTokens,
                    currencyAffixPositions,
                    currencyAffixUsesSpacing,
                    options.NumericCulture,
                    options.CancellationToken);

                if (options.AutoFitColumns) {
                    sheet.AutoFitColumns(ct: options.CancellationToken);
                }
                bool columnProportionsApplied = ApplyColumnProportions(sheet, extraction.Table.Columns, columns.Count, options);

                string actualTableName = FindActualTableName(workbook, sheet.Name, range, requestedTableName);
                results.Add(new PdfExcelTableImportEntry(
                    extraction.PageIndex,
                    extraction.PageNumber,
                    extraction.TableIndex,
                    extraction.DetectionKind,
                    sheet.Name,
                    actualTableName,
                    range,
                    columns.Count,
                    rows.Count,
                    totalRows,
                    rows.Count < totalRows,
                    group.Segments.Select(static segment => segment.PageNumber).ToArray(),
                    group.Segments.Count,
                    group.SuppressedRepeatedHeaderRows,
                    group.AdditionalHeaderRowCount,
                    columnKinds,
                    currencyTokens,
                    currencyAffixPositions,
                    currencyAffixUsesSpacing,
                    confirmedHeader,
                    columnProportionsApplied));
            }

            options.CancellationToken.ThrowIfCancellationRequested();
            return results.AsReadOnly();
        }

        private static void ApplyTypedColumnFormats(
            ExcelSheet sheet,
            DataTable table,
            IReadOnlyList<IReadOnlyList<string>> sourceRows,
            IReadOnlyList<ExcelPdfTableColumnKind> columnKinds,
            IReadOnlyList<string?> currencyTokens,
            IReadOnlyList<PdfCore.PdfLogicalCurrencyAffixPosition?> currencyAffixPositions,
            IReadOnlyList<bool?> currencyAffixUsesSpacing,
            CultureInfo numericCulture,
            CancellationToken cancellationToken) {
            for (int columnIndex = 0; columnIndex < columnKinds.Count; columnIndex++) {
                cancellationToken.ThrowIfCancellationRequested();
                if (columnKinds[columnIndex] == ExcelPdfTableColumnKind.Percentage) {
                    sheet.ColumnStyleByHeader(table.Columns[columnIndex].ColumnName).Percent(decimals: 2);
                } else if (columnKinds[columnIndex] == ExcelPdfTableColumnKind.Time) {
                    sheet.ColumnStyleByHeader(table.Columns[columnIndex].ColumnName).Time();
                } else if (columnKinds[columnIndex] == ExcelPdfTableColumnKind.Currency &&
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

        internal static string BuildCurrencyNumberFormat(
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
            string? detectedNegativeSign = FindEquivalentSign(normalized, negativeSign, negative: true);
            bool usesNegativeSign = detectedNegativeSign is not null;
            if (parsedValue > 0M || parsedValue == 0M && !usesAccountingNotation && !usesNegativeSign) {
                string positiveSign = numericCulture.NumberFormat.PositiveSign;
                string? detectedPositiveSign = FindEquivalentSign(normalized, positiveSign, negative: false);
                return detectedPositiveSign is not null
                    ? BuildSignedCurrencyPattern(
                        normalized,
                        detectedPositiveSign,
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
                    detectedNegativeSign ?? negativeSign,
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

        private static string? FindEquivalentSign(string value, string cultureSign, bool negative) {
            if (!string.IsNullOrEmpty(cultureSign) && value.IndexOf(cultureSign, StringComparison.Ordinal) >= 0) {
                return cultureSign;
            }

            string equivalents = negative
                ? "-\u2212\uFE63\uFF0D"
                : "+\uFE62\uFF0B";
            for (int index = 0; index < equivalents.Length; index++) {
                if (value.IndexOf(equivalents[index]) >= 0) return equivalents[index].ToString();
            }

            return null;
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

        private static void AddEmptyWorkbookSheet(ExcelDocument workbook, ExcelPdfImportOptions options) {
            ExcelSheet sheet = workbook.AddSheet(options.EmptyWorkbookSheetName, ExcelSheetNameValidationMode.Sanitize, options.CancellationToken);
            sheet.CellValue(1, 1, "No PDF tables detected.");
        }

        private static (
            DataTable Table,
            IReadOnlyList<ExcelPdfTableColumnKind> ColumnKinds,
            IReadOnlyList<string?> CurrencyTokens,
            IReadOnlyList<PdfCore.PdfLogicalCurrencyAffixPosition?> CurrencyAffixPositions,
            IReadOnlyList<bool?> CurrencyAffixUsesSpacing) ToDataTable(
            string tableName,
            IReadOnlyList<string> columns,
            IReadOnlyList<IReadOnlyList<string>> rows,
            ExcelPdfImportOptions options) {
            var table = new DataTable(tableName) {
                Locale = CultureInfo.InvariantCulture
            };

            (ExcelPdfTableColumnKind[] columnKinds, string?[] currencyTokens,
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
            ExcelPdfTableColumnKind[] Kinds,
            string?[] CurrencyTokens,
            PdfCore.PdfLogicalCurrencyAffixPosition?[] CurrencyAffixPositions,
            bool?[] CurrencyAffixUsesSpacing) DetectColumnKinds(
            IReadOnlyList<string> columns,
            IReadOnlyList<IReadOnlyList<string>> rows,
            ExcelPdfImportOptions options) {
            IReadOnlyList<PdfCore.PdfLogicalTableValueProfile> profiles =
                PdfCore.PdfLogicalTableValueAnalysis.Analyze(
                    columns,
                    rows,
                    new PdfCore.PdfLogicalTableValueAnalysisOptions {
                        NumericCulture = options.NumericCulture,
                        DateTimeCulture = options.DateTimeCulture
                    });
            var kinds = new ExcelPdfTableColumnKind[profiles.Count];
            var currencyTokens = new string?[profiles.Count];
            var currencyAffixPositions = new PdfCore.PdfLogicalCurrencyAffixPosition?[profiles.Count];
            var currencyAffixUsesSpacing = new bool?[profiles.Count];
            for (int columnIndex = 0; columnIndex < profiles.Count; columnIndex++) {
                kinds[columnIndex] = profiles[columnIndex].Kind switch {
                    PdfCore.PdfLogicalTableValueKind.Boolean when options.ConvertBooleanColumns => ExcelPdfTableColumnKind.Boolean,
                    PdfCore.PdfLogicalTableValueKind.Percentage when options.ConvertPercentageColumns => ExcelPdfTableColumnKind.Percentage,
                    PdfCore.PdfLogicalTableValueKind.Time when options.ConvertDateTimeColumns => ExcelPdfTableColumnKind.Time,
                    PdfCore.PdfLogicalTableValueKind.Number when options.ConvertNumericColumns => ExcelPdfTableColumnKind.Number,
                    PdfCore.PdfLogicalTableValueKind.Currency when options.ConvertNumericColumns => ExcelPdfTableColumnKind.Currency,
                    PdfCore.PdfLogicalTableValueKind.DateTime when options.ConvertDateTimeColumns => ExcelPdfTableColumnKind.DateTime,
                    _ => ExcelPdfTableColumnKind.Text
                };
                currencyTokens[columnIndex] = kinds[columnIndex] == ExcelPdfTableColumnKind.Currency
                    ? profiles[columnIndex].CurrencyToken
                    : null;
                currencyAffixPositions[columnIndex] = kinds[columnIndex] == ExcelPdfTableColumnKind.Currency
                    ? profiles[columnIndex].CurrencyAffixPosition
                    : null;
                currencyAffixUsesSpacing[columnIndex] = kinds[columnIndex] == ExcelPdfTableColumnKind.Currency
                    ? profiles[columnIndex].CurrencyAffixUsesSpacing
                    : null;
            }

            return (kinds, currencyTokens, currencyAffixPositions, currencyAffixUsesSpacing);
        }

        private static void AddTypedColumn(DataTable table, string columnName, ExcelPdfTableColumnKind kind) {
            switch (kind) {
                case ExcelPdfTableColumnKind.Number:
                case ExcelPdfTableColumnKind.Percentage:
                case ExcelPdfTableColumnKind.Currency:
                    table.Columns.Add(columnName, typeof(decimal));
                    break;
                case ExcelPdfTableColumnKind.Boolean:
                    table.Columns.Add(columnName, typeof(bool));
                    break;
                case ExcelPdfTableColumnKind.Time:
                    table.Columns.Add(columnName, typeof(TimeSpan));
                    break;
                case ExcelPdfTableColumnKind.DateTime:
                    table.Columns.Add(columnName, typeof(DateTime));
                    break;
                default:
                    table.Columns.Add(columnName, typeof(string));
                    break;
            }
        }

        private static object ConvertValue(
            string value,
            ExcelPdfTableColumnKind kind,
            ExcelPdfImportOptions options) {
            if (kind == ExcelPdfTableColumnKind.Text) return value;
            if (string.IsNullOrWhiteSpace(value)) return DBNull.Value;
            return kind switch {
                ExcelPdfTableColumnKind.Number when PdfCore.PdfLogicalTableAnalysis.TryParseNumericValue(value, options.NumericCulture, out decimal number) => number,
                ExcelPdfTableColumnKind.Currency when PdfCore.PdfLogicalTableValueParser.TryParseCurrency(value, options.NumericCulture, out decimal currency, out _) => currency,
                ExcelPdfTableColumnKind.Percentage when PdfCore.PdfLogicalTableValueParser.TryParsePercentage(value, options.NumericCulture, out decimal percentage) => percentage,
                ExcelPdfTableColumnKind.Boolean when PdfCore.PdfLogicalTableValueParser.TryParseBoolean(value, out bool boolean) => boolean,
                ExcelPdfTableColumnKind.Time when PdfCore.PdfLogicalTableValueParser.TryParseTime(value, options.DateTimeCulture, out TimeSpan time) => time,
                ExcelPdfTableColumnKind.DateTime when PdfCore.PdfLogicalTableValueParser.TryParseDateTime(value, options.DateTimeCulture, out DateTime dateTime) => dateTime,
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
