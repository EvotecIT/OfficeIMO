using System.Data;
using System.Globalization;
using System.ComponentModel;
using System.Text;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private sealed class DirectDataSetWorkbookModel {
            private DirectDataSetWorkbookModel(
                IReadOnlyList<DirectDataSetSheetModel> sheets,
                IReadOnlyList<ExcelDataSetImportResult> results,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                ExcelDateSystem dateSystem = ExcelDateSystem.NineteenHundred) {
                Sheets = sheets;
                Results = results;
                DateTimeOffsetWriteStrategy = dateTimeOffsetWriteStrategy;
                DateSystem = dateSystem;
            }

            internal IReadOnlyList<DirectDataSetSheetModel> Sheets { get; }

            internal IReadOnlyList<ExcelDataSetImportResult> Results { get; }

            internal Func<DateTimeOffset, DateTime> DateTimeOffsetWriteStrategy { get; }

            internal ExcelDateSystem DateSystem { get; }

            internal DirectDataSetWorkbookModel WithWorksheetMetadata(IReadOnlyList<DirectWorksheetMetadata?> metadata) {
                if (metadata == null) throw new ArgumentNullException(nameof(metadata));
                if (metadata.Count != Sheets.Count) {
                    throw new ArgumentException("Metadata count must match the sheet count.", nameof(metadata));
                }

                var sheets = new DirectDataSetSheetModel[Sheets.Count];
                for (int i = 0; i < Sheets.Count; i++) {
                    sheets[i] = Sheets[i].WithMetadata(metadata[i]);
                }

                return new DirectDataSetWorkbookModel(sheets, Results, DateTimeOffsetWriteStrategy, DateSystem);
            }

            internal DirectDataSetWorkbookModel WithAutoFitColumns(
                string sheetName,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                CancellationToken ct) {
                var sheets = new DirectDataSetSheetModel[Sheets.Count];
                bool canCancel = ct.CanBeCanceled;
                for (int i = 0; i < Sheets.Count; i++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    var sheet = Sheets[i];
                    if (!string.Equals(sheet.SheetName, sheetName, StringComparison.Ordinal)) {
                        sheets[i] = sheet;
                        continue;
                    }

                    double[] columnWidths = sheet.Table.CalculateColumnWidths(sheet.IncludeHeaders, dateTimeOffsetWriteStrategy, ct);
                    sheets[i] = new DirectDataSetSheetModel(
                        sheet.Index,
                        sheet.SheetName,
                        sheet.TableName,
                        sheet.Range,
                        sheet.Table,
                        sheet.TableStyle,
                        sheet.IncludeHeaders,
                        sheet.IncludeAutoFilter,
                        sheet.HasTable,
                        autoFitColumns: true,
                        sheet.OmitBlankCells,
                        columnWidths,
                        sheet.UseCellValueNumberFormats,
                        sheet.Metadata,
                        sheet.ColumnNumberFormats,
                        sheet.ShowFirstColumn,
                        sheet.ShowLastColumn,
                        sheet.ShowRowStripes,
                        sheet.ShowColumnStripes);
                }

                return new DirectDataSetWorkbookModel(sheets, Results, dateTimeOffsetWriteStrategy ?? DateTimeOffsetWriteStrategy, DateSystem);
            }

            internal DirectDataSetWorkbookModel WithTableAutoFilter(string sheetName, bool includeAutoFilter) {
                var sheets = new DirectDataSetSheetModel[Sheets.Count];
                for (int i = 0; i < Sheets.Count; i++) {
                    var sheet = Sheets[i];
                    if (!string.Equals(sheet.SheetName, sheetName, StringComparison.Ordinal)) {
                        sheets[i] = sheet;
                        continue;
                    }

                    sheets[i] = new DirectDataSetSheetModel(
                        sheet.Index,
                        sheet.SheetName,
                        sheet.TableName,
                        sheet.Range,
                        sheet.Table,
                        sheet.TableStyle,
                        sheet.IncludeHeaders,
                        includeAutoFilter,
                        sheet.HasTable,
                        sheet.AutoFitColumns,
                        sheet.OmitBlankCells,
                        sheet.ColumnWidths,
                        sheet.UseCellValueNumberFormats,
                        sheet.Metadata,
                        sheet.ColumnNumberFormats,
                        sheet.ShowFirstColumn,
                        sheet.ShowLastColumn,
                        sheet.ShowRowStripes,
                        sheet.ShowColumnStripes);
                }

                return new DirectDataSetWorkbookModel(sheets, Results, DateTimeOffsetWriteStrategy, DateSystem);
            }

            internal bool TryWithTableStyle(
                string sheetName,
                string tableOrRange,
                TableStyle tableStyle,
                bool? showFirstColumn,
                bool? showLastColumn,
                bool? showRowStripes,
                bool? showColumnStripes,
                out DirectDataSetWorkbookModel model) {
                var sheets = new DirectDataSetSheetModel[Sheets.Count];
                bool matched = false;
                for (int i = 0; i < Sheets.Count; i++) {
                    var sheet = Sheets[i];
                    if (!matched
                        && sheet.HasTable
                        && string.Equals(sheet.SheetName, sheetName, StringComparison.Ordinal)
                        && (string.Equals(sheet.Range, tableOrRange, StringComparison.OrdinalIgnoreCase)
                            || string.Equals(sheet.TableName, tableOrRange, StringComparison.OrdinalIgnoreCase))) {
                        sheets[i] = sheet.WithTableStyle(tableStyle, showFirstColumn, showLastColumn, showRowStripes, showColumnStripes);
                        matched = true;
                    } else {
                        sheets[i] = sheet;
                    }
                }

                model = matched ? new DirectDataSetWorkbookModel(sheets, Results, DateTimeOffsetWriteStrategy, DateSystem) : this;
                return matched;
            }

            internal DirectDataSetWorkbookModel WithAutoFitColumns(
                string sheetName,
                IReadOnlyList<int> columnIndexes,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                CancellationToken ct) {
                var sheets = new DirectDataSetSheetModel[Sheets.Count];
                bool canCancel = ct.CanBeCanceled;
                for (int i = 0; i < Sheets.Count; i++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    var sheet = Sheets[i];
                    if (!string.Equals(sheet.SheetName, sheetName, StringComparison.Ordinal)) {
                        sheets[i] = sheet;
                        continue;
                    }

                    double[] columnWidths = sheet.Table.CalculateColumnWidths(sheet.IncludeHeaders, dateTimeOffsetWriteStrategy, ct, columnIndexes);
                    if (sheet.ColumnWidths != null && sheet.ColumnWidths.Length == columnWidths.Length) {
                        var mergedWidths = new double[columnWidths.Length];
                        Array.Copy(sheet.ColumnWidths, mergedWidths, mergedWidths.Length);
                        for (int columnIndex = 0; columnIndex < columnIndexes.Count; columnIndex++) {
                            int widthIndex = columnIndexes[columnIndex] - 1;
                            if (widthIndex >= 0 && widthIndex < mergedWidths.Length) {
                                mergedWidths[widthIndex] = columnWidths[widthIndex];
                            }
                        }

                        columnWidths = mergedWidths;
                    }

                    sheets[i] = new DirectDataSetSheetModel(
                        sheet.Index,
                        sheet.SheetName,
                        sheet.TableName,
                        sheet.Range,
                        sheet.Table,
                        sheet.TableStyle,
                        sheet.IncludeHeaders,
                        sheet.IncludeAutoFilter,
                        sheet.HasTable,
                        autoFitColumns: false,
                        sheet.OmitBlankCells,
                        columnWidths,
                        sheet.UseCellValueNumberFormats,
                        sheet.Metadata,
                        sheet.ColumnNumberFormats,
                        sheet.ShowFirstColumn,
                        sheet.ShowLastColumn,
                        sheet.ShowRowStripes,
                        sheet.ShowColumnStripes);
                }

                return new DirectDataSetWorkbookModel(sheets, Results, dateTimeOffsetWriteStrategy ?? DateTimeOffsetWriteStrategy, DateSystem);
            }

            internal DirectDataSetWorkbookModel WithTable(
                string sheetName,
                string tableName,
                bool includeHeaders,
                TableStyle tableStyle,
                bool includeAutoFilter,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                CancellationToken ct) {
                var sheets = new DirectDataSetSheetModel[Sheets.Count];
                var results = new ExcelDataSetImportResult[Sheets.Count];
                bool canCancel = ct.CanBeCanceled;
                for (int i = 0; i < Sheets.Count; i++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    var sheet = Sheets[i];
                    if (!string.Equals(sheet.SheetName, sheetName, StringComparison.Ordinal)) {
                        sheets[i] = sheet;
                        results[i] = new ExcelDataSetImportResult(sheet.SheetName, sheet.TableName, sheet.Range, sheet.Table.RowCount, sheet.Table.ColumnCount);
                        continue;
                    }

                    var table = includeHeaders ? sheet.Table : sheet.Table.WithGeneratedColumnNames();
                    double[]? columnWidths = sheet.ColumnWidths;
                    if (sheet.AutoFitColumns && !ReferenceEquals(table, sheet.Table)) {
                        columnWidths = table.CalculateColumnWidths(includeHeaders, dateTimeOffsetWriteStrategy, ct);
                    }

                    sheets[i] = new DirectDataSetSheetModel(
                        sheet.Index,
                        sheet.SheetName,
                        tableName,
                        sheet.Range,
                        table,
                        tableStyle,
                        includeHeaders,
                        includeAutoFilter,
                        hasTable: true,
                        sheet.AutoFitColumns,
                        sheet.OmitBlankCells,
                        columnWidths,
                        sheet.UseCellValueNumberFormats,
                        sheet.Metadata,
                        sheet.ColumnNumberFormats,
                        sheet.ShowFirstColumn,
                        sheet.ShowLastColumn,
                        sheet.ShowRowStripes,
                        sheet.ShowColumnStripes);
                    results[i] = new ExcelDataSetImportResult(sheet.SheetName, tableName, sheet.Range, table.RowCount, table.ColumnCount);
                }

                return new DirectDataSetWorkbookModel(sheets, results, dateTimeOffsetWriteStrategy ?? DateTimeOffsetWriteStrategy, DateSystem);
            }

            internal DirectDataSetWorkbookModel WithColumnNumberFormat(string sheetName, int columnIndex, string numberFormat) {
                var sheets = new DirectDataSetSheetModel[Sheets.Count];
                for (int i = 0; i < Sheets.Count; i++) {
                    var sheet = Sheets[i];
                    sheets[i] = string.Equals(sheet.SheetName, sheetName, StringComparison.Ordinal)
                        ? sheet.WithColumnNumberFormat(columnIndex, numberFormat)
                        : sheet;
                }

                return new DirectDataSetWorkbookModel(sheets, Results, DateTimeOffsetWriteStrategy, DateSystem);
            }


            internal static DirectDataSetWorkbookModel Create(
                DataSet dataSet,
                bool createTables,
                TableStyle tableStyle,
                bool includeHeaders,
                bool includeAutoFilter,
                bool autoFit,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                CancellationToken ct,
                IReadOnlyList<ExcelDataSetImportResult>? importResults = null,
                bool snapshotTables = false,
                bool omitBlankCells = false,
                ExcelDateSystem dateSystem = ExcelDateSystem.NineteenHundred) {
                var sheets = new List<DirectDataSetSheetModel>();
                var results = new List<ExcelDataSetImportResult>();
                var usedSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var usedTableNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                int index = 1;
                bool canCancel = ct.CanBeCanceled;
                foreach (DataTable table in dataSet.Tables) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    var tableModel = snapshotTables
                        ? DirectDataSetTableModel.Snapshot(table, ct)
                        : DirectDataSetTableModel.Reference(table);
                    string requestedName = string.IsNullOrWhiteSpace(table.TableName)
                        ? "Table" + index.ToString(CultureInfo.InvariantCulture)
                        : table.TableName;
                    ExcelDataSetImportResult? importResult = importResults != null && index <= importResults.Count
                        ? importResults[index - 1]
                        : null;
                    string sheetName = importResult?.SheetName ?? GetUniqueSheetName(SanitizeSheetName(requestedName), usedSheetNames);
                    usedSheetNames.Add(sheetName);
                    string tableName = importResult?.TableName ?? GetUniqueName(SanitizeTableName(requestedName), usedTableNames, 255);
                    usedTableNames.Add(tableName);
                    int rowCount = tableModel.RowCount + (includeHeaders ? 1 : 0);
                    ValidateWorksheetBounds(tableModel, rowCount, requestedName);
                    string range = importResult?.Range ?? (tableModel.ColumnCount == 0 || rowCount == 0
                        ? string.Empty
                        : "A1:" + A1.CellReference(rowCount, tableModel.ColumnCount));

                    bool hasTable = createTables && range.Length > 0;
                    double[]? columnWidths = autoFit && tableModel.ColumnCount > 0
                        ? tableModel.CalculateColumnWidths(includeHeaders, dateTimeOffsetWriteStrategy, ct)
                        : null;
                    var sheet = new DirectDataSetSheetModel(index, sheetName, hasTable ? tableName : null, range, tableModel, tableStyle, includeHeaders, includeAutoFilter, hasTable, autoFit, omitBlankCells, columnWidths);
                    sheets.Add(sheet);
                    results.Add(new ExcelDataSetImportResult(sheetName, hasTable ? tableName : null, range, tableModel.RowCount, tableModel.ColumnCount));
                    index++;
                }

                return new DirectDataSetWorkbookModel(sheets, results, dateTimeOffsetWriteStrategy ?? DefaultDateTimeOffsetWriteStrategy, dateSystem);
            }

            internal static DirectDataSetWorkbookModel CreateSingle(
                string sheetName,
                string requestedName,
                string? tableName,
                string range,
                DirectDataSetTableModel tableModel,
                bool createTable,
                TableStyle tableStyle,
                bool includeHeaders,
                bool includeAutoFilter,
                bool autoFit,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                CancellationToken ct,
                bool useCellValueNumberFormats = false,
                ExcelDateSystem dateSystem = ExcelDateSystem.NineteenHundred) {
                int rowCount = tableModel.RowCount + (includeHeaders ? 1 : 0);
                ValidateWorksheetBounds(tableModel, rowCount, requestedName);
                bool hasTable = createTable && range.Length > 0;
                string? resolvedTableName = hasTable
                    ? SanitizeTableName(string.IsNullOrWhiteSpace(tableName) ? requestedName : tableName!)
                    : null;
                double[]? columnWidths = autoFit && tableModel.ColumnCount > 0
                    ? tableModel.CalculateColumnWidths(includeHeaders, dateTimeOffsetWriteStrategy, ct)
                    : null;
                var sheet = new DirectDataSetSheetModel(1, sheetName, resolvedTableName, range, tableModel, tableStyle, includeHeaders, includeAutoFilter, hasTable, autoFit, omitBlankCells: false, columnWidths: columnWidths, useCellValueNumberFormats: useCellValueNumberFormats);
                var result = new ExcelDataSetImportResult(sheetName, resolvedTableName, range, tableModel.RowCount, tableModel.ColumnCount);
                return new DirectDataSetWorkbookModel([sheet], [result], dateTimeOffsetWriteStrategy ?? DefaultDateTimeOffsetWriteStrategy, dateSystem);
            }

            private static string GetUniqueSheetName(string baseName, HashSet<string> used) {
                string trimmed = baseName;
                if (string.IsNullOrWhiteSpace(trimmed)) {
                    int defaultIndex = 1;
                    string defaultCandidate = "Sheet1";
                    while (used.Contains(defaultCandidate)) {
                        defaultIndex++;
                        defaultCandidate = "Sheet" + defaultIndex.ToString(CultureInfo.InvariantCulture);
                    }

                    return defaultCandidate;
                }

                if (trimmed.Length > 31) {
                    trimmed = trimmed.Substring(0, 31);
                }

                if (!used.Contains(trimmed)) {
                    return trimmed;
                }

                int suffix = 2;
                while (true) {
                    string suffixText = " (" + suffix.ToString(CultureInfo.InvariantCulture) + ")";
                    int prefixLength = Math.Max(1, 31 - suffixText.Length);
                    string candidate = trimmed.Length > prefixLength
                        ? trimmed.Substring(0, prefixLength) + suffixText
                        : trimmed + suffixText;
                    if (!used.Contains(candidate)) {
                        return candidate;
                    }

                    suffix++;
                }
            }

            private static void ValidateWorksheetBounds(DirectDataSetTableModel table, int rowCount, string requestedName) {
                if (table.ColumnCount > A1.MaxColumns) {
                    throw new ArgumentException($"DataTable '{requestedName}' has {table.ColumnCount.ToString(CultureInfo.InvariantCulture)} columns, exceeding Excel's maximum of {A1.MaxColumns.ToString(CultureInfo.InvariantCulture)} columns.", nameof(table));
                }

                if (rowCount > A1.MaxRows) {
                    throw new ArgumentException($"DataTable '{requestedName}' has {rowCount.ToString(CultureInfo.InvariantCulture)} worksheet rows including headers, exceeding Excel's maximum of {A1.MaxRows.ToString(CultureInfo.InvariantCulture)} rows.", nameof(table));
                }
            }

            private static string GetUniqueName(string baseName, HashSet<string> used, int maxLength) {
                string trimmed = string.IsNullOrWhiteSpace(baseName) ? "Table" : baseName;
                if (trimmed.Length > maxLength) {
                    trimmed = trimmed.Substring(0, maxLength);
                }

                if (used.Add(trimmed)) {
                    return trimmed;
                }

                int suffix = 2;
                while (true) {
                    string suffixText = suffix.ToString(CultureInfo.InvariantCulture);
                    int prefixLength = Math.Max(1, maxLength - suffixText.Length);
                    string candidate = trimmed.Length > prefixLength
                        ? trimmed.Substring(0, prefixLength) + suffixText
                        : trimmed + suffixText;
                    if (used.Add(candidate)) {
                        return candidate;
                    }

                    suffix++;
                }
            }

            private static string SanitizeSheetName(string name) {
                string baseName = (name ?? string.Empty).Trim();
                baseName = baseName.Trim('\'', ' ');
                var builder = new StringBuilder(baseName.Length);
                foreach (char ch in baseName) {
                    builder.Append(ch is ':' or '\\' or '/' or '?' or '*' or '[' or ']' ? '_' : ch);
                }

                string value = _multipleUnderscoresRegex.Replace(builder.ToString().Trim(), "_");
                return value.Trim('_');
            }

            private static string SanitizeTableName(string name) {
                var builder = new StringBuilder(name.Length + 1);
                foreach (char ch in name) {
                    builder.Append(char.IsLetterOrDigit(ch) || ch == '_' ? ch : '_');
                }

                string value = builder.ToString();
                if (string.IsNullOrWhiteSpace(value)) {
                    value = "Table";
                }

                if (!char.IsLetter(value[0]) && value[0] != '_') {
                    value = "_" + value;
                }

                return value;
            }
        }

        private sealed class DirectDataSetSheetModel {
            internal DirectDataSetSheetModel(
                int index,
                string sheetName,
                string? tableName,
                string range,
                DirectDataSetTableModel table,
                TableStyle tableStyle,
                bool includeHeaders,
                bool includeAutoFilter,
                bool hasTable,
                bool autoFitColumns,
                bool omitBlankCells,
                double[]? columnWidths,
                bool useCellValueNumberFormats = false,
                DirectWorksheetMetadata? metadata = null,
                IReadOnlyList<string?>? columnNumberFormats = null,
                bool showFirstColumn = false,
                bool showLastColumn = false,
                bool showRowStripes = true,
                bool showColumnStripes = false) {
                Index = index;
                SheetName = sheetName;
                TableName = tableName;
                Range = range;
                Table = table;
                TableStyle = tableStyle;
                IncludeHeaders = includeHeaders;
                IncludeAutoFilter = includeAutoFilter;
                HasTable = hasTable;
                AutoFitColumns = autoFitColumns;
                OmitBlankCells = omitBlankCells;
                ColumnWidths = columnWidths;
                UseCellValueNumberFormats = useCellValueNumberFormats;
                Metadata = metadata;
                ColumnNumberFormats = columnNumberFormats;
                ShowFirstColumn = showFirstColumn;
                ShowLastColumn = showLastColumn;
                ShowRowStripes = showRowStripes;
                ShowColumnStripes = showColumnStripes;
            }

            internal DirectDataSetSheetModel WithMetadata(DirectWorksheetMetadata? metadata) {
                if (ReferenceEquals(Metadata, metadata)) {
                    return this;
                }

                return new DirectDataSetSheetModel(
                    Index,
                    SheetName,
                    TableName,
                    Range,
                    Table,
                    TableStyle,
                    IncludeHeaders,
                    IncludeAutoFilter,
                    HasTable,
                    AutoFitColumns,
                    OmitBlankCells,
                    ColumnWidths,
                    UseCellValueNumberFormats,
                    metadata,
                    ColumnNumberFormats,
                    ShowFirstColumn,
                    ShowLastColumn,
                    ShowRowStripes,
                    ShowColumnStripes);
            }

            internal DirectDataSetSheetModel WithTableStyle(
                TableStyle tableStyle,
                bool? showFirstColumn,
                bool? showLastColumn,
                bool? showRowStripes,
                bool? showColumnStripes) {
                return new DirectDataSetSheetModel(
                    Index,
                    SheetName,
                    TableName,
                    Range,
                    Table,
                    tableStyle,
                    IncludeHeaders,
                    IncludeAutoFilter,
                    HasTable,
                    AutoFitColumns,
                    OmitBlankCells,
                    ColumnWidths,
                    UseCellValueNumberFormats,
                    Metadata,
                    ColumnNumberFormats,
                    showFirstColumn ?? ShowFirstColumn,
                    showLastColumn ?? ShowLastColumn,
                    showRowStripes ?? ShowRowStripes,
                    showColumnStripes ?? ShowColumnStripes);
            }

            internal DirectDataSetSheetModel WithColumnNumberFormat(int columnIndex, string numberFormat) {
                if (columnIndex <= 0 || columnIndex > Table.ColumnCount) {
                    throw new ArgumentOutOfRangeException(nameof(columnIndex));
                }

                string?[] formats;
                if (ColumnNumberFormats == null || ColumnNumberFormats.Count != Table.ColumnCount) {
                    formats = new string?[Table.ColumnCount];
                } else {
                    formats = new string?[ColumnNumberFormats.Count];
                    for (int i = 0; i < formats.Length; i++) {
                        formats[i] = ColumnNumberFormats[i];
                    }
                }

                formats[columnIndex - 1] = numberFormat;
                return new DirectDataSetSheetModel(
                    Index,
                    SheetName,
                    TableName,
                    Range,
                    Table,
                    TableStyle,
                    IncludeHeaders,
                    IncludeAutoFilter,
                    HasTable,
                    AutoFitColumns,
                    OmitBlankCells,
                    ColumnWidths,
                    UseCellValueNumberFormats,
                    Metadata,
                    formats,
                    ShowFirstColumn,
                    ShowLastColumn,
                    ShowRowStripes,
                    ShowColumnStripes);
            }

            internal int Index { get; }

            internal string SheetName { get; }

            internal string? TableName { get; }

            internal string Range { get; }

            internal DirectDataSetTableModel Table { get; }

            internal TableStyle TableStyle { get; }

            internal bool IncludeHeaders { get; }

            internal bool IncludeAutoFilter { get; }

            internal bool HasTable { get; }

            internal bool AutoFitColumns { get; }

            internal bool OmitBlankCells { get; }

            internal double[]? ColumnWidths { get; }

            internal bool UseCellValueNumberFormats { get; }

            internal DirectWorksheetMetadata? Metadata { get; }

            internal IReadOnlyList<string?>? ColumnNumberFormats { get; }

            internal bool ShowFirstColumn { get; }

            internal bool ShowLastColumn { get; }

            internal bool ShowRowStripes { get; }

            internal bool ShowColumnStripes { get; }
        }

        private sealed class DirectWorksheetMetadata {
            internal static readonly DirectWorksheetMetadata Empty = new(
                null,
                null,
                null,
                null,
                null,
                Array.Empty<string>(),
                null,
                null,
                Array.Empty<string>(),
                Array.Empty<DirectOverlayCell>());

            internal DirectWorksheetMetadata(
                string? sheetPropertiesXml,
                string? sheetViewsXml,
                string? sheetFormatPropertiesXml,
                string? sheetProtectionXml,
                string? autoFilterXml,
                IReadOnlyList<string> conditionalFormattingXml,
                string? dataValidationsXml,
                string? drawingXml,
                IReadOnlyList<string> postDataValidationXml,
                IReadOnlyList<DirectOverlayCell> overlayCells) {
                SheetPropertiesXml = sheetPropertiesXml;
                SheetViewsXml = sheetViewsXml;
                SheetFormatPropertiesXml = sheetFormatPropertiesXml;
                SheetProtectionXml = sheetProtectionXml;
                AutoFilterXml = autoFilterXml;
                ConditionalFormattingXml = conditionalFormattingXml ?? Array.Empty<string>();
                DataValidationsXml = dataValidationsXml;
                DrawingXml = drawingXml;
                PostDataValidationXml = postDataValidationXml ?? Array.Empty<string>();
                OverlayCells = overlayCells ?? Array.Empty<DirectOverlayCell>();
            }

            internal DirectWorksheetMetadata WithSheetViewsXml(string? sheetViewsXml) {
                if (string.Equals(SheetViewsXml, sheetViewsXml, StringComparison.Ordinal)) {
                    return this;
                }

                return new DirectWorksheetMetadata(
                    SheetPropertiesXml,
                    sheetViewsXml,
                    SheetFormatPropertiesXml,
                    SheetProtectionXml,
                    AutoFilterXml,
                    ConditionalFormattingXml,
                    DataValidationsXml,
                    DrawingXml,
                    PostDataValidationXml,
                    OverlayCells);
            }

            internal DirectWorksheetMetadata WithAutoFilterXml(string? autoFilterXml) {
                if (string.Equals(AutoFilterXml, autoFilterXml, StringComparison.Ordinal)) {
                    return this;
                }

                return new DirectWorksheetMetadata(
                    SheetPropertiesXml,
                    SheetViewsXml,
                    SheetFormatPropertiesXml,
                    SheetProtectionXml,
                    autoFilterXml,
                    ConditionalFormattingXml,
                    DataValidationsXml,
                    DrawingXml,
                    PostDataValidationXml,
                    OverlayCells);
            }

            internal string? SheetPropertiesXml { get; }

            internal string? SheetViewsXml { get; }

            internal string? SheetFormatPropertiesXml { get; }

            internal string? SheetProtectionXml { get; }

            internal string? AutoFilterXml { get; }

            internal IReadOnlyList<string> ConditionalFormattingXml { get; }

            internal string? DataValidationsXml { get; }

            internal string? DrawingXml { get; }

            internal IReadOnlyList<string> PostDataValidationXml { get; }

            internal IReadOnlyList<DirectOverlayCell> OverlayCells { get; }

            internal bool IsEmpty
                => SheetPropertiesXml == null
                   && SheetViewsXml == null
                   && SheetFormatPropertiesXml == null
                   && SheetProtectionXml == null
                && AutoFilterXml == null
                && ConditionalFormattingXml.Count == 0
                && DataValidationsXml == null
                && DrawingXml == null
                && PostDataValidationXml.Count == 0
                && OverlayCells.Count == 0;
        }

        private readonly struct DirectOverlayCell {
            internal DirectOverlayCell(int row, int column, object? value, uint? styleIndex, string? numberFormat, bool isDeleted = false) {
                Row = row;
                Column = column;
                Value = value;
                StyleIndex = styleIndex;
                NumberFormat = numberFormat;
                IsDeleted = isDeleted;
            }

            internal int Row { get; }

            internal int Column { get; }

            internal object? Value { get; }

            internal uint? StyleIndex { get; }

            internal string? NumberFormat { get; }

            internal bool IsDeleted { get; }
        }

        private readonly struct DirectOverlayStyleResolution {
            internal DirectOverlayStyleResolution(bool supported, string? numberFormat) {
                Supported = supported;
                NumberFormat = numberFormat;
            }

            internal bool Supported { get; }

            internal string? NumberFormat { get; }
        }

        private readonly struct DirectBufferedRows {
            private readonly object?[][]? _arrayRows;
            private readonly List<object?[]>? _listRows;

            internal DirectBufferedRows(object?[][] rows) {
                _arrayRows = rows;
                _listRows = null;
            }

            internal DirectBufferedRows(List<object?[]> rows) {
                _arrayRows = null;
                _listRows = rows;
            }

            internal int Count => _arrayRows?.Length ?? _listRows!.Count;

            internal object?[] this[int index] => _arrayRows != null
                ? _arrayRows[index]
                : _listRows![index];
        }

        private readonly struct DirectCellValueRows {
            internal DirectCellValueRows(object?[] values, int columnCount, int rowCount, bool valuesMatchColumnTypes) {
                Values = values;
                ColumnCount = columnCount;
                Count = rowCount;
                ValuesMatchColumnTypes = valuesMatchColumnTypes;
            }

            internal object?[] Values { get; }

            internal int ColumnCount { get; }

            internal int Count { get; }

            internal bool ValuesMatchColumnTypes { get; }

            internal int GetRowOffset(int rowIndex) => rowIndex * ColumnCount;

            internal object? GetValue(int rowIndex, int columnIndex) {
                int index = GetRowOffset(rowIndex) + columnIndex;
                return Values[index];
            }
        }
    }
}
