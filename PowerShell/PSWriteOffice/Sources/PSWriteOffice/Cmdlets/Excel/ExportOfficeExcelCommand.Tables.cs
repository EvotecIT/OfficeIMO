using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Security;
using System.Threading.Tasks;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;
using System.Data;

namespace PSWriteOffice.Cmdlets.Excel;

public sealed partial class ExportOfficeExcelCommand {
    private bool CanExportReaderDirectly(bool isAppendingToExistingSheet)
    {
        return !isAppendingToExistingSheet && ExcludeProperty is not { Length: > 0 };
    }

    private bool TryExportObjectsThroughOfficeImoObjectPath(
        ExcelSheet sheet,
        bool isAppendingToExistingSheet,
        bool preserveWorkbook,
        int dataStartRow,
        bool includeHeaders,
        ExcelTableStyle style,
        ExcelColumnFormatPlan? columnFormatPlan,
        out string range)
    {
        range = string.Empty;
        if (!CanExportObjectsThroughOfficeImoObjectPath(isAppendingToExistingSheet, preserveWorkbook, dataStartRow, includeHeaders, columnFormatPlan))
        {
            return false;
        }

        var items = _input.Where(static item => item != null).ToList();
        if (items.Count == 0)
        {
            return false;
        }

        if (!CanUseOfficeImoObjectPathInput(items))
        {
            return false;
        }

        var normalizedItems = PowerShellObjectNormalizer
            .NormalizeItems(items, CreateNormalizerOptions())
            .ToList();
        sheet.WriteData<object?>(normalizedItems, includeHeaders, dataStartRow);
        range = sheet.UsedRange.Address;
        if (string.IsNullOrWhiteSpace(range))
        {
            return false;
        }

        if (!NoTable.IsPresent)
        {
            sheet.AddTable(range, includeHeaders, TableName ?? string.Empty, style, includeAutoFilter: !NoAutoFilter.IsPresent);
            ApplyTableStyleOptions(sheet, range, style);
        }

        ApplyColumnFormatPlan(sheet, columnFormatPlan, includeHeaders ? dataStartRow : null);
        return true;
    }

    private bool CanExportObjectsThroughOfficeImoObjectPath(
        bool isAppendingToExistingSheet,
        bool preserveWorkbook,
        int dataStartRow,
        bool includeHeaders,
        ExcelColumnFormatPlan? columnFormatPlan)
    {
        return !isAppendingToExistingSheet &&
            !preserveWorkbook &&
            includeHeaders &&
            dataStartRow == 1 &&
            StartColumn == 1 &&
            string.IsNullOrWhiteSpace(Title) &&
            ExcludeProperty is not { Length: > 0 } &&
            !AutoFit.IsPresent &&
            !AutoFitFormattedColumn.IsPresent &&
            !ColumnFormatRequiresMaterializedCells(columnFormatPlan) &&
            !BoldTopRow.IsPresent &&
            !FreezeTopRow.IsPresent &&
            !FreezeFirstColumn.IsPresent;
    }

    private static bool CanUseOfficeImoObjectPathInput(IReadOnlyList<object?> items)
    {
        foreach (var item in items)
        {
            if (item is PSObject)
            {
                return false;
            }

            if (item is Dictionary<string, object?> ||
                item is IReadOnlyDictionary<string, object?> ||
                item is IDictionary<string, object?> ||
                item is IDictionary)
            {
                continue;
            }

            return false;
        }

        return items.Count > 0;
    }

    private string ExportDataReader(ExcelSheet sheet, IDataReader reader, ExcelTableStyle style, ExcelColumnFormatPlan? columnFormatPlan)
    {
        var dataStartRow = StartRow;
        var fieldCount = reader.FieldCount;
        if (!string.IsNullOrWhiteSpace(Title))
        {
            sheet.Cell(dataStartRow, StartColumn).SetValue(Title!);
            sheet.Cell(dataStartRow, StartColumn).SetBold(true);
            dataStartRow++;
        }

        var range = sheet.InsertDataReader(
            reader,
            startRow: dataStartRow,
            startColumn: StartColumn,
            includeHeaders: !NoHeader.IsPresent,
            tableName: TableName,
            style: style,
            includeAutoFilter: !NoAutoFilter.IsPresent,
            createTable: !NoTable.IsPresent,
            autoFit: false);

        if (!string.IsNullOrWhiteSpace(range))
        {
            ApplyColumnFormatPlan(sheet, columnFormatPlan, !NoHeader.IsPresent ? dataStartRow : null, requireExplicitHeaderRow: NoHeader.IsPresent);

            if (!NoTable.IsPresent)
            {
                ApplyTableStyleOptions(sheet, range, style);
            }

            if (BoldTopRow.IsPresent && !NoHeader.IsPresent)
            {
                BoldRow(sheet, dataStartRow, StartColumn, fieldCount);
            }

            if (AutoFit.IsPresent)
            {
                sheet.AutoFitColumnsFor(Enumerable.Range(StartColumn, fieldCount));
            }

            if (FreezeTopRow.IsPresent || FreezeFirstColumn.IsPresent)
            {
                var frozenRows = FreezeTopRow.IsPresent ? Math.Max(1, !NoHeader.IsPresent ? dataStartRow : StartRow) : 0;
                var frozenColumns = FreezeFirstColumn.IsPresent ? Math.Max(1, StartColumn) : 0;
                sheet.View.Freeze(frozenRows, frozenColumns);
            }
        }

        return range;
    }

    private void ExportDataSet(ExcelDocument document, DataSet dataSet, ExcelTableStyle style, ExcelColumnFormatPlan? columnFormatPlan)
    {
        if (dataSet.Tables.Count == 0)
        {
            throw new PSArgumentException("DataSet must contain at least one DataTable.", nameof(InputObject));
        }

        var sheetNames = BuildDataSetWorksheetNames(document, dataSet, Append.IsPresent || ReplaceSheet.IsPresent);
        var usedTableNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var formattedHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var tableIndex = 1;
        foreach (DataTable sourceTable in dataSet.Tables)
        {
            var sheetName = sheetNames[tableIndex - 1];

            var sheetExists = SheetExists(document, sheetName);
            if (ReplaceSheet.IsPresent && sheetExists)
            {
                document.RemoveWorksheet(sheetName);
                sheetExists = false;
            }

            var sheet = GetOrAddDataSetSheet(document, sheetName);
            var table = PrepareDataTableForExport(sourceTable);
            if (table.Columns.Count == 0)
            {
                throw new InvalidOperationException($"Unable to infer columns from DataTable '{sourceTable.TableName}'.");
            }

            var isAppendingToExistingSheet = Append.IsPresent && sheetExists;
            var dataStartRow = ResolveDataStartRow(sheet, isAppendingToExistingSheet);
            var includeHeaders = !NoHeader.IsPresent && !isAppendingToExistingSheet;
            var appendTableName = ResolveAppendTableName(document, sheet, isAppendingToExistingSheet);
            var tableName = ResolveDataSetTableName(table, usedTableNames);
            var range = WriteData(sheet, table, dataStartRow, includeHeaders, style, isAppendingToExistingSheet, appendTableName, tableName);
            foreach (var result in ApplyColumnFormatPlan(
                         sheet,
                         columnFormatPlan,
                         ResolveColumnFormatHeaderRow(document, sheet, columnFormatPlan, isAppendingToExistingSheet, appendTableName, includeHeaders, dataStartRow),
                         throwOnMissing: false,
                         requireExplicitHeaderRow: !includeHeaders))
            {
                if (result.Applied)
                {
                    formattedHeaders.Add(result.Header);
                }
            }

            if (BoldTopRow.IsPresent && includeHeaders)
            {
                BoldRow(sheet, dataStartRow, StartColumn, table.Columns.Count);
            }

            if (AutoFit.IsPresent)
            {
                sheet.AutoFitColumnsFor(Enumerable.Range(StartColumn, table.Columns.Count));
            }

            if (FreezeTopRow.IsPresent || FreezeFirstColumn.IsPresent)
            {
                var frozenRows = ResolveFrozenTopRows(document, sheet, isAppendingToExistingSheet, appendTableName, dataStartRow, includeHeaders);
                var frozenColumns = FreezeFirstColumn.IsPresent ? Math.Max(1, StartColumn) : 0;
                sheet.View.Freeze(frozenRows, frozenColumns);
            }

            if (!string.IsNullOrWhiteSpace(range))
            {
                WriteVerbose($"Exported DataTable '{sourceTable.TableName}' to {sheet.Name}!{range}.");
            }

            tableIndex++;
        }

        ThrowForDataSetColumnFormatsMissingEverywhere(columnFormatPlan, formattedHeaders);
    }

    private DataTable PrepareDataTableForExport(DataTable sourceTable)
    {
        var table = ExcelTabularInputService.ToDataTable(
            new object?[] { sourceTable },
            sourceTable.TableName,
            copyExistingTables: ExcludeProperty is { Length: > 0 },
            normalizerOptions: CreateNormalizerOptions());
        return ApplyExcludedColumns(table);
    }

    private string? ResolveDataSetTableName(DataTable table, ISet<string> usedTableNames)
    {
        var tableName = ResolveTableName(table, allowExplicitOverride: false);
        if (string.IsNullOrWhiteSpace(tableName))
        {
            return null;
        }

        var candidate = tableName!;
        var suffix = 2;
        while (!usedTableNames.Add(candidate))
        {
            candidate = $"{tableName}_{suffix}";
            suffix++;
        }

        return candidate;
    }

    private ExcelSheet PrepareSheet(ExcelDocument document)
    {
        if (ReplaceSheet.IsPresent && SheetExists(document, WorksheetName))
        {
            document.RemoveWorksheet(WorksheetName);
        }

        return document.GetOrCreateSheet(WorksheetName, ExcelSheetNameValidationMode.Sanitize);
    }

    private DataTable BuildDataTable()
    {
        var table = ExcelTabularInputService.ToDataTable(
            _input,
            TableName,
            copyExistingTables: ExcludeProperty is { Length: > 0 },
            normalizerOptions: CreateNormalizerOptions());
        return ApplyExcludedColumns(table);
    }

    private PowerShellObjectNormalizerOptions CreateNormalizerOptions()
    {
        return new PowerShellObjectNormalizerOptions
        {
            CollectionSeparator = CollectionSeparator,
            DictionaryEntrySeparator = DictionaryEntrySeparator,
            DictionaryKeyValueSeparator = DictionaryKeyValueSeparator,
            MaxCollectionItems = MaxCollectionItems,
            MaxNestingDepth = MaxNestingDepth,
            IncludeUnexportableProperties = IncludeUnexportableProperties.IsPresent,
            PropertyErrorAction = PropertyConversionErrorAction,
            PropertyErrorCallback = PropertyConversionErrorAction == ActionPreference.Continue
                ? (name, exception) => WriteWarning($"Skipping property '{name}' because it could not be read: {exception.Message}")
                : null,
            UnexportablePropertyValueFactory = static (_, exception) => $"Property export failed: {exception.Message}"
        };
    }

    private DataTable ApplyExcludedColumns(DataTable table)
    {
        if (ExcludeProperty is { Length: > 0 })
        {
            var excluded = new HashSet<string>(
                ExcludeProperty.Where(static p => !string.IsNullOrWhiteSpace(p)).Select(static p => p.Trim()),
                StringComparer.OrdinalIgnoreCase);

            foreach (DataColumn column in table.Columns.Cast<DataColumn>().ToArray())
            {
                if (excluded.Contains(column.ColumnName))
                {
                    table.Columns.Remove(column);
                }
            }
        }

        return table;
    }

    private void WritePassThru(string resolvedPath)
    {
        if (PassThru.IsPresent)
        {
            WriteObject(new FileInfo(resolvedPath));
        }
    }

    private string WriteData(ExcelSheet sheet, DataTable table, int startRow, bool includeHeaders, ExcelTableStyle style, bool appendRawRows, string? appendTableName, string? tableName)
    {
        if (appendRawRows &&
            !NoTable.IsPresent &&
            !string.IsNullOrWhiteSpace(appendTableName) &&
            TryAppendDataTableToTable(sheet, table, appendTableName!, out var updatedTableRange))
        {
            return updatedTableRange;
        }

        if (!NoTable.IsPresent && !appendRawRows)
        {
            var range = sheet.AddTable(table, new OfficeIMO.Excel.ExcelTableOptions { AutoFit = false, UseDefaultHeaderlessColumnNames = true, IncludeHeaders = includeHeaders, Name = tableName, Style = style, IncludeAutoFilter = !NoAutoFilter.IsPresent }, row: startRow, column: StartColumn).Range!;
            ApplyTableStyleOptions(sheet, range, style);
            return range;
        }

        var cellCount = checked((table.Rows.Count + (includeHeaders ? 1 : 0)) * table.Columns.Count);
        var cells = new List<(int Row, int Column, object Value)>(cellCount);
        var row = startRow;

        if (includeHeaders)
        {
            for (var columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++)
            {
                cells.Add((row, StartColumn + columnIndex, table.Columns[columnIndex].ColumnName));
            }
            row++;
        }

        foreach (DataRow dataRow in table.Rows)
        {
            for (var columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++)
            {
                var value = dataRow[columnIndex];
                cells.Add((row, StartColumn + columnIndex, value is DBNull ? string.Empty : value));
            }
            row++;
        }

        sheet.SetValues(cells);
        var endRow = Math.Max(startRow, row - 1);
        var endColumn = StartColumn + table.Columns.Count - 1;
        return $"{ExcelA1.CellReference(startRow, StartColumn)}:{ExcelA1.CellReference(endRow, endColumn)}";
    }

    private string? ResolveAppendTableName(ExcelDocument document, ExcelSheet sheet, bool appendToExistingSheet) {
        if (AppendsToNamedTable && !appendToExistingSheet)
            throw new PSArgumentException($"Worksheet '{sheet.Name}' must exist when appending to a named table.");
        return AppendsToNamedTable ? TableName : null;
    }

    private string? ResolveTableName(DataTable table, bool allowExplicitOverride = true)
    {
        if (allowExplicitOverride && !string.IsNullOrWhiteSpace(TableName))
        {
            return TableName;
        }

        return string.IsNullOrWhiteSpace(table.TableName) ? null : table.TableName;
    }

    private bool TryAppendDataTableToTable(ExcelSheet sheet, DataTable table, string tableName, out string range)
    {
        range = sheet.AppendDataTableToTable(table, tableName, matchColumnsByHeader: true);
        return true;
    }

    private int ResolveDataStartRow(ExcelSheet sheet, bool appendToExistingSheet)
    {
        if (!appendToExistingSheet || StartRow > 1)
        {
            return StartRow;
        }

        var usedRange = sheet.UsedRange.Address;
        if (ExcelA1.TryParseRange(usedRange, out _, out _, out var lastRow, out _))
        {
            return lastRow + 1;
        }

        var (row, _) = ExcelA1.ParseCellRef(usedRange);
        return row > 0 ? row + 1 : StartRow;
    }

    private void AddInput(object? value)
    {
        if (value == null)
        {
            return;
        }

        if (value is PSObject psObject)
        {
            value = psObject.BaseObject is PSCustomObject ? psObject : psObject.BaseObject;
        }

        if (value is DataSet || value is DataTable || value is DataView || value is IDataReader)
        {
            _input.Add(value);
            return;
        }

        if (value is IEnumerable enumerable &&
            value is not string &&
            !PowerShellDictionaryAdapter.IsDictionaryLike(value))
        {
            foreach (var item in enumerable)
            {
                if (item != null)
                {
                    _input.Add(item);
                }
            }
            return;
        }

        _input.Add(value);
    }

    private static void BoldRow(ExcelSheet sheet, int row, int startColumn, int columnCount)
    {
        for (var column = startColumn; column < startColumn + columnCount; column++)
        {
            sheet.Cell(row, column).SetBold(true);
        }
    }


}
