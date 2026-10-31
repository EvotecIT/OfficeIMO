using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;
using PSWriteOffice.Services.Table;

namespace PSWriteOffice.Cmdlets.Excel;

public sealed partial class AddOfficeExcelChartCommand {
    /// <summary>Prepares source selections without modifying worksheet cells or drawings.</summary>
    private ChartPlan PlanWorksheet(ExcelSheet sheet) {
        if (ParameterSetName == RangeSet) {
            var bounds = ParseSourceRange(Range);
            int firstDataRow = bounds.Top + (NoHeader.IsPresent ? 0 : 1);
            if (firstDataRow > bounds.Bottom || bounds.Left == bounds.Right) {
                throw new PSArgumentException("Range must contain at least one category row and one value column.", nameof(Range));
            }

            string[] headers = Enumerable.Range(bounds.Left, bounds.Right - bounds.Left + 1)
                .Select(column => NoHeader.IsPresent ? "Series " + (column - bounds.Left).ToString(CultureInfo.InvariantCulture) : ReadText(sheet, bounds.Top, column))
                .ToArray();
            var data = ReadData(sheet, firstDataRow, bounds.Bottom, bounds.Left,
                Enumerable.Range(bounds.Left + 1, bounds.Right - bounds.Left).ToArray(), headers.Skip(1).ToArray());
            return new ChartPlan(data, Range, !NoHeader.IsPresent);
        }

        var table = sheet.GetTables().SingleOrDefault(item =>
            string.Equals(item.Name, Table, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(item.DisplayName, Table, StringComparison.OrdinalIgnoreCase));
        if (table == null) throw new PSArgumentException($"Table '{Table}' was not found on sheet '{sheet.Name}'.", nameof(Table));
        var tableBounds = ParseSourceRange(table.Range);
        int firstRow = tableBounds.Top + (table.HasHeaderRow ? 1 : 0);
        int lastRow = tableBounds.Bottom - (table.TotalsRowShown ? 1 : 0);
        if (firstRow > lastRow) throw new PSArgumentException($"Table '{Table}' contains no data rows.", nameof(Table));
        string[] names = table.Columns.Select(column => column.Name).ToArray();
        if (names.Length != tableBounds.Right - tableBounds.Left + 1) {
            throw new PSArgumentException($"Table '{Table}' has an inconsistent column definition.", nameof(Table));
        }

        int categoryIndex = FindColumn(names, Category ?? names[0], nameof(Category));
        int[] valueIndexes = SelectValues(names, categoryIndex,
            index => Enumerable.Range(firstRow, lastRow - firstRow + 1).All(row =>
                TryReadNumber(sheet, row, tableBounds.Left + index, out _)));
        int[] valueColumns = valueIndexes.Select(index => tableBounds.Left + index).ToArray();
        var chartData = ReadData(sheet, firstRow, lastRow, tableBounds.Left + categoryIndex,
            valueColumns, valueIndexes.Select(index => names[index]).ToArray());
        bool contiguous = valueIndexes.Select((index, offset) => index == categoryIndex + offset + 1).All(matches => matches);
        string? source = contiguous
            ? ExcelA1.CellReference(table.HasHeaderRow ? tableBounds.Top : firstRow, tableBounds.Left + categoryIndex) + ":" +
              ExcelA1.CellReference(lastRow, valueColumns[valueColumns.Length - 1])
            : null;
        return new ChartPlan(chartData, source, table.HasHeaderRow);
    }

    /// <summary>Normalizes pipeline input through the shared table owner and reserves unused worksheet rows.</summary>
    private ChartPlan PlanInput(ExcelSheet sheet) {
        var rows = TableInputCollector.RequireRows(_items, nameof(InputObject));
        var table = ExcelTabularInputService.ToDataTable(rows);
        if (table.Rows.Count == 0 || table.Columns.Count < 2) {
            throw new PSArgumentException("Provide at least one data row, a category and a value column.", nameof(InputObject));
        }
        var names = table.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToArray();
        int categoryIndex = FindColumn(names, Category!, nameof(Category));
        int[] valueIndexes = SelectValues(names, categoryIndex, _ => false);
        var categories = table.Rows.Cast<DataRow>().Select(row => CategoryText(row[categoryIndex])).ToArray();
        var series = valueIndexes.Select(index => new OfficeSeries(names[index],
            table.Rows.Cast<DataRow>().Select(row => RequireNumber(row[index], names[index])))).ToArray();
        var chartData = new OfficeChartData(categories, series);

        // Keep all supplied fields while putting the selected chart fields in one native source block.
        table.Columns[names[categoryIndex]]!.SetOrdinal(0);
        for (int i = 0; i < valueIndexes.Length; i++) table.Columns[names[valueIndexes[i]]]!.SetOrdinal(i + 1);
        var used = ParseSourceRange(sheet.UsedRangeA1);
        int startRow = used.Top == 1 && used.Bottom == 1 && used.Left == 1 && used.Right == 1 &&
            !sheet.TryGetCellText(1, 1, out _) ? 1 : checked(used.Bottom + 2);
        if ((long)startRow + table.Rows.Count > ExcelA1.MaxRows || table.Columns.Count > ExcelA1.MaxColumns) {
            throw new PSArgumentException("There is not enough unused worksheet space for the chart data table.", nameof(InputObject));
        }
        string range = ExcelA1.CellReference(startRow, 1) + ":" + ExcelA1.CellReference(startRow + table.Rows.Count, series.Length + 1);
        return new ChartPlan(chartData, range, true, table, startRow);
    }

    /// <summary>Resolves explicit headers or infers fully numeric columns, excluding the category and duplicate selections.</summary>
    private int[] SelectValues(string[] names, int category, Func<int, bool> isNumeric) {
        int[] indexes;
        if (Value == null) {
            indexes = Enumerable.Range(0, names.Length).Where(index => index != category && isNumeric(index)).ToArray();
        } else {
            if (Value.Length == 0) throw new PSArgumentException("Provide at least one value column.", nameof(Value));
            indexes = Value.Select(name => FindColumn(names, name, nameof(Value))).ToArray();
            if (indexes.Contains(category)) throw new PSArgumentException("A category column cannot also be a value column.", nameof(Value));
            if (indexes.Distinct().Count() != indexes.Length) throw new PSArgumentException("Value columns must be distinct.", nameof(Value));
        }
        if (indexes.Length == 0) throw new PSArgumentException("No numeric value columns were found; select numeric columns with -Value.", nameof(Value));
        return indexes;
    }

    private static int FindColumn(string[] names, string name, string parameter) {
        if (string.IsNullOrWhiteSpace(name)) throw new PSArgumentException("Column names cannot be empty.", parameter);
        var matches = names.Select((header, index) => new { header, index })
            .Where(item => string.Equals(item.header, name, StringComparison.OrdinalIgnoreCase)).ToArray();
        if (matches.Length != 1) throw new PSArgumentException($"Column '{name}' must match exactly one source header.", parameter);
        return matches[0].index;
    }

    private static OfficeChartData ReadData(ExcelSheet sheet, int firstRow, int lastRow, int categoryColumn,
        int[] valueColumns, string[] seriesNames) {
        int count = lastRow - firstRow + 1;
        var categories = Enumerable.Range(firstRow, count).Select(row => CategoryText(ReadText(sheet, row, categoryColumn))).ToArray();
        var series = valueColumns.Select((column, index) => new OfficeSeries(seriesNames[index],
            Enumerable.Range(firstRow, count).Select(row => TryReadNumber(sheet, row, column, out double value)
                ? value : throw new PSArgumentException($"Column '{seriesNames[index]}' contains a blank, nonnumeric or nonfinite value at {ExcelA1.CellReference(row, column)}.", nameof(Value))))).ToArray();
        return new OfficeChartData(categories, series);
    }

    private static bool TryReadNumber(ExcelSheet sheet, int row, int column, out double value) {
        value = 0;
        if (!sheet.TryGetCellValueSnapshot(row, column, out var cell) || cell == null || cell.DateTimeValue.HasValue ||
            cell.Kind is ExcelCellValueKind.Boolean or ExcelCellValueKind.Error or ExcelCellValueKind.DateTime) return false;
        if (cell.Kind == ExcelCellValueKind.Formula) {
            // Formula RawValue is the expression. Read its stored result through the native cache owner.
            if (cell.OpenXmlType.HasValue && cell.OpenXmlType.Value != ExcelCellValueType.Number) return false;
            return sheet.TryGetCachedFormulaValue(row, column, out string? cached) && TryNumber(cached, out value);
        }
        return cell.Kind switch {
            ExcelCellValueKind.Number => TryNumber(cell.RawValue, out value),
            // Text columns may contain numeric CSV values; resolve shared strings rather than parsing their indexes.
            ExcelCellValueKind.Text => TryNumber(cell.Text, out value),
            _ => false
        };
    }

    private static bool TryNumber(object? input, out double value) {
        value = 0;
        if (input == null || input == DBNull.Value || input is bool or DateTime or DateTimeOffset or TimeSpan ||
            input is not (string or byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal)) return false;
        return double.TryParse(Convert.ToString(input, CultureInfo.InvariantCulture), NumberStyles.Float,
            CultureInfo.InvariantCulture, out value) && !double.IsNaN(value) && !double.IsInfinity(value);
    }

    private static double RequireNumber(object value, string name) => TryNumber(value, out double number)
        ? number : throw new PSArgumentException($"Column '{name}' contains a blank, nonnumeric or nonfinite value.", nameof(Value));

    private static string ReadText(ExcelSheet sheet, int row, int column) => sheet.TryGetCellText(row, column, out string text) ? text : string.Empty;

    private static string CategoryText(object value) {
        string text = Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
        if (string.IsNullOrWhiteSpace(text)) throw new PSArgumentException("Category values cannot be blank.", nameof(Category));
        return text;
    }

    private static (int Top, int Left, int Bottom, int Right) ParseSourceRange(string range) {
        if (!ExcelA1.TryParseRange(range, out int top, out int left, out int bottom, out int right)) {
            var cell = ExcelA1.ParseCellRef(range);
            top = bottom = cell.Row;
            left = right = cell.Col;
        }
        if (top < 1 || left < 1 || bottom < top || right < left || bottom > ExcelA1.MaxRows || right > ExcelA1.MaxColumns) {
            throw new PSArgumentException($"'{range}' must be a bounded worksheet A1 range.", nameof(Range));
        }
        return (top, left, bottom, right);
    }

    private sealed class ChartPlan {
        internal ChartPlan(OfficeChartData data, string? sourceRange, bool hasHeaders, DataTable? inputTable = null, int startRow = 0) {
            Data = data;
            SourceRange = sourceRange;
            HasHeaders = hasHeaders;
            InputTable = inputTable;
            StartRow = startRow;
        }
        internal OfficeChartData Data { get; }
        internal string? SourceRange { get; }
        internal bool HasHeaders { get; }
        internal DataTable? InputTable { get; }
        internal int StartRow { get; }
    }
}
