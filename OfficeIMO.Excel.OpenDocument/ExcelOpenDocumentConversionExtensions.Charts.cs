using System.Globalization;
using OfficeIMO.Excel;
using OfficeIMO.OpenDocument;
using OfficeIMO.Spreadsheet;

namespace OfficeIMO.Excel.OpenDocument;

public static partial class ExcelOpenDocumentConversionExtensions {
    private const int MaximumConvertedChartSeries = 16;
    private const int MaximumConvertedChartPoints = 4096;

    private static int ConvertOdsCharts(OdsDocument source,
        IReadOnlyList<(OdsSheet Source, ExcelSheet Target)> chartTargets,
        ExcelOpenDocumentConversionOptions options, ref long expandedCells, ref bool truncated) {
        int converted = 0;
        long nextChartDataRow = 1;
        var readers = new Dictionary<string, ChartCellReader>(StringComparer.Ordinal);
        foreach ((OdsSheet odsSheet, ExcelSheet excelSheet) in chartTargets) {
            foreach (OdsChart chart in odsSheet.Charts) {
                if (!TryCreateExcelChartData(source, odsSheet.Name, chart, options, readers,
                    out ExcelChartData? data, out ExcelChartType chartType, out int row, out int column,
                    out int width, out int height, out bool sourceLimitExceeded)) {
                    if (sourceLimitExceeded) truncated = true;
                    continue;
                }
                if (data!.Series.Count + 1 > options.MaximumColumns) {
                    truncated = true;
                    continue;
                }
                long chartCells = ((long)data!.Categories.Count + 1) * (data.Series.Count + 1);
                if (chartCells > options.MaximumExpandedCells - expandedCells) {
                    truncated = true;
                    continue;
                }
                long reservedRows = (long)data.Categories.Count + 3;
                if (nextChartDataRow + reservedRows - 1 > options.MaximumRows) {
                    truncated = true;
                    continue;
                }
                excelSheet.AddChart(data!, row, column, width, height, chartType, chart.Title);
                expandedCells += chartCells;
                nextChartDataRow += reservedRows;
                converted++;
            }
        }
        return converted;
    }

    private static bool TryCreateExcelChartData(OdsDocument document, string hostSheetName, OdsChart chart,
        ExcelOpenDocumentConversionOptions options, Dictionary<string, ChartCellReader> readers,
        out ExcelChartData? data,
        out ExcelChartType type, out int row, out int column, out int width, out int height,
        out bool sourceLimitExceeded) {
        data = null;
        type = ExcelChartType.ColumnClustered;
        row = column = width = height = 0;
        sourceLimitExceeded = false;
        if (chart.AnchorRow.HasValue && chart.AnchorColumn.HasValue &&
            (chart.AnchorRow.Value >= options.MaximumRows || chart.AnchorColumn.Value >= options.MaximumColumns)) {
            sourceLimitExceeded = true;
            return false;
        }
        if (chart.IsStacked || chart.IsPercentage || chart.IsThreeDimensional
            || chart.TitleCellRangeAddress != null
            || chart.Series.Count < 1 || chart.Series.Count > MaximumConvertedChartSeries
            || !chart.AnchorRow.HasValue || !chart.AnchorColumn.HasValue
            || !chart.Bounds.Width.TryToPoints(out double widthPoints)
            || !chart.Bounds.Height.TryToPoints(out double heightPoints)
            || widthPoints <= 0 || heightPoints <= 0
            || widthPoints > 1500 || heightPoints > 1500) return false;
        switch (chart.ChartClass) {
            case "chart:bar":
                type = chart.VerticalBars == true ? ExcelChartType.BarClustered : ExcelChartType.ColumnClustered;
                break;
            case "chart:line": type = ExcelChartType.Line; break;
            default: return false;
        }
        if (!TryReadChartCells(document, hostSheetName, chart.CategoriesAddress, options, readers,
            out OdsCellValue[] categories, out bool categoryLimitExceeded)) {
            sourceLimitExceeded = categoryLimitExceeded;
            return false;
        }
        if (categories.Length < 1 || categories.Length > MaximumConvertedChartPoints) return false;
        var labels = new string[categories.Length];
        for (int index = 0; index < labels.Length; index++) {
            labels[index] = categories[index].DisplayText.Length > 0
                ? categories[index].DisplayText : categories[index].LexicalValue;
        }
        var series = new List<ExcelChartSeries>(chart.Series.Count);
        for (int seriesIndex = 0; seriesIndex < chart.Series.Count; seriesIndex++) {
            OdsChartSeries sourceSeries = chart.Series[seriesIndex];
            if (sourceSeries.ChartClass != null && sourceSeries.ChartClass != chart.ChartClass) return false;
            if (!TryReadChartCells(document, hostSheetName, sourceSeries.ValuesAddress, options, readers,
                out OdsCellValue[] values, out bool valuesLimitExceeded)) {
                sourceLimitExceeded = valuesLimitExceeded;
                return false;
            }
            if (values.Length != labels.Length) return false;
            string name = "Series " + (seriesIndex + 1).ToString(CultureInfo.InvariantCulture);
            if (sourceSeries.LabelAddress != null) {
                if (!TryReadChartCells(document, hostSheetName, sourceSeries.LabelAddress, options, readers,
                    out OdsCellValue[] nameCell, out bool labelLimitExceeded)) {
                    sourceLimitExceeded = labelLimitExceeded;
                    return false;
                }
                if (nameCell.Length != 1) return false;
                name = nameCell[0].DisplayText.Length > 0 ? nameCell[0].DisplayText : nameCell[0].LexicalValue;
                if (string.IsNullOrWhiteSpace(name)) return false;
            }
            var numbers = new double[values.Length];
            for (int index = 0; index < values.Length; index++) {
                if (values[index].Kind is not (OdsCellValueKind.Number or OdsCellValueKind.Percentage or OdsCellValueKind.Currency)
                    || !double.TryParse(values[index].LexicalValue, NumberStyles.Float,
                        CultureInfo.InvariantCulture, out double number)
                    || double.IsNaN(number) || double.IsInfinity(number)) return false;
                numbers[index] = number;
            }
            series.Add(new ExcelChartSeries(name, numbers));
        }
        data = new ExcelChartData(labels, series);
        row = checked((int)chart.AnchorRow.Value + 1);
        column = checked((int)chart.AnchorColumn.Value + 1);
        width = Math.Max(1, (int)Math.Round(widthPoints * 96D / 72D));
        height = Math.Max(1, (int)Math.Round(heightPoints * 96D / 72D));
        return true;
    }

    private static bool TryReadChartCells(OdsDocument document, string hostSheetName, string? address,
        ExcelOpenDocumentConversionOptions options, Dictionary<string, ChartCellReader> readers,
        out OdsCellValue[] values, out bool sourceLimitExceeded) {
        values = Array.Empty<OdsCellValue>();
        sourceLimitExceeded = false;
        if (!SpreadsheetRangeReference.TryParse(address, SpreadsheetAddressDialect.OpenDocument,
                out SpreadsheetRangeReference? range) || !range!.Start.IsCell) return false;
        SpreadsheetCellReference start = range.Start;
        SpreadsheetCellReference end = range.End ?? start;
        string startSheetName = start.SheetName ?? hostSheetName;
        string endSheetName = end.SheetName ?? hostSheetName;
        if (!end.IsCell || !string.Equals(startSheetName, endSheetName, StringComparison.Ordinal)) return false;
        long firstRow = start.Row!.Value, lastRow = end.Row!.Value;
        int firstColumn = start.Column!.Value, lastColumn = end.Column!.Value;
        if (lastRow > options.MaximumRows || lastColumn > options.MaximumColumns) {
            sourceLimitExceeded = true;
            return false;
        }
        if (firstRow > lastRow || firstColumn > lastColumn
            || (firstRow != lastRow && firstColumn != lastColumn)) return false;
        long count = firstRow == lastRow ? lastColumn - firstColumn + 1L : lastRow - firstRow + 1;
        if (count < 1 || count > MaximumConvertedChartPoints) return false;
        if (!readers.TryGetValue(startSheetName, out ChartCellReader? reader)) {
            OdsSheet? sheet = document.GetSheet(startSheetName);
            if (sheet == null) return false;
            reader = new ChartCellReader(sheet);
            readers.Add(startSheetName, reader);
        }
        values = new OdsCellValue[checked((int)count)];
        for (int index = 0; index < values.Length; index++) {
            long row = firstRow + (firstRow == lastRow ? 0 : index);
            int column = firstColumn + (firstColumn == lastColumn ? 0 : index);
            values[index] = reader.GetValue(row - 1, column - 1);
        }
        return true;
    }

    private sealed class ChartCellReader {
        private readonly IReadOnlyList<OdsRowRun> _rows;
        private readonly Dictionary<int, IReadOnlyList<OdsCellRun>> _cells = new Dictionary<int, IReadOnlyList<OdsCellRun>>();

        internal ChartCellReader(OdsSheet sheet) => _rows = sheet.RowRuns;

        internal OdsCellValue GetValue(long row, int column) {
            int low = 0, high = _rows.Count - 1;
            while (low <= high) {
                int middle = low + (high - low) / 2;
                OdsRowRun run = _rows[middle];
                if (row < run.StartRow) high = middle - 1;
                else if (row - run.StartRow >= run.RepeatCount) low = middle + 1;
                else {
                    if (!_cells.TryGetValue(middle, out IReadOnlyList<OdsCellRun>? cells)) {
                        cells = run.CellRuns;
                        _cells.Add(middle, cells);
                    }
                    int cellLow = 0, cellHigh = cells.Count - 1;
                    while (cellLow <= cellHigh) {
                        int cellMiddle = cellLow + (cellHigh - cellLow) / 2;
                        OdsCellRun cell = cells[cellMiddle];
                        if (column < cell.StartColumn) cellHigh = cellMiddle - 1;
                        else if (column - cell.StartColumn >= cell.RepeatCount) cellLow = cellMiddle + 1;
                        else return cell.Value;
                    }
                    return OdsCellValue.Empty;
                }
            }
            return OdsCellValue.Empty;
        }
    }
}
