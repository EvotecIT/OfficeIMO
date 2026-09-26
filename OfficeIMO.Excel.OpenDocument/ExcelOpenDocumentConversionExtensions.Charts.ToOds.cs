using OfficeIMO.Excel;
using OfficeIMO.OpenDocument;
using OfficeIMO.Drawing;
using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel.OpenDocument;

public static partial class ExcelOpenDocumentConversionExtensions {
    private static int ConvertExcelCharts(ExcelDocument source, OdsDocument target,
        ExcelOpenDocumentConversionOptions options,
        Dictionary<string, HashSet<(int Row, int Column)>> convertedCellsBySheet,
        ref long materializedCells, ref bool truncated, out int sourceChartFrames) {
        int converted = 0;
        sourceChartFrames = 0;
        foreach (ExcelSheet sourceSheet in source.Sheets) {
            ExcelChart[] charts = sourceSheet.Charts.ToArray();
            sourceChartFrames += charts.Length;
            OdsSheet? targetSheet = target.GetSheet(sourceSheet.Name);
            if (targetSheet == null || !convertedCellsBySheet.TryGetValue(sourceSheet.Name,
                    out HashSet<(int Row, int Column)>? anchorCells)) continue;
            foreach (ExcelChart chart in charts) {
                try {
                    if (chart.IsPivotChart || chart.HasAbsoluteAnchor
                        || !TryGetOdsChartType(chart.ChartType, out OdsChartType type)
                        || !chart.TryGetSnapshot(out ExcelChartSnapshot snapshot)
                        || chart.DataRange is not ExcelChartDataRange range
                        || !chart.HasCanonicalWorksheetReferences()
                        || !range.HasHeaderRow
                        || range.CategoryCount < 1 || range.CategoryCount > MaximumConvertedChartPoints
                        || range.SeriesCount < 1 || range.SeriesCount > MaximumConvertedChartSeries
                        || snapshot.Data.Categories.Count != range.CategoryCount
                        || snapshot.Data.Series.Count != range.SeriesCount
                        || snapshot.Data.Series.Any(series =>
                            series.ChartType.HasValue && series.ChartType.Value != chart.ChartType
                            || series.AxisGroup != OfficeChartAxisGroup.Primary)
                        || snapshot.RowIndex < 1 || snapshot.ColumnIndex < 1
                        || snapshot.RowIndex > options.MaximumRows || snapshot.ColumnIndex > options.MaximumColumns
                        || snapshot.OffsetXPixels < 0 || snapshot.OffsetYPixels < 0
                        || snapshot.WidthPixels < 1 || snapshot.HeightPixels < 1
                        || snapshot.WidthPixels > 2000 || snapshot.HeightPixels > 2000
                        || snapshot.Title?.Length > 32767
                        || source.Sheets.FirstOrDefault(sheet => string.Equals(sheet.Name, range.SheetName,
                            StringComparison.OrdinalIgnoreCase)) is not ExcelSheet dataSourceSheet
                        || !convertedCellsBySheet.TryGetValue(dataSourceSheet.Name,
                            out HashSet<(int Row, int Column)>? dataCells)
                        || target.GetSheet(dataSourceSheet.Name) is not OdsSheet dataSheet
                        || !HasAllChartSourceCells(range, dataCells, dataSheet, snapshot.Data)) continue;
                    var series = new OdsChartSeries[range.SeriesCount];
                    bool valid = true;
                    for (int index = 0; index < series.Length; index++) {
                        if (snapshot.Data.Series[index].Values.Count != range.CategoryCount ||
                            snapshot.Data.Series[index].Values.Any(number => double.IsNaN(number) || double.IsInfinity(number))) {
                            valid = false;
                            break;
                        }
                        string values = SpreadsheetAddressConverter.ExcelRangeToOpenAddress(
                            range.SeriesValuesRangeA1(index), dataSourceSheet.Name);
                        string label = SpreadsheetAddressConverter.ExcelRangeToOpenAddress(
                            range.SeriesNameCellA1(index), dataSourceSheet.Name);
                        if (values.Length == 0 || label.Length == 0) { valid = false; break; }
                        series[index] = new OdsChartSeries(values, label);
                    }
                    if (!valid) continue;
                    string categories = SpreadsheetAddressConverter.ExcelRangeToOpenAddress(
                        range.CategoriesRangeA1, dataSourceSheet.Name);
                    if (categories.Length == 0) continue;
                    var coordinate = (snapshot.RowIndex, snapshot.ColumnIndex);
                    if (!anchorCells.Contains(coordinate) && materializedCells >= options.MaximumExpandedCells) {
                        truncated = true;
                        continue;
                    }
                    if (targetSheet.Cell(snapshot.RowIndex - 1L, snapshot.ColumnIndex - 1L).IsCovered)
                        continue;
                    targetSheet.AddChart(type, categories, series,
                        snapshot.RowIndex - 1L, snapshot.ColumnIndex - 1L,
                        new OdfRect(OdfLength.Points(snapshot.OffsetXPixels * 0.75D),
                            OdfLength.Points(snapshot.OffsetYPixels * 0.75D),
                            OdfLength.Points(snapshot.WidthPixels * 0.75D),
                            OdfLength.Points(snapshot.HeightPixels * 0.75D)),
                        snapshot.Title, snapshot.Name);
                    if (anchorCells.Add(coordinate)) materializedCells++;
                    converted++;
                } catch (Exception exception) when (exception is InvalidOperationException or
                    InvalidCastException or ArgumentException or KeyNotFoundException or InvalidDataException or
                    OpenXmlPackageException or System.Xml.XmlException) {
                    // A malformed drawing relationship or chart part is an unsupported chart,
                    // not a failure of the whole workbook conversion.
                }
            }
        }
        return converted;
    }

    private static int CountReferencedChartParts(ExcelDocument source) {
        var referenced = new HashSet<Uri>();
        foreach (WorksheetPart worksheet in source.OpenXmlDocument.WorkbookPart?.WorksheetParts
                     ?? Enumerable.Empty<WorksheetPart>()) {
            DrawingsPart? drawing = worksheet.DrawingsPart;
            if (drawing?.WorksheetDrawing == null) continue;
            foreach (Xdr.GraphicFrame frame in drawing.WorksheetDrawing.Descendants<Xdr.GraphicFrame>()) {
                string? relationshipId = frame.Graphic?.GraphicData?.GetFirstChild<C.ChartReference>()?.Id?.Value;
                if (string.IsNullOrWhiteSpace(relationshipId)) continue;
                try {
                    if (drawing.GetPartById(relationshipId!) is ChartPart part) referenced.Add(part.Uri);
                } catch (Exception exception) when (exception is ArgumentException or InvalidOperationException or
                    OpenXmlPackageException) {
                    // The frame is counted separately even when its relationship is missing.
                }
            }
        }
        return referenced.Count;
    }

    private static bool TryGetOdsChartType(ExcelChartType source, out OdsChartType type) {
        switch (source) {
            case ExcelChartType.ColumnClustered: type = OdsChartType.Column; return true;
            case ExcelChartType.BarClustered: type = OdsChartType.Bar; return true;
            case ExcelChartType.Line: type = OdsChartType.Line; return true;
            default: type = default; return false;
        }
    }

    private static bool HasAllChartSourceCells(ExcelChartDataRange range,
        HashSet<(int Row, int Column)> converted, OdsSheet dataSheet, ExcelChartData data) {
        var reader = new ChartCellReader(dataSheet);
        if (range.HasHeaderRow) {
            for (int series = 0; series < range.SeriesCount; series++) {
                int row = range.Orientation == ExcelChartDataOrientation.Vertical
                    ? range.StartRow : range.SeriesStartRow + series;
                int column = range.Orientation == ExcelChartDataOrientation.Vertical
                    ? range.SeriesStartColumn + series : range.StartColumn;
                if (!converted.Contains((row, column))) return false;
            }
        }
        for (int point = 0; point < range.CategoryCount; point++) {
            int categoryRow = range.Orientation == ExcelChartDataOrientation.Vertical
                ? range.CategoryStartRow + point : range.CategoryStartRow;
            int categoryColumn = range.Orientation == ExcelChartDataOrientation.Vertical
                ? range.CategoryStartColumn : range.CategoryStartColumn + point;
            if (!converted.Contains((categoryRow, categoryColumn))) return false;
            for (int series = 0; series < range.SeriesCount; series++) {
                int valueRow = range.Orientation == ExcelChartDataOrientation.Vertical
                    ? range.SeriesStartRow + point : range.SeriesStartRow + series;
                int valueColumn = range.Orientation == ExcelChartDataOrientation.Vertical
                    ? range.SeriesStartColumn + series : range.SeriesStartColumn + point;
                if (!converted.Contains((valueRow, valueColumn))) return false;
                OdsCellValue value = reader.GetValue(valueRow - 1L, valueColumn - 1);
                if (value.Kind != OdsCellValueKind.Number
                    || !double.TryParse(value.LexicalValue, NumberStyles.Float,
                        CultureInfo.InvariantCulture, out double number)
                    || double.IsNaN(number) || double.IsInfinity(number)
                    || number != data.Series[series].Values[point]) return false;
            }
        }
        return data.Series.All(series => !string.IsNullOrWhiteSpace(series.Name));
    }
}
