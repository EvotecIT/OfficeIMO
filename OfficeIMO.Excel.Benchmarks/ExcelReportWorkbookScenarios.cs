using System.Data;
using System.Globalization;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeIMO.Excel.Benchmarks;

internal static partial class ExcelLibraryComparisonRunner {
    private const string ReportWorkbookScenario = "report-workbook";
    private const string ReportWorkbookCoreScenario = "report-workbook-core";
    private const string ReportWorkbookDataTableScenario = "report-workbook-datatable";
    private const string ReportWorkbookDataTableCoreScenario = "report-workbook-datatable-core";
    private const string ReportWorkbookTableName = "Data";

    private static void AddReportWorkbookScenarioGroups(
        List<ExcelLibraryComparisonScenario> scenarios,
        IReadOnlySet<string>? scenarioFilter,
        IReadOnlyList<Dictionary<string, object?>> rows,
        DataTable dataTable,
        int warmupIterations,
        int measuredIterations) {
        AddScenarioGroup(scenarios, scenarioFilter, ReportWorkbookScenario, warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Create the PSWriteOffice report workbook shape with table, AutoFit, freeze top row, conditional formatting, list validation, number formats, clustered column chart, pivot table, and save.", () => OfficeImoWriteReportWorkbook(rows)),
            new LibraryComparisonCase("EPPlus", "Create the equivalent report workbook shape from the same mixed object rows with table, AutoFit, freeze top row, conditional formatting, list validation, number formats, clustered column chart, pivot table, and save.", () => EpPlusWriteReportWorkbook(rows))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, ReportWorkbookCoreScenario, warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Create the report workbook table/core formatting shape without chart or pivot table.", () => OfficeImoWriteReportWorkbookCore(rows)),
            new LibraryComparisonCase("ClosedXML", "Create the report workbook table/core formatting shape from the same mixed object rows without chart or pivot table.", () => ClosedXmlWriteReportWorkbookCore(rows)),
            new LibraryComparisonCase("EPPlus", "Create the report workbook table/core formatting shape from the same mixed object rows without chart or pivot table.", () => EpPlusWriteReportWorkbookCore(rows))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, ReportWorkbookDataTableScenario, warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Create the report workbook shape from the same typed DataTable with table, AutoFit, freeze top row, conditional formatting, list validation, number formats, clustered column chart, pivot table, and save.", () => OfficeImoWriteReportWorkbookDataTable(dataTable)),
            new LibraryComparisonCase("EPPlus", "Create the equivalent report workbook shape from the same typed DataTable with table, AutoFit, freeze top row, conditional formatting, list validation, number formats, clustered column chart, pivot table, and save.", () => EpPlusWriteReportWorkbookDataTable(dataTable))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, ReportWorkbookDataTableCoreScenario, warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Create the report workbook table/core formatting shape from the same typed DataTable without chart or pivot table.", () => OfficeImoWriteReportWorkbookDataTableCore(dataTable)),
            new LibraryComparisonCase("ClosedXML", "Create the report workbook table/core formatting shape from the same typed DataTable without chart or pivot table.", () => ClosedXmlWriteReportWorkbookDataTableCore(dataTable)),
            new LibraryComparisonCase("EPPlus", "Create the report workbook table/core formatting shape from the same typed DataTable without chart or pivot table.", () => EpPlusWriteReportWorkbookDataTableCore(dataTable))
        ]);
    }

    private static void AddReportWorkbookPackageProfileGroups(
        List<ExcelPackageProfileScenario> scenarios,
        IReadOnlySet<string>? scenarioFilter,
        IReadOnlyList<Dictionary<string, object?>> rows,
        DataTable dataTable,
        int warmupIterations,
        int measuredIterations) {
        AddPackageProfileGroup(scenarios, scenarioFilter, ReportWorkbookScenario, warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Create the PSWriteOffice report workbook shape with table, AutoFit, freeze top row, conditional formatting, list validation, number formats, clustered column chart, pivot table, and save.", () => OfficeImoWriteReportWorkbookBytes(rows)),
            new PackageProfileCase("EPPlus", "Create the equivalent report workbook shape from the same mixed object rows with table, AutoFit, freeze top row, conditional formatting, list validation, number formats, clustered column chart, pivot table, and save.", () => EpPlusWriteReportWorkbookBytes(rows))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, ReportWorkbookCoreScenario, warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Create the report workbook table/core formatting shape without chart or pivot table.", () => OfficeImoWriteReportWorkbookCoreBytes(rows)),
            new PackageProfileCase("ClosedXML", "Create the report workbook table/core formatting shape from the same mixed object rows without chart or pivot table.", () => ClosedXmlWriteReportWorkbookCoreBytes(rows)),
            new PackageProfileCase("EPPlus", "Create the report workbook table/core formatting shape from the same mixed object rows without chart or pivot table.", () => EpPlusWriteReportWorkbookCoreBytes(rows))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, ReportWorkbookDataTableScenario, warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Create the report workbook shape from the same typed DataTable with table, AutoFit, freeze top row, conditional formatting, list validation, number formats, clustered column chart, pivot table, and save.", () => OfficeImoWriteReportWorkbookDataTableBytes(dataTable)),
            new PackageProfileCase("EPPlus", "Create the equivalent report workbook shape from the same typed DataTable with table, AutoFit, freeze top row, conditional formatting, list validation, number formats, clustered column chart, pivot table, and save.", () => EpPlusWriteReportWorkbookDataTableBytes(dataTable))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, ReportWorkbookDataTableCoreScenario, warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Create the report workbook table/core formatting shape from the same typed DataTable without chart or pivot table.", () => OfficeImoWriteReportWorkbookDataTableCoreBytes(dataTable)),
            new PackageProfileCase("ClosedXML", "Create the report workbook table/core formatting shape from the same typed DataTable without chart or pivot table.", () => ClosedXmlWriteReportWorkbookDataTableCoreBytes(dataTable)),
            new PackageProfileCase("EPPlus", "Create the report workbook table/core formatting shape from the same typed DataTable without chart or pivot table.", () => EpPlusWriteReportWorkbookDataTableCoreBytes(dataTable))
        ]);
    }

    private static int OfficeImoWriteReportWorkbook(IReadOnlyList<Dictionary<string, object?>> rows)
        => ByteCount(OfficeImoWriteReportWorkbookBytes(rows));

    private static byte[] OfficeImoWriteReportWorkbookBytes(IReadOnlyList<Dictionary<string, object?>> rows) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorksheet("Data");
            sheet.InsertObjects(rows);
            ApplyOfficeImoReportWorkbookCore(sheet, rows.Count);
            AddOfficeImoReportWorkbookChart(sheet, rows.Count);
            AddOfficeImoReportWorkbookPivotTable(sheet, rows.Count);
            document.Save(stream);
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteReportWorkbookCore(IReadOnlyList<Dictionary<string, object?>> rows)
        => ByteCount(OfficeImoWriteReportWorkbookCoreBytes(rows));

    private static byte[] OfficeImoWriteReportWorkbookCoreBytes(IReadOnlyList<Dictionary<string, object?>> rows) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorksheet("Data");
            sheet.InsertObjects(rows);
            ApplyOfficeImoReportWorkbookCore(sheet, rows.Count);
            document.Save(stream);
            AssertOfficeImoDirectPackageWriter(document, "report workbook core");
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteReportWorkbookDataTable(DataTable dataTable)
        => ByteCount(OfficeImoWriteReportWorkbookDataTableBytes(dataTable));

    private static byte[] OfficeImoWriteReportWorkbookDataTableBytes(DataTable dataTable) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorksheet("Data");
            sheet.InsertDataTable(dataTable);
            ApplyOfficeImoReportWorkbookCore(sheet, dataTable.Rows.Count);
            AddOfficeImoReportWorkbookChart(sheet, dataTable.Rows.Count);
            AddOfficeImoReportWorkbookPivotTable(sheet, dataTable.Rows.Count);
            document.Save(stream);
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteReportWorkbookDataTableCore(DataTable dataTable)
        => ByteCount(OfficeImoWriteReportWorkbookDataTableCoreBytes(dataTable));

    private static byte[] OfficeImoWriteReportWorkbookDataTableCoreBytes(DataTable dataTable) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorksheet("Data");
            sheet.InsertDataTable(dataTable);
            ApplyOfficeImoReportWorkbookCore(sheet, dataTable.Rows.Count);
            document.Save(stream);
            AssertOfficeImoDirectPackageWriter(document, "report workbook DataTable core");
        }

        return stream.ToArray();
    }

    private static int ClosedXmlWriteReportWorkbookCore(IReadOnlyList<Dictionary<string, object?>> rows)
        => ByteCount(ClosedXmlWriteReportWorkbookCoreBytes(rows));

    private static byte[] ClosedXmlWriteReportWorkbookCoreBytes(IReadOnlyList<Dictionary<string, object?>> rows) {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook()) {
            var worksheet = workbook.Worksheets.Add("Data");
            WriteClosedXmlReportWorkbookRows(worksheet, rows);
            var table = worksheet.Range(1, 1, rows.Count + 1, PowerShellMixedColumnNames.Length).CreateTable(ReportWorkbookTableName);
            ExcelBenchmarkScenarioFactory.StyleClosedXmlTable(table);
            worksheet.ColumnsUsed().AdjustToContents();
            worksheet.SheetView.FreezeRows(1);
            ApplyClosedXmlReportWorkbookCore(worksheet, rows.Count);
            workbook.SaveAs(stream);
        }

        return stream.ToArray();
    }

    private static int ClosedXmlWriteReportWorkbookDataTableCore(DataTable dataTable)
        => ByteCount(ClosedXmlWriteReportWorkbookDataTableCoreBytes(dataTable));

    private static byte[] ClosedXmlWriteReportWorkbookDataTableCoreBytes(DataTable dataTable) {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook()) {
            var worksheet = workbook.Worksheets.Add("Data");
            var table = worksheet.Cell(1, 1).InsertTable(dataTable, ReportWorkbookTableName, true);
            ExcelBenchmarkScenarioFactory.StyleClosedXmlTable(table);
            worksheet.ColumnsUsed().AdjustToContents();
            worksheet.SheetView.FreezeRows(1);
            ApplyClosedXmlReportWorkbookCore(worksheet, dataTable.Rows.Count);
            workbook.SaveAs(stream);
        }

        return stream.ToArray();
    }

    private static int EpPlusWriteReportWorkbook(IReadOnlyList<Dictionary<string, object?>> rows)
        => ByteCount(EpPlusWriteReportWorkbookBytes(rows));

    private static byte[] EpPlusWriteReportWorkbookBytes(IReadOnlyList<Dictionary<string, object?>> rows) {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage(stream)) {
            var worksheet = package.Workbook.Worksheets.Add("Data");
            PopulateEpPlusReportWorkbookData(worksheet, rows);
            ApplyEpPlusReportWorkbookCore(worksheet, rows.Count);
            AddEpPlusReportWorkbookChart(worksheet, rows.Count);
            AddEpPlusReportWorkbookPivotTable(package, worksheet, rows.Count);
            package.Save();
        }

        return stream.ToArray();
    }

    private static int EpPlusWriteReportWorkbookCore(IReadOnlyList<Dictionary<string, object?>> rows)
        => ByteCount(EpPlusWriteReportWorkbookCoreBytes(rows));

    private static byte[] EpPlusWriteReportWorkbookCoreBytes(IReadOnlyList<Dictionary<string, object?>> rows) {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage(stream)) {
            var worksheet = package.Workbook.Worksheets.Add("Data");
            PopulateEpPlusReportWorkbookData(worksheet, rows);
            ApplyEpPlusReportWorkbookCore(worksheet, rows.Count);
            package.Save();
        }

        return stream.ToArray();
    }

    private static int EpPlusWriteReportWorkbookDataTable(DataTable dataTable)
        => ByteCount(EpPlusWriteReportWorkbookDataTableBytes(dataTable));

    private static byte[] EpPlusWriteReportWorkbookDataTableBytes(DataTable dataTable) {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage(stream)) {
            var worksheet = package.Workbook.Worksheets.Add("Data");
            PopulateEpPlusReportWorkbookData(worksheet, dataTable);
            ApplyEpPlusReportWorkbookCore(worksheet, dataTable.Rows.Count);
            AddEpPlusReportWorkbookChart(worksheet, dataTable.Rows.Count);
            AddEpPlusReportWorkbookPivotTable(package, worksheet, dataTable.Rows.Count);
            package.Save();
        }

        return stream.ToArray();
    }

    private static int EpPlusWriteReportWorkbookDataTableCore(DataTable dataTable)
        => ByteCount(EpPlusWriteReportWorkbookDataTableCoreBytes(dataTable));

    private static byte[] EpPlusWriteReportWorkbookDataTableCoreBytes(DataTable dataTable) {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage(stream)) {
            var worksheet = package.Workbook.Worksheets.Add("Data");
            PopulateEpPlusReportWorkbookData(worksheet, dataTable);
            ApplyEpPlusReportWorkbookCore(worksheet, dataTable.Rows.Count);
            package.Save();
        }

        return stream.ToArray();
    }

    private static void ApplyOfficeImoReportWorkbookCore(ExcelSheet sheet, int rowCount) {
        int lastRow = rowCount + 1;
        string range = BuildReportWorkbookRange(rowCount);
        sheet.AddTable(range, hasHeader: true, name: ReportWorkbookTableName, style: OfficeIMO.Excel.TableStyle.TableStyleMedium2, includeAutoFilter: true);
        sheet.AutoFitColumns();
        sheet.Freeze(topRows: 1);
        sheet.AddConditionalRule($"G2:G{lastRow.ToString(CultureInfo.InvariantCulture)}", ConditionalFormattingOperatorValues.GreaterThan, "700");
        sheet.AddConditionalDataBar($"G2:G{lastRow.ToString(CultureInfo.InvariantCulture)}", OfficeColor.SteelBlue);
        sheet.AddConditionalColorScale($"I2:I{lastRow.ToString(CultureInfo.InvariantCulture)}", OfficeColor.LightPink, OfficeColor.LightGreen);
        sheet.AddConditionalIconSet($"I2:I{lastRow.ToString(CultureInfo.InvariantCulture)}", IconSetValues.ThreeTrafficLights1, showValue: true, reverseIconOrder: false);
        sheet.ValidationList($"D2:D{lastRow.ToString(CultureInfo.InvariantCulture)}", new[] { "NA", "EU", "APAC", "LATAM" });
        sheet.ColumnStyleByHeader("Score").NumberFormat("#,##0.000");
        sheet.ColumnStyleByHeader("Created").DateTime("yyyy-mm-dd hh:mm");
    }

    private static void AddOfficeImoReportWorkbookChart(ExcelSheet sheet, int rowCount) {
        int lastRow = rowCount + 1;
        sheet.AddChartFromRange($"F1:G{lastRow.ToString(CultureInfo.InvariantCulture)}", row: 2, column: 12, widthPixels: 720, heightPixels: 320, type: ExcelChartType.ColumnClustered, title: "Score by Created", includeCachedData: false);
    }

    private static void AddOfficeImoReportWorkbookPivotTable(ExcelSheet sheet, int rowCount) {
        sheet.AddPivotTable(
            sourceRange: BuildReportWorkbookRange(rowCount),
            destinationCell: "L24",
            name: "ReportPivot",
            rowFields: new[] { "Region" },
            columnFields: new[] { "Department" },
            dataFields: new[] {
                new ExcelPivotDataField("Score", DataConsolidateFunctionValues.Average, "Average Score"),
                new ExcelPivotDataField("TicketCount", DataConsolidateFunctionValues.Sum, "Sum TicketCount")
            },
            pivotStyleName: "PivotStyleMedium9");
    }

    private static void PopulateEpPlusReportWorkbookData(ExcelWorksheet worksheet, IReadOnlyList<Dictionary<string, object?>> rows) {
        WriteEpPlusReportWorkbookRows(worksheet, rows);
        var table = worksheet.Tables.Add(worksheet.Cells[1, 1, rows.Count + 1, PowerShellMixedColumnNames.Length], ReportWorkbookTableName);
        table.TableStyle = OfficeOpenXml.Table.TableStyles.Medium2;
        if (worksheet.Dimension != null) {
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
        }
    }

    private static void PopulateEpPlusReportWorkbookData(ExcelWorksheet worksheet, DataTable dataTable) {
        worksheet.Cells["A1"].LoadFromDataTable(dataTable, true);
        var table = worksheet.Tables.Add(worksheet.Cells[1, 1, dataTable.Rows.Count + 1, dataTable.Columns.Count], ReportWorkbookTableName);
        table.TableStyle = OfficeOpenXml.Table.TableStyles.Medium2;
        if (worksheet.Dimension != null) {
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
        }
    }

    private static void ApplyEpPlusReportWorkbookCore(ExcelWorksheet worksheet, int rowCount) {
        int lastRow = rowCount + 1;
        worksheet.View.FreezePanes(2, 1);
        var highScore = worksheet.ConditionalFormatting.AddGreaterThan(worksheet.Cells[2, 7, lastRow, 7]);
        highScore.Formula = "700";
        highScore.Style.Fill.PatternType = ExcelFillStyle.Solid;
        highScore.Style.Fill.BackgroundColor.Color = System.Drawing.Color.LightGreen;
        worksheet.ConditionalFormatting.AddDatabar(worksheet.Cells[2, 7, lastRow, 7], System.Drawing.Color.SteelBlue);
        worksheet.ConditionalFormatting.AddTwoColorScale(worksheet.Cells[2, 9, lastRow, 9]);
        worksheet.ConditionalFormatting.AddThreeIconSet(worksheet.Cells[2, 9, lastRow, 9], OfficeOpenXml.ConditionalFormatting.eExcelconditionalFormatting3IconsSetType.TrafficLights1);
        var validation = worksheet.DataValidations.AddListValidation($"D2:D{lastRow.ToString(CultureInfo.InvariantCulture)}");
        validation.Formula.Values.Add("NA");
        validation.Formula.Values.Add("EU");
        validation.Formula.Values.Add("APAC");
        validation.Formula.Values.Add("LATAM");
        worksheet.Cells[2, 7, lastRow, 7].Style.Numberformat.Format = "#,##0.000";
        worksheet.Cells[2, 6, lastRow, 6].Style.Numberformat.Format = "yyyy-mm-dd hh:mm";
    }

    private static void AddEpPlusReportWorkbookChart(ExcelWorksheet worksheet, int rowCount) {
        int lastRow = rowCount + 1;
        var chart = worksheet.Drawings.AddChart("ReportScoreChart", eChartType.ColumnClustered);
        chart.Title.Text = "Score by Created";
        chart.SetPosition(1, 0, 11, 0);
        chart.SetSize(720, 320);
        chart.Series.Add(worksheet.Cells[2, 7, lastRow, 7], worksheet.Cells[2, 6, lastRow, 6]);
    }

    private static void AddEpPlusReportWorkbookPivotTable(ExcelPackage package, ExcelWorksheet dataWorksheet, int rowCount) {
        var source = dataWorksheet.Cells[1, 1, rowCount + 1, PowerShellMixedColumnNames.Length];
        var pivot = dataWorksheet.PivotTables.Add(dataWorksheet.Cells["L24"], source, "ReportPivot");
        pivot.RowFields.Add(pivot.Fields["Region"]);
        pivot.ColumnFields.Add(pivot.Fields["Department"]);
        var score = pivot.DataFields.Add(pivot.Fields["Score"]);
        score.Function = DataFieldFunctions.Average;
        score.Name = "Average Score";
        var tickets = pivot.DataFields.Add(pivot.Fields["TicketCount"]);
        tickets.Function = DataFieldFunctions.Sum;
        tickets.Name = "Sum TicketCount";
    }

    private static void ApplyClosedXmlReportWorkbookCore(IXLWorksheet worksheet, int rowCount) {
        int lastRow = rowCount + 1;
        worksheet.Range(2, 7, lastRow, 7).AddConditionalFormat().WhenGreaterThan(700).Fill.SetBackgroundColor(XLColor.LightGreen);
        worksheet.Range(2, 4, lastRow, 4).CreateDataValidation().List("\"NA,EU,APAC,LATAM\"");
        worksheet.Range(2, 7, lastRow, 7).Style.NumberFormat.Format = "#,##0.000";
        worksheet.Range(2, 6, lastRow, 6).Style.DateFormat.Format = "yyyy-mm-dd hh:mm";
    }

    private static void WriteClosedXmlReportWorkbookRows(IXLWorksheet worksheet, IReadOnlyList<Dictionary<string, object?>> rows) {
        for (int i = 0; i < PowerShellMixedColumnNames.Length; i++) {
            worksheet.Cell(1, i + 1).Value = PowerShellMixedColumnNames[i];
        }

        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            Dictionary<string, object?> row = rows[rowIndex];
            int targetRow = rowIndex + 2;
            for (int columnIndex = 0; columnIndex < PowerShellMixedColumnNames.Length; columnIndex++) {
                row.TryGetValue(PowerShellMixedColumnNames[columnIndex], out object? value);
                worksheet.Cell(targetRow, columnIndex + 1).Value = XLCellValue.FromObject(value);
            }
        }
    }

    private static void WriteEpPlusReportWorkbookRows(ExcelWorksheet worksheet, IReadOnlyList<Dictionary<string, object?>> rows) {
        for (int i = 0; i < PowerShellMixedColumnNames.Length; i++) {
            worksheet.Cells[1, i + 1].Value = PowerShellMixedColumnNames[i];
        }

        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            Dictionary<string, object?> row = rows[rowIndex];
            int targetRow = rowIndex + 2;
            for (int columnIndex = 0; columnIndex < PowerShellMixedColumnNames.Length; columnIndex++) {
                row.TryGetValue(PowerShellMixedColumnNames[columnIndex], out object? value);
                worksheet.Cells[targetRow, columnIndex + 1].Value = value;
            }
        }
    }

    private static string BuildReportWorkbookRange(int rowCount)
        => "A1:J" + (rowCount + 1).ToString(CultureInfo.InvariantCulture);
}
