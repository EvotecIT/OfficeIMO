using System.Globalization;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeIMO.Excel.Benchmarks;

internal static partial class ExcelLibraryComparisonRunner {
    private const string RealWorldReportScenario = "realworld-report-all-in-one";
    private const string RealWorldReportCoreScenario = "realworld-report-core";
    private const string RealWorldFreezePanesScenario = "realworld-freeze-panes";
    private const string RealWorldAutoFilterScenario = "realworld-autofilter";
    private const string RealWorldConditionalFormattingScenario = "realworld-conditional-formatting";
    private const string RealWorldDataValidationScenario = "realworld-data-validation";
    private const string RealWorldChartsScenario = "realworld-charts";
    private const string RealWorldPivotTableScenario = "realworld-pivot-table";
    private const string RealWorldReportNoAutoFitScenario = "realworld-report-no-autofit";
    private const string RealWorldReportChartFirstScenario = "realworld-report-chart-first";
    private const string RealWorldReportShuffledColumnsScenario = "realworld-report-shuffled-columns";
    private const string RealWorldReportExtraColumnScenario = "realworld-report-extra-column";
    private const string RealWorldReportPostMutationScenario = "realworld-report-post-mutation";

    private static void AddRealWorldScenarioGroups(
        List<ExcelLibraryComparisonScenario> scenarios,
        IReadOnlySet<string>? scenarioFilter,
        IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows,
        int warmupIterations,
        int measuredIterations) {
        AddScenarioGroup(scenarios, scenarioFilter, RealWorldReportScenario, warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Create a sales workbook with table, AutoFit, freeze panes, filters, conditional formatting, pivot table, chart, and save.", () => OfficeImoWriteRealWorldReport(rows)),
            new LibraryComparisonCase("EPPlus", "Create a sales workbook with table, AutoFit, freeze panes, filters, conditional formatting, pivot table, chart, and save.", () => EpPlusWriteRealWorldReport(rows))
        ]);

        AddRealWorldVariantScenarioGroups(scenarios, scenarioFilter, rows, warmupIterations, measuredIterations);

        AddScenarioGroup(scenarios, scenarioFilter, RealWorldReportCoreScenario, warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Create a sales workbook with table, AutoFit, frozen header, AutoFilter, conditional formatting, data validation, and save.", () => OfficeImoWriteRealWorldCoreReport(rows)),
            new LibraryComparisonCase("ClosedXML", "Create a sales workbook with table, AutoFit, frozen header, AutoFilter, conditional formatting, data validation, and save.", () => ClosedXmlWriteRealWorldCoreReport(rows)),
            new LibraryComparisonCase("EPPlus", "Create a sales workbook with table, AutoFit, frozen header, AutoFilter, conditional formatting, data validation, and save.", () => EpPlusWriteRealWorldCoreReport(rows))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, RealWorldFreezePanesScenario, warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Write a sales table, freeze the header row and first column, and save.", () => OfficeImoWriteRealWorldFreezePanes(rows)),
            new LibraryComparisonCase("ClosedXML", "Write a sales table, freeze the header row and first column, and save.", () => ClosedXmlWriteRealWorldFreezePanes(rows)),
            new LibraryComparisonCase("EPPlus", "Write a sales table, freeze the header row and first column, and save.", () => EpPlusWriteRealWorldFreezePanes(rows))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, RealWorldAutoFilterScenario, warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Write a sales table, add worksheet-level AutoFilter, and save.", () => OfficeImoWriteRealWorldAutoFilter(rows)),
            new LibraryComparisonCase("ClosedXML", "Write a sales table, add worksheet-level AutoFilter, and save.", () => ClosedXmlWriteRealWorldAutoFilter(rows)),
            new LibraryComparisonCase("EPPlus", "Write a sales table, add worksheet-level AutoFilter, and save.", () => EpPlusWriteRealWorldAutoFilter(rows))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, RealWorldConditionalFormattingScenario, warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Write a sales table, add value rules plus visual conditional formatting, and save.", () => OfficeImoWriteRealWorldConditionalFormatting(rows)),
            new LibraryComparisonCase("ClosedXML", "Write a sales table, add equivalent value rules, and save.", () => ClosedXmlWriteRealWorldConditionalFormatting(rows)),
            new LibraryComparisonCase("EPPlus", "Write a sales table, add equivalent value rules, and save.", () => EpPlusWriteRealWorldConditionalFormatting(rows))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, RealWorldDataValidationScenario, warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Write a sales table, add whole-number data validation to the Units column, and save.", () => OfficeImoWriteRealWorldDataValidation(rows)),
            new LibraryComparisonCase("ClosedXML", "Write a sales table, add whole-number data validation to the Units column, and save.", () => ClosedXmlWriteRealWorldDataValidation(rows)),
            new LibraryComparisonCase("EPPlus", "Write a sales table, add whole-number data validation to the Units column, and save.", () => EpPlusWriteRealWorldDataValidation(rows))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, RealWorldChartsScenario, warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Write sales data, add a clustered column chart over regional totals, and save.", () => OfficeImoWriteRealWorldCharts(rows)),
            new LibraryComparisonCase("EPPlus", "Write sales data, add a clustered column chart over regional totals, and save.", () => EpPlusWriteRealWorldCharts(rows))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, RealWorldPivotTableScenario, warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Write sales data, add a pivot table with row, column, and sum data fields, and save.", () => OfficeImoWriteRealWorldPivotTable(rows)),
            new LibraryComparisonCase("EPPlus", "Write sales data, add a pivot table with row, column, and sum data fields, and save.", () => EpPlusWriteRealWorldPivotTable(rows))
        ]);
    }

    private static void AddRealWorldPackageProfileGroups(
        List<ExcelPackageProfileScenario> scenarios,
        IReadOnlySet<string>? scenarioFilter,
        IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows,
        int warmupIterations,
        int measuredIterations) {
        AddPackageProfileGroup(scenarios, scenarioFilter, RealWorldReportScenario, warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Create a sales workbook with table, AutoFit, freeze panes, filters, conditional formatting, pivot table, chart, and save.", () => OfficeImoWriteRealWorldReportBytes(rows)),
            new PackageProfileCase("EPPlus", "Create a sales workbook with table, AutoFit, freeze panes, filters, conditional formatting, pivot table, chart, and save.", () => EpPlusWriteRealWorldReportBytes(rows))
        ]);

        AddRealWorldVariantPackageProfileGroups(scenarios, scenarioFilter, rows, warmupIterations, measuredIterations);

        AddPackageProfileGroup(scenarios, scenarioFilter, RealWorldReportCoreScenario, warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Create a sales workbook with table, AutoFit, frozen header, AutoFilter, conditional formatting, data validation, and save.", () => OfficeImoWriteRealWorldCoreReportBytes(rows)),
            new PackageProfileCase("ClosedXML", "Create a sales workbook with table, AutoFit, frozen header, AutoFilter, conditional formatting, data validation, and save.", () => ClosedXmlWriteRealWorldCoreReportBytes(rows)),
            new PackageProfileCase("EPPlus", "Create a sales workbook with table, AutoFit, frozen header, AutoFilter, conditional formatting, data validation, and save.", () => EpPlusWriteRealWorldCoreReportBytes(rows))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, RealWorldFreezePanesScenario, warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Write a sales table, freeze the header row and first column, and save.", () => OfficeImoWriteRealWorldFreezePanesBytes(rows)),
            new PackageProfileCase("ClosedXML", "Write a sales table, freeze the header row and first column, and save.", () => ClosedXmlWriteRealWorldFreezePanesBytes(rows)),
            new PackageProfileCase("EPPlus", "Write a sales table, freeze the header row and first column, and save.", () => EpPlusWriteRealWorldFreezePanesBytes(rows))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, RealWorldAutoFilterScenario, warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Write a sales table, add worksheet-level AutoFilter, and save.", () => OfficeImoWriteRealWorldAutoFilterBytes(rows)),
            new PackageProfileCase("ClosedXML", "Write a sales table, add worksheet-level AutoFilter, and save.", () => ClosedXmlWriteRealWorldAutoFilterBytes(rows)),
            new PackageProfileCase("EPPlus", "Write a sales table, add worksheet-level AutoFilter, and save.", () => EpPlusWriteRealWorldAutoFilterBytes(rows))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, RealWorldConditionalFormattingScenario, warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Write a sales table, add value rules plus visual conditional formatting, and save.", () => OfficeImoWriteRealWorldConditionalFormattingBytes(rows)),
            new PackageProfileCase("ClosedXML", "Write a sales table, add equivalent value rules, and save.", () => ClosedXmlWriteRealWorldConditionalFormattingBytes(rows)),
            new PackageProfileCase("EPPlus", "Write a sales table, add equivalent value rules, and save.", () => EpPlusWriteRealWorldConditionalFormattingBytes(rows))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, RealWorldDataValidationScenario, warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Write a sales table, add whole-number data validation to the Units column, and save.", () => OfficeImoWriteRealWorldDataValidationBytes(rows)),
            new PackageProfileCase("ClosedXML", "Write a sales table, add whole-number data validation to the Units column, and save.", () => ClosedXmlWriteRealWorldDataValidationBytes(rows)),
            new PackageProfileCase("EPPlus", "Write a sales table, add whole-number data validation to the Units column, and save.", () => EpPlusWriteRealWorldDataValidationBytes(rows))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, RealWorldChartsScenario, warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Write sales data, add a clustered column chart over regional totals, and save.", () => OfficeImoWriteRealWorldChartsBytes(rows)),
            new PackageProfileCase("EPPlus", "Write sales data, add a clustered column chart over regional totals, and save.", () => EpPlusWriteRealWorldChartsBytes(rows))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, RealWorldPivotTableScenario, warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Write sales data, add a pivot table with row, column, and sum data fields, and save.", () => OfficeImoWriteRealWorldPivotTableBytes(rows)),
            new PackageProfileCase("EPPlus", "Write sales data, add a pivot table with row, column, and sum data fields, and save.", () => EpPlusWriteRealWorldPivotTableBytes(rows))
        ]);
    }

    private static int OfficeImoWriteRealWorldReport(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(OfficeImoWriteRealWorldReportBytes(rows));

    private static byte[] OfficeImoWriteRealWorldReportBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream, autoSave: false)) {
            var sheet = document.AddWorkSheet("Data");
            ExcelBenchmarkScenarioFactory.InsertOfficeImoObjects(sheet, rows);
            ApplyOfficeImoTable(sheet, rows.Count);
            ApplyOfficeImoNavigation(sheet, rows.Count);
            ApplyOfficeImoConditionalFormatting(sheet, rows.Count);
            ApplyOfficeImoDataValidation(sheet, rows.Count);
            AddOfficeImoPivotTable(sheet, rows.Count);
            AddOfficeImoRegionalChart(sheet, rows);
            document.Save(stream);
        }

        return stream.ToArray();
    }

    private static int EpPlusWriteRealWorldReport(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(EpPlusWriteRealWorldReportBytes(rows));

    private static byte[] EpPlusWriteRealWorldReportBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage(stream)) {
            var worksheet = package.Workbook.Worksheets.Add("Data");
            PopulateEpPlusWorksheet(worksheet, rows, includeTable: true, autoFit: true);
            ApplyEpPlusNavigation(worksheet, rows.Count);
            ApplyEpPlusConditionalFormatting(worksheet, rows.Count);
            ApplyEpPlusDataValidation(worksheet, rows.Count);
            AddEpPlusPivotTable(package, worksheet, rows.Count);
            AddEpPlusRegionalChart(package, rows);
            package.Save();
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteRealWorldCoreReport(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(OfficeImoWriteRealWorldCoreReportBytes(rows));

    private static byte[] OfficeImoWriteRealWorldCoreReportBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream, autoSave: false)) {
            var sheet = document.AddWorkSheet("Data");
            ExcelBenchmarkScenarioFactory.InsertOfficeImoObjects(sheet, rows);
            ApplyOfficeImoTable(sheet, rows.Count);
            ApplyOfficeImoNavigation(sheet, rows.Count);
            ApplyOfficeImoConditionalFormatting(sheet, rows.Count);
            ApplyOfficeImoDataValidation(sheet, rows.Count);
            document.Save(stream);
        }

        return stream.ToArray();
    }

    private static int ClosedXmlWriteRealWorldCoreReport(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(ClosedXmlWriteRealWorldCoreReportBytes(rows));

    private static byte[] ClosedXmlWriteRealWorldCoreReportBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook()) {
            var worksheet = workbook.Worksheets.Add("Data");
            WriteSalesRows(worksheet, rows, includeAllColumns: true);
            var table = worksheet.Range(1, 1, rows.Count + 1, 8).CreateTable("SalesData");
            table.Theme = XLTableTheme.TableStyleMedium2;
            worksheet.ColumnsUsed().AdjustToContents();
            worksheet.SheetView.FreezeRows(1);
            ApplyClosedXmlConditionalFormatting(worksheet, rows.Count);
            ApplyClosedXmlDataValidation(worksheet, rows.Count);
            workbook.SaveAs(stream);
        }

        return stream.ToArray();
    }

    private static int EpPlusWriteRealWorldCoreReport(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(EpPlusWriteRealWorldCoreReportBytes(rows));

    private static byte[] EpPlusWriteRealWorldCoreReportBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage(stream)) {
            var worksheet = package.Workbook.Worksheets.Add("Data");
            PopulateEpPlusWorksheet(worksheet, rows, includeTable: true, autoFit: true);
            ApplyEpPlusNavigation(worksheet, rows.Count);
            ApplyEpPlusConditionalFormatting(worksheet, rows.Count);
            ApplyEpPlusDataValidation(worksheet, rows.Count);
            package.Save();
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteRealWorldFreezePanes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(OfficeImoWriteRealWorldFreezePanesBytes(rows));

    private static byte[] OfficeImoWriteRealWorldFreezePanesBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream, autoSave: false)) {
            var sheet = document.AddWorkSheet("Data");
            ExcelBenchmarkScenarioFactory.InsertOfficeImoObjects(sheet, rows);
            sheet.Freeze(topRows: 1, leftCols: 1);
            document.Save(stream);
        }

        return stream.ToArray();
    }

    private static int ClosedXmlWriteRealWorldFreezePanes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(ClosedXmlWriteRealWorldFreezePanesBytes(rows));

    private static byte[] ClosedXmlWriteRealWorldFreezePanesBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook()) {
            var worksheet = workbook.Worksheets.Add("Data");
            WriteSalesRows(worksheet, rows, includeAllColumns: true);
            worksheet.SheetView.FreezeRows(1);
            worksheet.SheetView.FreezeColumns(1);
            workbook.SaveAs(stream);
        }

        return stream.ToArray();
    }

    private static int EpPlusWriteRealWorldFreezePanes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(EpPlusWriteRealWorldFreezePanesBytes(rows));

    private static byte[] EpPlusWriteRealWorldFreezePanesBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage(stream)) {
            var worksheet = package.Workbook.Worksheets.Add("Data");
            WriteSalesRows(worksheet, rows, includeAllColumns: true);
            worksheet.View.FreezePanes(2, 2);
            package.Save();
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteRealWorldAutoFilter(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(OfficeImoWriteRealWorldAutoFilterBytes(rows));

    private static byte[] OfficeImoWriteRealWorldAutoFilterBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream, autoSave: false)) {
            var sheet = document.AddWorkSheet("Data");
            ExcelBenchmarkScenarioFactory.InsertOfficeImoObjects(sheet, rows);
            sheet.AddAutoFilter(BuildSalesRange(rows.Count));
            document.Save(stream);
        }

        return stream.ToArray();
    }

    private static int ClosedXmlWriteRealWorldAutoFilter(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(ClosedXmlWriteRealWorldAutoFilterBytes(rows));

    private static byte[] ClosedXmlWriteRealWorldAutoFilterBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook()) {
            var worksheet = workbook.Worksheets.Add("Data");
            WriteSalesRows(worksheet, rows, includeAllColumns: true);
            worksheet.Range(1, 1, rows.Count + 1, 8).SetAutoFilter();
            workbook.SaveAs(stream);
        }

        return stream.ToArray();
    }

    private static int EpPlusWriteRealWorldAutoFilter(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(EpPlusWriteRealWorldAutoFilterBytes(rows));

    private static byte[] EpPlusWriteRealWorldAutoFilterBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage(stream)) {
            var worksheet = package.Workbook.Worksheets.Add("Data");
            WriteSalesRows(worksheet, rows, includeAllColumns: true);
            worksheet.Cells[1, 1, rows.Count + 1, 8].AutoFilter = true;
            package.Save();
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteRealWorldConditionalFormatting(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(OfficeImoWriteRealWorldConditionalFormattingBytes(rows));

    private static byte[] OfficeImoWriteRealWorldConditionalFormattingBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream, autoSave: false)) {
            var sheet = document.AddWorkSheet("Data");
            ExcelBenchmarkScenarioFactory.InsertOfficeImoObjects(sheet, rows);
            ApplyOfficeImoConditionalFormatting(sheet, rows.Count);
            document.Save(stream);
        }

        return stream.ToArray();
    }

    private static int ClosedXmlWriteRealWorldConditionalFormatting(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(ClosedXmlWriteRealWorldConditionalFormattingBytes(rows));

    private static byte[] ClosedXmlWriteRealWorldConditionalFormattingBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook()) {
            var worksheet = workbook.Worksheets.Add("Data");
            WriteSalesRows(worksheet, rows, includeAllColumns: true);
            ApplyClosedXmlConditionalFormatting(worksheet, rows.Count);
            workbook.SaveAs(stream);
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteRealWorldDataValidation(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(OfficeImoWriteRealWorldDataValidationBytes(rows));

    private static byte[] OfficeImoWriteRealWorldDataValidationBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream, autoSave: false)) {
            var sheet = document.AddWorkSheet("Data");
            ExcelBenchmarkScenarioFactory.InsertOfficeImoObjects(sheet, rows);
            ApplyOfficeImoDataValidation(sheet, rows.Count);
            document.Save(stream);
        }

        return stream.ToArray();
    }

    private static int ClosedXmlWriteRealWorldDataValidation(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(ClosedXmlWriteRealWorldDataValidationBytes(rows));

    private static byte[] ClosedXmlWriteRealWorldDataValidationBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook()) {
            var worksheet = workbook.Worksheets.Add("Data");
            WriteSalesRows(worksheet, rows, includeAllColumns: true);
            ApplyClosedXmlDataValidation(worksheet, rows.Count);
            workbook.SaveAs(stream);
        }

        return stream.ToArray();
    }

    private static int EpPlusWriteRealWorldDataValidation(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(EpPlusWriteRealWorldDataValidationBytes(rows));

    private static byte[] EpPlusWriteRealWorldDataValidationBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage(stream)) {
            var worksheet = package.Workbook.Worksheets.Add("Data");
            WriteSalesRows(worksheet, rows, includeAllColumns: true);
            ApplyEpPlusDataValidation(worksheet, rows.Count);
            package.Save();
        }

        return stream.ToArray();
    }

    private static int EpPlusWriteRealWorldConditionalFormatting(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(EpPlusWriteRealWorldConditionalFormattingBytes(rows));

    private static byte[] EpPlusWriteRealWorldConditionalFormattingBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage(stream)) {
            var worksheet = package.Workbook.Worksheets.Add("Data");
            WriteSalesRows(worksheet, rows, includeAllColumns: true);
            ApplyEpPlusConditionalFormatting(worksheet, rows.Count);
            package.Save();
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteRealWorldCharts(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(OfficeImoWriteRealWorldChartsBytes(rows));

    private static byte[] OfficeImoWriteRealWorldChartsBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream, autoSave: false)) {
            var sheet = document.AddWorkSheet("Data");
            ExcelBenchmarkScenarioFactory.InsertOfficeImoObjects(sheet, rows);
            AddOfficeImoRegionalChart(sheet, rows);
            document.Save(stream);
        }

        return stream.ToArray();
    }

    private static int EpPlusWriteRealWorldCharts(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(EpPlusWriteRealWorldChartsBytes(rows));

    private static byte[] EpPlusWriteRealWorldChartsBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage(stream)) {
            var worksheet = package.Workbook.Worksheets.Add("Data");
            WriteSalesRows(worksheet, rows, includeAllColumns: true);
            AddEpPlusRegionalChart(package, rows);
            package.Save();
        }

        return stream.ToArray();
    }

    private static int OfficeImoWriteRealWorldPivotTable(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(OfficeImoWriteRealWorldPivotTableBytes(rows));

    private static byte[] OfficeImoWriteRealWorldPivotTableBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream, autoSave: false)) {
            var sheet = document.AddWorkSheet("Data");
            ExcelBenchmarkScenarioFactory.InsertOfficeImoObjects(sheet, rows);
            AddOfficeImoPivotTable(sheet, rows.Count);
            document.Save(stream);
        }

        return stream.ToArray();
    }

    private static int EpPlusWriteRealWorldPivotTable(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => ByteCount(EpPlusWriteRealWorldPivotTableBytes(rows));

    private static byte[] EpPlusWriteRealWorldPivotTableBytes(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage(stream)) {
            var worksheet = package.Workbook.Worksheets.Add("Data");
            WriteSalesRows(worksheet, rows, includeAllColumns: true);
            AddEpPlusPivotTable(package, worksheet, rows.Count);
            package.Save();
        }

        return stream.ToArray();
    }

    private static void AddRealWorldVariantScenarioGroups(
        List<ExcelLibraryComparisonScenario> scenarios,
        IReadOnlySet<string>? scenarioFilter,
        IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows,
        int warmupIterations,
        int measuredIterations) {
        AddScenarioGroup(scenarios, scenarioFilter, RealWorldReportNoAutoFitScenario, warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Create the real-world report workbook without AutoFit to check the fast path is not tied to width measurement.", () => OfficeImoWriteRealWorldVariant(rows, RealWorldDefaultColumns, new RealWorldVariantOptions(AutoFit: false))),
            new LibraryComparisonCase("EPPlus", "Create the equivalent real-world report workbook without AutoFit.", () => EpPlusWriteRealWorldVariant(rows, RealWorldDefaultColumns, new RealWorldVariantOptions(AutoFit: false)))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, RealWorldReportChartFirstScenario, warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Create the real-world report workbook with chart creation before pivot creation.", () => OfficeImoWriteRealWorldVariant(rows, RealWorldDefaultColumns, new RealWorldVariantOptions(ChartBeforePivot: true))),
            new LibraryComparisonCase("EPPlus", "Create the equivalent real-world report workbook with chart creation before pivot creation.", () => EpPlusWriteRealWorldVariant(rows, RealWorldDefaultColumns, new RealWorldVariantOptions(ChartBeforePivot: true)))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, RealWorldReportShuffledColumnsScenario, warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Create the real-world report workbook with the same fields in a different column order.", () => OfficeImoWriteRealWorldVariant(rows, RealWorldShuffledColumns, RealWorldVariantOptions.Default)),
            new LibraryComparisonCase("EPPlus", "Create the equivalent real-world report workbook with the same fields in a different column order.", () => EpPlusWriteRealWorldVariant(rows, RealWorldShuffledColumns, RealWorldVariantOptions.Default))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, RealWorldReportExtraColumnScenario, warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Create the real-world report workbook with an extra derived column.", () => OfficeImoWriteRealWorldVariant(rows, RealWorldExtraColumns, RealWorldVariantOptions.Default)),
            new LibraryComparisonCase("EPPlus", "Create the equivalent real-world report workbook with an extra derived column.", () => EpPlusWriteRealWorldVariant(rows, RealWorldExtraColumns, RealWorldVariantOptions.Default))
        ]);

        AddScenarioGroup(scenarios, scenarioFilter, RealWorldReportPostMutationScenario, warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", "Create the real-world report workbook and then make a normal cell edit after report features are added.", () => OfficeImoWriteRealWorldVariant(rows, RealWorldDefaultColumns, new RealWorldVariantOptions(PostMutation: true))),
            new LibraryComparisonCase("EPPlus", "Create the equivalent real-world report workbook and then make a normal cell edit after report features are added.", () => EpPlusWriteRealWorldVariant(rows, RealWorldDefaultColumns, new RealWorldVariantOptions(PostMutation: true)))
        ]);
    }

    private static void AddRealWorldVariantPackageProfileGroups(
        List<ExcelPackageProfileScenario> scenarios,
        IReadOnlySet<string>? scenarioFilter,
        IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows,
        int warmupIterations,
        int measuredIterations) {
        AddPackageProfileGroup(scenarios, scenarioFilter, RealWorldReportNoAutoFitScenario, warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Create the real-world report workbook without AutoFit to check the fast path is not tied to width measurement.", () => OfficeImoWriteRealWorldVariantBytes(rows, RealWorldDefaultColumns, new RealWorldVariantOptions(AutoFit: false))),
            new PackageProfileCase("EPPlus", "Create the equivalent real-world report workbook without AutoFit.", () => EpPlusWriteRealWorldVariantBytes(rows, RealWorldDefaultColumns, new RealWorldVariantOptions(AutoFit: false)))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, RealWorldReportChartFirstScenario, warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Create the real-world report workbook with chart creation before pivot creation.", () => OfficeImoWriteRealWorldVariantBytes(rows, RealWorldDefaultColumns, new RealWorldVariantOptions(ChartBeforePivot: true))),
            new PackageProfileCase("EPPlus", "Create the equivalent real-world report workbook with chart creation before pivot creation.", () => EpPlusWriteRealWorldVariantBytes(rows, RealWorldDefaultColumns, new RealWorldVariantOptions(ChartBeforePivot: true)))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, RealWorldReportShuffledColumnsScenario, warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Create the real-world report workbook with the same fields in a different column order.", () => OfficeImoWriteRealWorldVariantBytes(rows, RealWorldShuffledColumns, RealWorldVariantOptions.Default)),
            new PackageProfileCase("EPPlus", "Create the equivalent real-world report workbook with the same fields in a different column order.", () => EpPlusWriteRealWorldVariantBytes(rows, RealWorldShuffledColumns, RealWorldVariantOptions.Default))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, RealWorldReportExtraColumnScenario, warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Create the real-world report workbook with an extra derived column.", () => OfficeImoWriteRealWorldVariantBytes(rows, RealWorldExtraColumns, RealWorldVariantOptions.Default)),
            new PackageProfileCase("EPPlus", "Create the equivalent real-world report workbook with an extra derived column.", () => EpPlusWriteRealWorldVariantBytes(rows, RealWorldExtraColumns, RealWorldVariantOptions.Default))
        ]);

        AddPackageProfileGroup(scenarios, scenarioFilter, RealWorldReportPostMutationScenario, warmupIterations, measuredIterations, [
            new PackageProfileCase("OfficeIMO.Excel", "Create the real-world report workbook and then make a normal cell edit after report features are added.", () => OfficeImoWriteRealWorldVariantBytes(rows, RealWorldDefaultColumns, new RealWorldVariantOptions(PostMutation: true))),
            new PackageProfileCase("EPPlus", "Create the equivalent real-world report workbook and then make a normal cell edit after report features are added.", () => EpPlusWriteRealWorldVariantBytes(rows, RealWorldDefaultColumns, new RealWorldVariantOptions(PostMutation: true)))
        ]);
    }

    private static int OfficeImoWriteRealWorldVariant(
        IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows,
        IReadOnlyList<RealWorldColumnSpec> columns,
        RealWorldVariantOptions options)
        => ByteCount(OfficeImoWriteRealWorldVariantBytes(rows, columns, options));

    private static byte[] OfficeImoWriteRealWorldVariantBytes(
        IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows,
        IReadOnlyList<RealWorldColumnSpec> columns,
        RealWorldVariantOptions options) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream, autoSave: false)) {
            var sheet = document.AddWorkSheet("Data");
            InsertOfficeImoVariantRows(sheet, rows, columns);
            string range = BuildVariantRange(rows.Count, columns.Count);
            ApplyOfficeImoVariantTable(sheet, range, options.AutoFit);
            ApplyOfficeImoNavigation(sheet, range);
            ApplyOfficeImoVariantConditionalFormatting(sheet, rows.Count, columns);
            ApplyOfficeImoVariantDataValidation(sheet, rows.Count, columns);

            if (options.ChartBeforePivot) {
                AddOfficeImoRegionalChart(sheet, rows);
                AddOfficeImoVariantPivotTable(sheet, range);
            } else {
                AddOfficeImoVariantPivotTable(sheet, range);
                AddOfficeImoRegionalChart(sheet, rows);
            }

            if (options.PostMutation) {
                sheet.CellValue(rows.Count + 4, 1, "Manual note after report features");
            }

            document.Save(stream);
        }

        return stream.ToArray();
    }

    private static int EpPlusWriteRealWorldVariant(
        IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows,
        IReadOnlyList<RealWorldColumnSpec> columns,
        RealWorldVariantOptions options)
        => ByteCount(EpPlusWriteRealWorldVariantBytes(rows, columns, options));

    private static byte[] EpPlusWriteRealWorldVariantBytes(
        IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows,
        IReadOnlyList<RealWorldColumnSpec> columns,
        RealWorldVariantOptions options) {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage(stream)) {
            var worksheet = package.Workbook.Worksheets.Add("Data");
            WriteEpPlusVariantRows(worksheet, rows, columns);
            ApplyEpPlusVariantTable(worksheet, rows.Count, columns.Count, options.AutoFit);
            ApplyEpPlusNavigation(worksheet, rows.Count, columns.Count);
            ApplyEpPlusVariantConditionalFormatting(worksheet, rows.Count, columns);
            ApplyEpPlusVariantDataValidation(worksheet, rows.Count, columns);

            if (options.ChartBeforePivot) {
                AddEpPlusRegionalChart(package, rows);
                AddEpPlusVariantPivotTable(package, worksheet, rows.Count, columns.Count);
            } else {
                AddEpPlusVariantPivotTable(package, worksheet, rows.Count, columns.Count);
                AddEpPlusRegionalChart(package, rows);
            }

            if (options.PostMutation) {
                worksheet.Cells[rows.Count + 4, 1].Value = "Manual note after report features";
            }

            package.Save();
        }

        return stream.ToArray();
    }

    private static void InsertOfficeImoVariantRows(
        ExcelSheet sheet,
        IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows,
        IReadOnlyList<RealWorldColumnSpec> columns) {
        var selectors = new (string Header, Func<ExcelBenchmarkScenarioFactory.SalesRecord, object?> Selector)[columns.Count];
        for (int i = 0; i < columns.Count; i++) {
            selectors[i] = (columns[i].Header, columns[i].Selector);
        }

        sheet.InsertObjects(rows, selectors);
    }

    private static void WriteEpPlusVariantRows(
        ExcelWorksheet worksheet,
        IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows,
        IReadOnlyList<RealWorldColumnSpec> columns) {
        for (int column = 0; column < columns.Count; column++) {
            worksheet.Cells[1, column + 1].Value = columns[column].Header;
        }

        for (int row = 0; row < rows.Count; row++) {
            var source = rows[row];
            for (int column = 0; column < columns.Count; column++) {
                worksheet.Cells[row + 2, column + 1].Value = columns[column].Selector(source);
            }
        }
    }

    private static void ApplyOfficeImoVariantTable(ExcelSheet sheet, string range, bool autoFit) {
        sheet.AddTable(range, hasHeader: true, name: "SalesData", style: OfficeIMO.Excel.TableStyle.TableStyleMedium2, includeAutoFilter: true);
        if (autoFit) {
            sheet.AutoFitColumns();
        }
    }

    private static void ApplyEpPlusVariantTable(ExcelWorksheet worksheet, int rowCount, int columnCount, bool autoFit) {
        var table = worksheet.Tables.Add(worksheet.Cells[1, 1, rowCount + 1, columnCount], "SalesData");
        table.TableStyle = OfficeOpenXml.Table.TableStyles.Medium2;
        if (autoFit) {
            worksheet.Cells[1, 1, rowCount + 1, columnCount].AutoFitColumns();
        }
    }

    private static void ApplyOfficeImoNavigation(ExcelSheet sheet, string range) {
        sheet.Freeze(topRows: 1, leftCols: 1);
        sheet.AddAutoFilter(range);
    }

    private static void ApplyEpPlusNavigation(ExcelWorksheet worksheet, int rowCount, int columnCount) {
        worksheet.View.FreezePanes(2, 2);
        worksheet.Cells[1, 1, rowCount + 1, columnCount].AutoFilter = true;
    }

    private static void ApplyOfficeImoVariantConditionalFormatting(ExcelSheet sheet, int rowCount, IReadOnlyList<RealWorldColumnSpec> columns) {
        int amountColumn = GetColumnIndex(columns, "Amount");
        int unitsColumn = GetColumnIndex(columns, "Units");
        int lastRow = rowCount + 1;
        string amountRange = BuildColumnRange(amountColumn, 2, lastRow);
        string unitsRange = BuildColumnRange(unitsColumn, 2, lastRow);

        sheet.AddConditionalRule(amountRange, ConditionalFormattingOperatorValues.GreaterThan, "3000");
        sheet.AddConditionalRule(unitsRange, ConditionalFormattingOperatorValues.LessThan, "5");
        sheet.AddConditionalColorScale(amountRange, OfficeColor.LightPink, OfficeColor.LightGreen);
        sheet.AddConditionalDataBar(unitsRange, OfficeColor.SteelBlue);
    }

    private static void ApplyEpPlusVariantConditionalFormatting(ExcelWorksheet worksheet, int rowCount, IReadOnlyList<RealWorldColumnSpec> columns) {
        int amountColumn = GetColumnIndex(columns, "Amount");
        int unitsColumn = GetColumnIndex(columns, "Units");
        int lastRow = rowCount + 1;

        var highAmount = worksheet.ConditionalFormatting.AddGreaterThan(worksheet.Cells[2, amountColumn, lastRow, amountColumn]);
        highAmount.Formula = "3000";
        highAmount.Style.Fill.PatternType = ExcelFillStyle.Solid;
        highAmount.Style.Fill.BackgroundColor.Color = System.Drawing.Color.LightGreen;

        var lowUnits = worksheet.ConditionalFormatting.AddLessThan(worksheet.Cells[2, unitsColumn, lastRow, unitsColumn]);
        lowUnits.Formula = "5";
        lowUnits.Style.Fill.PatternType = ExcelFillStyle.Solid;
        lowUnits.Style.Fill.BackgroundColor.Color = System.Drawing.Color.LightPink;
    }

    private static void ApplyOfficeImoVariantDataValidation(ExcelSheet sheet, int rowCount, IReadOnlyList<RealWorldColumnSpec> columns) {
        int unitsColumn = GetColumnIndex(columns, "Units");
        int lastRow = rowCount + 1;
        sheet.ValidationWholeNumber(BuildColumnRange(unitsColumn, 2, lastRow), DataValidationOperatorValues.Between, 1, 24);
    }

    private static void ApplyEpPlusVariantDataValidation(ExcelWorksheet worksheet, int rowCount, IReadOnlyList<RealWorldColumnSpec> columns) {
        int unitsColumn = GetColumnIndex(columns, "Units");
        int lastRow = rowCount + 1;
        var validation = worksheet.DataValidations.AddIntegerValidation(BuildColumnRange(unitsColumn, 2, lastRow));
        validation.Operator = OfficeOpenXml.DataValidation.ExcelDataValidationOperator.between;
        validation.Formula.Value = 1;
        validation.Formula2.Value = 24;
    }

    private static void AddOfficeImoVariantPivotTable(ExcelSheet sheet, string range) {
        sheet.AddPivotTable(
            sourceRange: range,
            destinationCell: "J3",
            name: "SalesPivot",
            rowFields: new[] { "Region" },
            columnFields: new[] { "Owner" },
            dataFields: new[] { new ExcelPivotDataField("Amount", DataConsolidateFunctionValues.Sum, "Total Amount") },
            pivotStyleName: "PivotStyleMedium9");
    }

    private static void AddEpPlusVariantPivotTable(ExcelPackage package, ExcelWorksheet dataWorksheet, int rowCount, int columnCount) {
        var pivotSheet = package.Workbook.Worksheets.Add("Pivot");
        var source = dataWorksheet.Cells[1, 1, rowCount + 1, columnCount];
        var pivot = pivotSheet.PivotTables.Add(pivotSheet.Cells["A3"], source, "SalesPivot");
        pivot.RowFields.Add(pivot.Fields["Region"]);
        pivot.ColumnFields.Add(pivot.Fields["Owner"]);
        var amount = pivot.DataFields.Add(pivot.Fields["Amount"]);
        amount.Function = DataFieldFunctions.Sum;
        amount.Name = "Total Amount";
    }

    private static string BuildVariantRange(int rowCount, int columnCount)
        => "A1:" + GetColumnLetter(columnCount) + (rowCount + 1).ToString(CultureInfo.InvariantCulture);

    private static string BuildColumnRange(int columnIndex, int firstRow, int lastRow) {
        string column = GetColumnLetter(columnIndex);
        return column + firstRow.ToString(CultureInfo.InvariantCulture) + ":" + column + lastRow.ToString(CultureInfo.InvariantCulture);
    }

    private static int GetColumnIndex(IReadOnlyList<RealWorldColumnSpec> columns, string header) {
        for (int i = 0; i < columns.Count; i++) {
            if (string.Equals(columns[i].Header, header, StringComparison.OrdinalIgnoreCase)) {
                return i + 1;
            }
        }

        throw new InvalidOperationException($"Column '{header}' was not found in the benchmark variant.");
    }

    private static string GetColumnLetter(int columnIndex) {
        Span<char> buffer = stackalloc char[8];
        int position = buffer.Length;
        int value = columnIndex;
        while (value > 0) {
            value--;
            buffer[--position] = (char)('A' + (value % 26));
            value /= 26;
        }

        return new string(buffer[position..]);
    }

    private static readonly RealWorldColumnSpec[] RealWorldDefaultColumns = [
        new("Id", static item => item.Id),
        new("Region", static item => item.Region),
        new("Owner", static item => item.Owner),
        new("CreatedOn", static item => item.CreatedOn),
        new("Amount", static item => item.Amount),
        new("Units", static item => item.Units),
        new("Active", static item => item.Active),
        new("Notes", static item => item.Notes)
    ];

    private static readonly RealWorldColumnSpec[] RealWorldShuffledColumns = [
        new("Owner", static item => item.Owner),
        new("Region", static item => item.Region),
        new("Id", static item => item.Id),
        new("Amount", static item => item.Amount),
        new("CreatedOn", static item => item.CreatedOn),
        new("Units", static item => item.Units),
        new("Notes", static item => item.Notes),
        new("Active", static item => item.Active)
    ];

    private static readonly RealWorldColumnSpec[] RealWorldExtraColumns = [
        new("Id", static item => item.Id),
        new("Region", static item => item.Region),
        new("Owner", static item => item.Owner),
        new("CreatedOn", static item => item.CreatedOn),
        new("Amount", static item => item.Amount),
        new("Units", static item => item.Units),
        new("Active", static item => item.Active),
        new("Notes", static item => item.Notes),
        new("AmountBand", static item => item.Amount >= 3000 ? "High" : item.Amount >= 1000 ? "Medium" : "Low")
    ];

    private sealed record RealWorldColumnSpec(string Header, Func<ExcelBenchmarkScenarioFactory.SalesRecord, object?> Selector);

    private sealed record RealWorldVariantOptions(bool AutoFit = true, bool ChartBeforePivot = false, bool PostMutation = false) {
        public static readonly RealWorldVariantOptions Default = new();
    }

    private static void ApplyOfficeImoTable(ExcelSheet sheet, int rowCount) {
        sheet.AddTable(
            BuildSalesRange(rowCount),
            hasHeader: true,
            name: "SalesData",
            style: OfficeIMO.Excel.TableStyle.TableStyleMedium2,
            includeAutoFilter: true);
        sheet.AutoFitColumns();
    }

    private static void ApplyOfficeImoNavigation(ExcelSheet sheet, int rowCount) {
        sheet.Freeze(topRows: 1, leftCols: 1);
        sheet.AddAutoFilter(BuildSalesRange(rowCount));
    }

    private static void ApplyOfficeImoConditionalFormatting(ExcelSheet sheet, int rowCount) {
        int lastRow = rowCount + 1;
        sheet.AddConditionalRule($"E2:E{lastRow.ToString(CultureInfo.InvariantCulture)}", ConditionalFormattingOperatorValues.GreaterThan, "3000");
        sheet.AddConditionalRule($"F2:F{lastRow.ToString(CultureInfo.InvariantCulture)}", ConditionalFormattingOperatorValues.LessThan, "5");
        sheet.AddConditionalColorScale($"E2:E{lastRow.ToString(CultureInfo.InvariantCulture)}", OfficeColor.LightPink, OfficeColor.LightGreen);
        sheet.AddConditionalDataBar($"F2:F{lastRow.ToString(CultureInfo.InvariantCulture)}", OfficeColor.SteelBlue);
    }

    private static void ApplyOfficeImoDataValidation(ExcelSheet sheet, int rowCount) {
        int lastRow = rowCount + 1;
        sheet.ValidationWholeNumber($"F2:F{lastRow.ToString(CultureInfo.InvariantCulture)}", DataValidationOperatorValues.Between, 1, 24);
    }

    private static void AddOfficeImoPivotTable(ExcelSheet sheet, int rowCount) {
        sheet.AddPivotTable(
            sourceRange: BuildSalesRange(rowCount),
            destinationCell: "J3",
            name: "SalesPivot",
            rowFields: new[] { "Region" },
            columnFields: new[] { "Owner" },
            dataFields: new[] { new ExcelPivotDataField("Amount", DataConsolidateFunctionValues.Sum, "Total Amount") },
            pivotStyleName: "PivotStyleMedium9");
    }

    private static void AddOfficeImoRegionalChart(ExcelSheet sheet, IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        var summaries = BuildRegionSummaries(rows);
        var chartData = new ExcelChartData(
            summaries.Select(static item => item.Region),
            new[] {
                new ExcelChartSeries("Amount", summaries.Select(static item => item.Amount)),
                new ExcelChartSeries("Units", summaries.Select(static item => (double)item.Units))
            });

        sheet.AddChart(chartData, row: 18, column: 10, widthPixels: 720, heightPixels: 360, type: ExcelChartType.ColumnClustered, title: "Regional Sales");
    }

    private static void ApplyEpPlusNavigation(ExcelWorksheet worksheet, int rowCount) {
        worksheet.View.FreezePanes(2, 2);
        worksheet.Cells[1, 1, rowCount + 1, 8].AutoFilter = true;
    }

    private static void ApplyEpPlusConditionalFormatting(ExcelWorksheet worksheet, int rowCount) {
        int lastRow = rowCount + 1;
        var highAmount = worksheet.ConditionalFormatting.AddGreaterThan(worksheet.Cells[2, 5, lastRow, 5]);
        highAmount.Formula = "3000";
        highAmount.Style.Fill.PatternType = ExcelFillStyle.Solid;
        highAmount.Style.Fill.BackgroundColor.Color = System.Drawing.Color.LightGreen;

        var lowUnits = worksheet.ConditionalFormatting.AddLessThan(worksheet.Cells[2, 6, lastRow, 6]);
        lowUnits.Formula = "5";
        lowUnits.Style.Fill.PatternType = ExcelFillStyle.Solid;
        lowUnits.Style.Fill.BackgroundColor.Color = System.Drawing.Color.LightPink;
    }

    private static void ApplyClosedXmlConditionalFormatting(IXLWorksheet worksheet, int rowCount) {
        int lastRow = rowCount + 1;
        worksheet.Range(2, 5, lastRow, 5).AddConditionalFormat().WhenGreaterThan(3000).Fill.SetBackgroundColor(XLColor.LightGreen);
        worksheet.Range(2, 6, lastRow, 6).AddConditionalFormat().WhenLessThan(5).Fill.SetBackgroundColor(XLColor.LightPink);
    }

    private static void ApplyClosedXmlDataValidation(IXLWorksheet worksheet, int rowCount) {
        int lastRow = rowCount + 1;
        worksheet.Range(2, 6, lastRow, 6).CreateDataValidation().WholeNumber.Between(1, 24);
    }

    private static void ApplyEpPlusDataValidation(ExcelWorksheet worksheet, int rowCount) {
        int lastRow = rowCount + 1;
        var validation = worksheet.DataValidations.AddIntegerValidation($"F2:F{lastRow.ToString(CultureInfo.InvariantCulture)}");
        validation.Operator = OfficeOpenXml.DataValidation.ExcelDataValidationOperator.between;
        validation.Formula.Value = 1;
        validation.Formula2.Value = 24;
    }

    private static void AddEpPlusPivotTable(ExcelPackage package, ExcelWorksheet dataWorksheet, int rowCount) {
        var pivotSheet = package.Workbook.Worksheets.Add("Pivot");
        var source = dataWorksheet.Cells[1, 1, rowCount + 1, 8];
        var pivot = pivotSheet.PivotTables.Add(pivotSheet.Cells["A3"], source, "SalesPivot");
        pivot.RowFields.Add(pivot.Fields["Region"]);
        pivot.ColumnFields.Add(pivot.Fields["Owner"]);
        var amount = pivot.DataFields.Add(pivot.Fields["Amount"]);
        amount.Function = DataFieldFunctions.Sum;
        amount.Name = "Total Amount";
    }

    private static void AddEpPlusRegionalChart(ExcelPackage package, IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        var summaries = BuildRegionSummaries(rows);
        var chartSheet = package.Workbook.Worksheets.Add("ChartData");
        chartSheet.Cells[1, 1].Value = "Region";
        chartSheet.Cells[1, 2].Value = "Amount";
        chartSheet.Cells[1, 3].Value = "Units";
        for (int i = 0; i < summaries.Count; i++) {
            int row = i + 2;
            chartSheet.Cells[row, 1].Value = summaries[i].Region;
            chartSheet.Cells[row, 2].Value = summaries[i].Amount;
            chartSheet.Cells[row, 3].Value = summaries[i].Units;
        }

        var chart = chartSheet.Drawings.AddChart("RegionalSales", eChartType.ColumnClustered);
        chart.Title.Text = "Regional Sales";
        chart.SetPosition(1, 0, 5, 0);
        chart.SetSize(720, 360);
        chart.Series.Add(chartSheet.Cells[2, 2, summaries.Count + 1, 2], chartSheet.Cells[2, 1, summaries.Count + 1, 1]);
        chart.Series.Add(chartSheet.Cells[2, 3, summaries.Count + 1, 3], chartSheet.Cells[2, 1, summaries.Count + 1, 1]);
    }

    private static string BuildSalesRange(int rowCount)
        => "A1:H" + (rowCount + 1).ToString(CultureInfo.InvariantCulture);

    private static IReadOnlyList<RegionSummary> BuildRegionSummaries(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => rows
            .GroupBy(static row => row.Region, StringComparer.Ordinal)
            .OrderBy(static group => group.Key, StringComparer.Ordinal)
            .Select(static group => new RegionSummary(
                group.Key,
                Math.Round(group.Sum(static row => row.Amount), 2),
                group.Sum(static row => row.Units)))
            .ToArray();

    private sealed record RegionSummary(string Region, double Amount, int Units);
}
