using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;

const int DefaultRowCount = 2500;
const int SparseLastRow = 100_001;
const int WarmupIterations = 3;
const int MeasuredIterations = 5;
const string LibraryName = "EPPlus 4.5.3.3";

string outputPath = ParseOutputPath(args)
    ?? Path.Combine("Docs", "benchmarks", "officeimo.excel.legacy-epplus-comparison.json");
int rowCount = ParseRowCount(args);
if (rowCount <= 0) {
    throw new ArgumentOutOfRangeException(nameof(rowCount));
}
int warmupIterations = ParsePositiveOption(args, "--warmup", "--warmups") ?? WarmupIterations;
int measuredIterations = ParsePositiveOption(args, "--iterations", "--measured-iterations", "--samples") ?? MeasuredIterations;
var scenarioFilter = BuildScenarioFilter(ParseOptionValues(args, "--scenario", "--scenarios"));

var rows = CreateSalesRecords(rowCount);
var salesDataTable = CreateSalesDataTableFromRows(rows, "SalesData");
var salesDataSet = CreateSalesDataSet(rows);
var powerShellMixedRows = CreatePowerShellMixedRows(rowCount);
var powerShellMixedDataTable = CreatePowerShellMixedDataTable(powerShellMixedRows, "PowerShellMixed");
int topDataRows = Math.Min(rowCount, 100);
byte[] workbookBytes = CreateWorkbookBytes(rows);
byte[] formulaWorkbookBytes = CreateFormulaWorkbookBytes(rowCount);
byte[] sharedStringWorkbookBytes = CreateSharedStringWorkbookBytes(rowCount);
byte[] sparseWorkbookBytes = CreateSparseWorkbookBytes(SparseLastRow);
RealWorldColumnSpec[] RealWorldDefaultColumns = [
    new("Id", static item => item.Id),
    new("Region", static item => item.Region),
    new("Owner", static item => item.Owner),
    new("CreatedOn", static item => item.CreatedOn),
    new("Amount", static item => item.Amount),
    new("Units", static item => item.Units),
    new("Active", static item => item.Active),
    new("Notes", static item => item.Notes)
];
RealWorldColumnSpec[] RealWorldShuffledColumns = [
    new("Owner", static item => item.Owner),
    new("Region", static item => item.Region),
    new("Id", static item => item.Id),
    new("Amount", static item => item.Amount),
    new("CreatedOn", static item => item.CreatedOn),
    new("Units", static item => item.Units),
    new("Notes", static item => item.Notes),
    new("Active", static item => item.Active)
];
RealWorldColumnSpec[] RealWorldExtraColumns = [
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
var scenarios = new List<LegacyComparisonScenario>();

AddScenario(scenarios, scenarioFilter, "write-bulk-report", LibraryName, "EPPlus 4.x manual row population, add table, autofit, save.", () => WriteBulkReport(rows), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "write-dataset-tables", LibraryName, "EPPlus 4.x import prepared DataTables as two styled worksheet tables and save.", () => WriteDataSetTables(salesDataSet), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "write-datatable-direct", LibraryName, "EPPlus 4.x import a prepared DataTable and save.", () => WriteDataTable(salesDataTable), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "write-datatable-table-direct", LibraryName, "EPPlus 4.x import a prepared DataTable as a styled worksheet table and save.", () => WriteDataTable(salesDataTable, includeTable: true), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "write-datareader-table", LibraryName, "EPPlus 4.x import equivalent prepared data as a styled worksheet table and save.", () => WriteDataTable(salesDataTable, includeTable: true), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "write-datareader-table-autofit", LibraryName, "EPPlus 4.x import equivalent prepared data as a styled worksheet table, autofit columns, and save.", () => WriteDataTable(salesDataTable, includeTable: true, autoFit: true), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "write-datareader-plain", LibraryName, "EPPlus 4.x import equivalent prepared data as plain worksheet rows and save.", () => WriteDataTable(salesDataTable), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "write-cellvalues-rectangle-direct", LibraryName, "EPPlus 4.x write the same complete A1 rectangle and save.", () => WriteEquivalentSalesRows(rows), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "write-cellvalue-strings", LibraryName, "EPPlus 4.x assign repeated and distinct text-heavy cells one by one and save.", () => WriteSharedStrings(rowCount), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "write-cellvalue-numbers", LibraryName, "EPPlus 4.x assign numeric cells one by one and save.", () => WriteCellValueNumbers(rowCount), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "write-cellvalue-scalars", LibraryName, "EPPlus 4.x assign decimal and boolean cells one by one and save.", () => WriteCellValueScalars(rowCount), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "write-cellvalue-temporal", LibraryName, "EPPlus 4.x assign date and duration cells one by one and save.", () => WriteCellValueTemporal(rowCount), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "write-cellvalue-object-mixed", LibraryName, "EPPlus 4.x assign mixed object-typed cells one by one and save.", () => WriteCellValueObjectMixed(rowCount), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "write-cellformula", LibraryName, "EPPlus 4.x assign numeric cells and row formulas one by one and save.", () => WriteCellFormula(rowCount), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "write-insertobjects-direct", LibraryName, "EPPlus 4.x import equivalent typed object data and save.", () => WriteDataTable(salesDataTable), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "write-fluent-rowsfrom-direct", LibraryName, "EPPlus 4.x import equivalent typed row data and save.", () => WriteDataTable(salesDataTable), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "append-plain-rows", LibraryName, "EPPlus 4.x append equivalent row/cell values.", () => AppendPlainRows(rows), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "copy-worksheet-package", LibraryName, "EPPlus 4.x copy one worksheet between workbooks with the library worksheet-copy API.", () => CopyWorksheet(workbookBytes), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "read-range", LibraryName, "EPPlus 4.x iterate used data cells from workbook.", () => ReadRange(workbookBytes), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "read-top-range", LibraryName, "EPPlus 4.x read the first 100 data rows from a larger sheet.", () => ReadRange(workbookBytes, topDataRows), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "read-datatable", LibraryName, "EPPlus 4.x manual DataTable materialization from worksheet rows.", () => ReadDataTable(workbookBytes), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "read-range-stream", LibraryName, "EPPlus 4.x iterate used data cells row-by-row.", () => ReadRange(workbookBytes), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "read-top-range-stream", LibraryName, "EPPlus 4.x read the first 100 data rows from a larger sheet row-by-row.", () => ReadRange(workbookBytes, topDataRows), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "large-sparse-column-read", LibraryName, "EPPlus 4.x read A1:A100001 with only first and last rows populated.", () => ReadSparseColumn(sparseWorkbookBytes, SparseLastRow), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "large-sparse-row-read", LibraryName, "EPPlus 4.x read A1:A100001 as rows with only first and last rows populated.", () => ReadSparseColumn(sparseWorkbookBytes, SparseLastRow), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "read-objects", LibraryName, "EPPlus 4.x manual typed materialization from worksheet rows.", () => ReadObjects(workbookBytes), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "read-objects-stream", LibraryName, "EPPlus 4.x manual typed materialization from worksheet rows.", () => ReadObjects(workbookBytes), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "autofit-existing", LibraryName, "EPPlus 4.x load existing workbook, autofit columns, save.", () => AutoFitExisting(workbookBytes), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "realworld-report-all-in-one", LibraryName, "EPPlus 4.x create a sales workbook with table, AutoFit, freeze panes, filters, conditional formatting, data validation, pivot table, chart, and save.", () => WriteRealWorldReport(rows), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "realworld-report-no-autofit", LibraryName, "EPPlus 4.x create the real-world report workbook without AutoFit.", () => WriteRealWorldVariant(rows, RealWorldDefaultColumns, new RealWorldVariantOptions(AutoFit: false)), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "realworld-report-chart-first", LibraryName, "EPPlus 4.x create the real-world report workbook with chart creation before pivot creation.", () => WriteRealWorldVariant(rows, RealWorldDefaultColumns, new RealWorldVariantOptions(ChartBeforePivot: true)), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "realworld-report-shuffled-columns", LibraryName, "EPPlus 4.x create the real-world report workbook with the same fields in a different column order.", () => WriteRealWorldVariant(rows, RealWorldShuffledColumns, RealWorldVariantOptions.Default), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "realworld-report-extra-column", LibraryName, "EPPlus 4.x create the real-world report workbook with an extra derived column.", () => WriteRealWorldVariant(rows, RealWorldExtraColumns, RealWorldVariantOptions.Default), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "realworld-report-post-mutation", LibraryName, "EPPlus 4.x create the real-world report workbook and then make a normal cell edit after report features are added.", () => WriteRealWorldVariant(rows, RealWorldDefaultColumns, new RealWorldVariantOptions(PostMutation: true)), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "realworld-report-core", LibraryName, "EPPlus 4.x create a sales workbook with table, AutoFit, frozen header, AutoFilter, conditional formatting, data validation, and save.", () => WriteRealWorldCoreReport(rows), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "realworld-freeze-panes", LibraryName, "EPPlus 4.x write a sales table, freeze the header row and first column, and save.", () => WriteRealWorldFreezePanes(rows), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "realworld-autofilter", LibraryName, "EPPlus 4.x write a sales table, add worksheet-level AutoFilter, and save.", () => WriteRealWorldAutoFilter(rows), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "realworld-conditional-formatting", LibraryName, "EPPlus 4.x write a sales table, add equivalent value rules, and save.", () => WriteRealWorldConditionalFormatting(rows), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "realworld-data-validation", LibraryName, "EPPlus 4.x write a sales table, add whole-number data validation to the Units column, and save.", () => WriteRealWorldDataValidation(rows), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "realworld-charts", LibraryName, "EPPlus 4.x write sales data, add a clustered column chart over regional totals, and save.", () => WriteRealWorldCharts(rows), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "realworld-pivot-table", LibraryName, "EPPlus 4.x write sales data, add a pivot table with row, column, and sum data fields, and save.", () => WriteRealWorldPivotTable(rows), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "report-workbook", LibraryName, "EPPlus 4.x create the PSWriteOffice report workbook shape from the same mixed object rows with table, AutoFit, freeze top row, conditional formatting, list validation, number formats, clustered column chart, pivot table, and save.", () => WriteReportWorkbook(powerShellMixedRows), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "report-workbook-core", LibraryName, "EPPlus 4.x create the report workbook table/core formatting shape from the same mixed object rows without chart or pivot table.", () => WriteReportWorkbookCore(powerShellMixedRows), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "report-workbook-datatable", LibraryName, "EPPlus 4.x create the report workbook shape from the same typed DataTable with table, AutoFit, freeze top row, conditional formatting, list validation, number formats, clustered column chart, pivot table, and save.", () => WriteReportWorkbookDataTable(powerShellMixedDataTable), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "report-workbook-datatable-core", LibraryName, "EPPlus 4.x create the report workbook table/core formatting shape from the same typed DataTable without chart or pivot table.", () => WriteReportWorkbookDataTableCore(powerShellMixedDataTable), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "write-text-heavy-default", LibraryName, "EPPlus 4.x write repeated and distinct text-heavy cells using its valid default storage strategy.", () => WriteSharedStrings(rowCount), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "formula-heavy-read", LibraryName, "EPPlus 4.x read formula text from formula cells.", () => ReadFormulaText(formulaWorkbookBytes, rowCount), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "read-shared-strings-materialized", LibraryName, "EPPlus 4.x materialize the indexed shared-string-table payload.", () => ReadSharedStrings(sharedStringWorkbookBytes, rowCount), warmupIterations, measuredIterations);

if (scenarios.Count == 0) {
    throw new ArgumentException("No legacy EPPlus comparison scenarios matched the requested --scenario filter.");
}

var profile = new LegacyComparisonProfile {
    GeneratedAtUtc = DateTime.UtcNow,
    Framework = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
    MachineName = Environment.MachineName,
    RowCount = rowCount,
    WarmupIterations = warmupIterations,
    MeasuredIterations = measuredIterations,
    Notes = "Local opt-in legacy EPPlus comparison. Not intended for CI gating.",
    Scenarios = scenarios
};

string? directory = Path.GetDirectoryName(outputPath);
if (!string.IsNullOrEmpty(directory)) {
    Directory.CreateDirectory(directory);
}

File.WriteAllText(outputPath, JsonSerializer.Serialize(profile, new JsonSerializerOptions { WriteIndented = true }));
Console.WriteLine($"Legacy EPPlus comparison written to '{outputPath}'.");

static LegacyComparisonScenario Measure(
    string scenario,
    string library,
    string notes,
    Func<int> action,
    int warmupIterations,
    int measuredIterations) {
    Console.WriteLine($"Running {scenario} / {library}...");
    var measurement = BenchmarkMeasurement.Measure(warmupIterations, measuredIterations, action);
    Console.WriteLine(
        string.Create(
            CultureInfo.InvariantCulture,
            $"{scenario} / {library}: avg {measurement.AverageMilliseconds:F2} ms, median {measurement.MedianMilliseconds:F2} ms"));

    return new LegacyComparisonScenario {
        Scenario = scenario,
        Library = library,
        Notes = notes,
        OutputMetric = measurement.OutputMetric,
        AverageMilliseconds = measurement.AverageMilliseconds,
        MedianMilliseconds = measurement.MedianMilliseconds,
        SamplesMilliseconds = measurement.SamplesMilliseconds.ToList()
    };
}

static HashSet<string>? BuildScenarioFilter(IReadOnlyCollection<string> scenarioFilters) {
    if (scenarioFilters.Count == 0) {
        return null;
    }

    var filter = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    foreach (string scenario in scenarioFilters) {
        if (!string.IsNullOrWhiteSpace(scenario)) {
            filter.Add(scenario.Trim());
        }
    }

    return filter.Count == 0 ? null : filter;
}

static int? ParsePositiveOption(string[] args, params string[] optionNames) {
    for (int i = 0; i < args.Length; i++) {
        if (!optionNames.Any(name => string.Equals(args[i], name, StringComparison.OrdinalIgnoreCase))) {
            continue;
        }

        if (i + 1 >= args.Length || args[i + 1].StartsWith("-", StringComparison.Ordinal)) {
            throw new ArgumentException($"Missing value for {args[i]}.");
        }

        string value = args[i + 1].Replace(",", string.Empty, StringComparison.Ordinal).Replace("_", string.Empty, StringComparison.Ordinal);
        if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) || parsed <= 0) {
            throw new ArgumentException($"{args[i]} must be a positive integer.");
        }

        return parsed;
    }

    return null;
}

static void AddScenario(
    List<LegacyComparisonScenario> scenarios,
    IReadOnlySet<string>? scenarioFilter,
    string scenario,
    string library,
    string notes,
    Func<int> action,
    int warmupIterations,
    int measuredIterations) {
    if (scenarioFilter != null && !scenarioFilter.Contains(scenario)) {
        return;
    }

    scenarios.Add(Measure(scenario, library, notes, action, warmupIterations, measuredIterations));
}

static int WriteBulkReport(IReadOnlyList<SalesRecord> rows) {
    using var stream = new MemoryStream();
    using (var package = new ExcelPackage(stream)) {
        var worksheet = package.Workbook.Worksheets.Add("Data");
        PopulateWorksheet(worksheet, rows, includeTable: true, autoFit: true);
        package.Save();
    }

    return checked((int)stream.Length);
}

static int WriteDataSetTables(DataSet dataSet) {
    using var stream = new MemoryStream();
    using (var package = new ExcelPackage(stream)) {
        foreach (DataTable dataTable in dataSet.Tables) {
            var worksheet = package.Workbook.Worksheets.Add(dataTable.TableName);
            worksheet.Cells["A1"].LoadFromDataTable(dataTable, true, TableStyles.Medium2);
        }

        package.Save();
    }

    return checked((int)stream.Length);
}

static int WriteDataTable(DataTable dataTable, bool includeTable = false, bool autoFit = false) {
    using var stream = new MemoryStream();
    using (var package = new ExcelPackage(stream)) {
        var worksheet = package.Workbook.Worksheets.Add("Data");
        worksheet.Cells["A1"].LoadFromDataTable(dataTable, true, includeTable ? TableStyles.Medium2 : TableStyles.None);
        if (autoFit && worksheet.Dimension != null) {
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
        }

        package.Save();
    }

    return checked((int)stream.Length);
}

static int WriteEquivalentSalesRows(IReadOnlyList<SalesRecord> rows) {
    using var stream = new MemoryStream();
    using (var package = new ExcelPackage(stream)) {
        var worksheet = package.Workbook.Worksheets.Add("Data");
        PopulateWorksheet(worksheet, rows, includeTable: false, autoFit: false);
        package.Save();
    }

    return checked((int)stream.Length);
}

static int AppendPlainRows(IReadOnlyList<SalesRecord> rows) {
    using var stream = new MemoryStream();
    using (var package = new ExcelPackage(stream)) {
        var worksheet = package.Workbook.Worksheets.Add("Data");
        WriteAppendHeaders(worksheet);
        for (int i = 0; i < rows.Count; i++) {
            var row = rows[i];
            int r = i + 2;
            worksheet.Cells[r, 1].Value = row.Id;
            worksheet.Cells[r, 2].Value = row.Region;
            worksheet.Cells[r, 3].Value = row.Owner;
            worksheet.Cells[r, 4].Value = row.Amount;
        }

        package.Save();
    }

    return checked((int)stream.Length);
}

static int CopyWorksheet(byte[] workbookBytes) {
    using var input = new MemoryStream(workbookBytes, writable: false);
    using var sourcePackage = new ExcelPackage(input);
    var sourceWorksheet = sourcePackage.Workbook.Worksheets["Data"]
        ?? throw new InvalidOperationException("Source worksheet 'Data' was not found.");

    using var output = new MemoryStream();
    using (var targetPackage = new ExcelPackage(output)) {
        targetPackage.Workbook.Worksheets.Add("CopiedData", sourceWorksheet);
        targetPackage.Save();
    }

    return checked((int)output.Length);
}

static int ReadRange(byte[] workbookBytes, int maxDataRows = int.MaxValue) {
    using var stream = new MemoryStream(workbookBytes, writable: false);
    using var package = new ExcelPackage(stream);
    var worksheet = package.Workbook.Worksheets["Data"];
    int lastRow = worksheet.Dimension?.End.Row ?? 0;
    if (maxDataRows != int.MaxValue) {
        lastRow = Math.Min(lastRow, maxDataRows + 1);
    }

    int metric = AddSalesHeadersMetric(0);

    for (int row = 2; row <= lastRow; row++) {
        metric = AddSalesRangeMetric(
            metric,
            Convert.ToInt32(worksheet.Cells[row, 1].Value, CultureInfo.InvariantCulture),
            Convert.ToString(worksheet.Cells[row, 2].Value, CultureInfo.InvariantCulture) ?? string.Empty,
            Convert.ToString(worksheet.Cells[row, 3].Value, CultureInfo.InvariantCulture) ?? string.Empty,
            ReadDateCell(worksheet.Cells[row, 4].Value),
            Convert.ToDouble(worksheet.Cells[row, 5].Value, CultureInfo.InvariantCulture),
            Convert.ToInt32(worksheet.Cells[row, 6].Value, CultureInfo.InvariantCulture),
            Convert.ToBoolean(worksheet.Cells[row, 7].Value, CultureInfo.InvariantCulture),
            Convert.ToString(worksheet.Cells[row, 8].Value, CultureInfo.InvariantCulture) ?? string.Empty);
    }

    return metric;
}

static int ReadDataTable(byte[] workbookBytes) {
    using var stream = new MemoryStream(workbookBytes, writable: false);
    using var package = new ExcelPackage(stream);
    var worksheet = package.Workbook.Worksheets["Data"];
    int lastRow = worksheet.Dimension?.End.Row ?? 0;
    DataTable table = CreateSalesDataTable();

    for (int row = 2; row <= lastRow; row++) {
        table.Rows.Add(
            Convert.ToInt32(worksheet.Cells[row, 1].Value, CultureInfo.InvariantCulture),
            Convert.ToString(worksheet.Cells[row, 2].Value, CultureInfo.InvariantCulture) ?? string.Empty,
            Convert.ToString(worksheet.Cells[row, 3].Value, CultureInfo.InvariantCulture) ?? string.Empty,
            ReadDateCell(worksheet.Cells[row, 4].Value),
            Convert.ToDouble(worksheet.Cells[row, 5].Value, CultureInfo.InvariantCulture),
            Convert.ToInt32(worksheet.Cells[row, 6].Value, CultureInfo.InvariantCulture),
            Convert.ToBoolean(worksheet.Cells[row, 7].Value, CultureInfo.InvariantCulture),
            Convert.ToString(worksheet.Cells[row, 8].Value, CultureInfo.InvariantCulture) ?? string.Empty);
    }

    return AddSalesDataTableMetric(table);
}

static int ReadSparseColumn(byte[] workbookBytes, int expectedRows) {
    using var stream = new MemoryStream(workbookBytes, writable: false);
    using var package = new ExcelPackage(stream);
    var worksheet = package.Workbook.Worksheets["Data"];
    int metric = 0;

    for (int row = 1; row <= expectedRows; row++) {
        metric = AddSparseMetric(metric, row, expectedRows, worksheet.Cells[row, 1].Text);
    }

    return metric;
}

static int ReadObjects(byte[] workbookBytes) {
    using var stream = new MemoryStream(workbookBytes, writable: false);
    using var package = new ExcelPackage(stream);
    var worksheet = package.Workbook.Worksheets["Data"];
    int lastRow = worksheet.Dimension?.End.Row ?? 0;
    int metric = 0;

    for (int row = 2; row <= lastRow; row++) {
        var record = new ReadSalesRecord {
            Id = Convert.ToInt32(worksheet.Cells[row, 1].Value, CultureInfo.InvariantCulture),
            Region = Convert.ToString(worksheet.Cells[row, 2].Value, CultureInfo.InvariantCulture) ?? string.Empty,
            Owner = Convert.ToString(worksheet.Cells[row, 3].Value, CultureInfo.InvariantCulture) ?? string.Empty,
            CreatedOn = ReadDateCell(worksheet.Cells[row, 4].Value),
            Amount = Convert.ToDouble(worksheet.Cells[row, 5].Value, CultureInfo.InvariantCulture),
            Units = Convert.ToInt32(worksheet.Cells[row, 6].Value, CultureInfo.InvariantCulture),
            Active = Convert.ToBoolean(worksheet.Cells[row, 7].Value, CultureInfo.InvariantCulture),
            Notes = Convert.ToString(worksheet.Cells[row, 8].Value, CultureInfo.InvariantCulture) ?? string.Empty
        };
        metric = AddSalesRecordMetric(metric, record);
    }

    return metric;
}

static int AutoFitExisting(byte[] workbookBytes) {
    using var input = new MemoryStream(workbookBytes, writable: false);
    using var output = new MemoryStream();
    using (var package = new ExcelPackage(input)) {
        var worksheet = package.Workbook.Worksheets["Data"];
        if (worksheet.Dimension != null) {
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
        }

        package.SaveAs(output);
    }

    return checked((int)output.Length);
}

static int WriteRealWorldReport(IReadOnlyList<SalesRecord> rows) {
    using var stream = new MemoryStream();
    using (var package = new ExcelPackage(stream)) {
        var worksheet = package.Workbook.Worksheets.Add("Data");
        PopulateWorksheet(worksheet, rows, includeTable: true, autoFit: true);
        ApplyNavigation(worksheet, rows.Count);
        ApplyConditionalFormatting(worksheet, rows.Count);
        ApplyDataValidation(worksheet, rows.Count);
        AddPivotTable(package, worksheet, rows.Count);
        AddRegionalChart(package, rows);
        package.Save();
    }

    return checked((int)stream.Length);
}

static int WriteRealWorldCoreReport(IReadOnlyList<SalesRecord> rows) {
    using var stream = new MemoryStream();
    using (var package = new ExcelPackage(stream)) {
        var worksheet = package.Workbook.Worksheets.Add("Data");
        PopulateWorksheet(worksheet, rows, includeTable: true, autoFit: true);
        ApplyNavigation(worksheet, rows.Count);
        ApplyConditionalFormatting(worksheet, rows.Count);
        ApplyDataValidation(worksheet, rows.Count);
        package.Save();
    }

    return checked((int)stream.Length);
}

static int WriteRealWorldFreezePanes(IReadOnlyList<SalesRecord> rows) {
    using var stream = new MemoryStream();
    using (var package = new ExcelPackage(stream)) {
        var worksheet = package.Workbook.Worksheets.Add("Data");
        PopulateWorksheet(worksheet, rows, includeTable: false, autoFit: false);
        worksheet.View.FreezePanes(2, 2);
        package.Save();
    }

    return checked((int)stream.Length);
}

static int WriteRealWorldAutoFilter(IReadOnlyList<SalesRecord> rows) {
    using var stream = new MemoryStream();
    using (var package = new ExcelPackage(stream)) {
        var worksheet = package.Workbook.Worksheets.Add("Data");
        PopulateWorksheet(worksheet, rows, includeTable: false, autoFit: false);
        worksheet.Cells[1, 1, rows.Count + 1, 8].AutoFilter = true;
        package.Save();
    }

    return checked((int)stream.Length);
}

static int WriteRealWorldConditionalFormatting(IReadOnlyList<SalesRecord> rows) {
    using var stream = new MemoryStream();
    using (var package = new ExcelPackage(stream)) {
        var worksheet = package.Workbook.Worksheets.Add("Data");
        PopulateWorksheet(worksheet, rows, includeTable: false, autoFit: false);
        ApplyConditionalFormatting(worksheet, rows.Count);
        package.Save();
    }

    return checked((int)stream.Length);
}

static int WriteRealWorldDataValidation(IReadOnlyList<SalesRecord> rows) {
    using var stream = new MemoryStream();
    using (var package = new ExcelPackage(stream)) {
        var worksheet = package.Workbook.Worksheets.Add("Data");
        PopulateWorksheet(worksheet, rows, includeTable: false, autoFit: false);
        ApplyDataValidation(worksheet, rows.Count);
        package.Save();
    }

    return checked((int)stream.Length);
}

static int WriteRealWorldCharts(IReadOnlyList<SalesRecord> rows) {
    using var stream = new MemoryStream();
    using (var package = new ExcelPackage(stream)) {
        var worksheet = package.Workbook.Worksheets.Add("Data");
        PopulateWorksheet(worksheet, rows, includeTable: false, autoFit: false);
        AddRegionalChart(package, rows);
        package.Save();
    }

    return checked((int)stream.Length);
}

static int WriteRealWorldPivotTable(IReadOnlyList<SalesRecord> rows) {
    using var stream = new MemoryStream();
    using (var package = new ExcelPackage(stream)) {
        var worksheet = package.Workbook.Worksheets.Add("Data");
        PopulateWorksheet(worksheet, rows, includeTable: false, autoFit: false);
        AddPivotTable(package, worksheet, rows.Count);
        package.Save();
    }

    return checked((int)stream.Length);
}

static int WriteRealWorldVariant(
    IReadOnlyList<SalesRecord> rows,
    IReadOnlyList<RealWorldColumnSpec> columns,
    RealWorldVariantOptions options) {
    using var stream = new MemoryStream();
    using (var package = new ExcelPackage(stream)) {
        var worksheet = package.Workbook.Worksheets.Add("Data");
        WriteVariantRows(worksheet, rows, columns);
        ApplyVariantTable(worksheet, rows.Count, columns.Count, options.AutoFit);
        ApplyVariantNavigation(worksheet, rows.Count, columns.Count);
        ApplyVariantConditionalFormatting(worksheet, rows.Count, columns);
        ApplyVariantDataValidation(worksheet, rows.Count, columns);

        if (options.ChartBeforePivot) {
            AddRegionalChart(package, rows);
            AddVariantPivotTable(package, worksheet, rows.Count, columns.Count);
        } else {
            AddVariantPivotTable(package, worksheet, rows.Count, columns.Count);
            AddRegionalChart(package, rows);
        }

        if (options.PostMutation) {
            worksheet.Cells[rows.Count + 4, 1].Value = "Manual note after report features";
        }

        package.Save();
    }

    return checked((int)stream.Length);
}

static int WriteReportWorkbook(IReadOnlyList<Dictionary<string, object?>> rows) {
    using var stream = new MemoryStream();
    using (var package = new ExcelPackage(stream)) {
        var worksheet = package.Workbook.Worksheets.Add("Data");
        PopulateReportWorkbookData(worksheet, rows);
        ApplyReportWorkbookCore(worksheet, rows.Count);
        AddReportWorkbookChart(worksheet, rows.Count);
        AddReportWorkbookPivotTable(worksheet, rows.Count);
        package.Save();
    }

    return checked((int)stream.Length);
}

static int WriteReportWorkbookCore(IReadOnlyList<Dictionary<string, object?>> rows) {
    using var stream = new MemoryStream();
    using (var package = new ExcelPackage(stream)) {
        var worksheet = package.Workbook.Worksheets.Add("Data");
        PopulateReportWorkbookData(worksheet, rows);
        ApplyReportWorkbookCore(worksheet, rows.Count);
        package.Save();
    }

    return checked((int)stream.Length);
}

static int WriteReportWorkbookDataTable(DataTable dataTable) {
    using var stream = new MemoryStream();
    using (var package = new ExcelPackage(stream)) {
        var worksheet = package.Workbook.Worksheets.Add("Data");
        PopulateReportWorkbookDataTable(worksheet, dataTable);
        ApplyReportWorkbookCore(worksheet, dataTable.Rows.Count);
        AddReportWorkbookChart(worksheet, dataTable.Rows.Count);
        AddReportWorkbookPivotTable(worksheet, dataTable.Rows.Count);
        package.Save();
    }

    return checked((int)stream.Length);
}

static int WriteReportWorkbookDataTableCore(DataTable dataTable) {
    using var stream = new MemoryStream();
    using (var package = new ExcelPackage(stream)) {
        var worksheet = package.Workbook.Worksheets.Add("Data");
        PopulateReportWorkbookDataTable(worksheet, dataTable);
        ApplyReportWorkbookCore(worksheet, dataTable.Rows.Count);
        package.Save();
    }

    return checked((int)stream.Length);
}

static int WriteSharedStrings(int rowCount) {
    using var stream = new MemoryStream();
    using (var package = new ExcelPackage(stream)) {
        var worksheet = package.Workbook.Worksheets.Add("Strings");
        for (int row = 1; row <= rowCount; row++) {
            worksheet.Cells[row, 1].Value = "Repeated value " + (row % 12);
            worksheet.Cells[row, 2].Value = "Distinct value " + row.ToString(CultureInfo.InvariantCulture);
            worksheet.Cells[row, 3].Value = "Long segment " + new string((char)('A' + (row % 26)), 48);
        }

        package.Save();
    }

    return checked((int)stream.Length);
}

static int WriteCellValueNumbers(int rowCount) {
    using var stream = new MemoryStream();
    using (var package = new ExcelPackage(stream)) {
        var worksheet = package.Workbook.Worksheets.Add("Numbers");
        for (int row = 1; row <= rowCount; row++) {
            worksheet.Cells[row, 1].Value = row * 1.25d;
            worksheet.Cells[row, 2].Value = row + 0.5d;
            worksheet.Cells[row, 3].Value = row % 17;
        }

        package.Save();
    }

    return checked((int)stream.Length);
}

static int WriteCellValueScalars(int rowCount) {
    using var stream = new MemoryStream();
    using (var package = new ExcelPackage(stream)) {
        var worksheet = package.Workbook.Worksheets.Add("Scalars");
        for (int row = 1; row <= rowCount; row++) {
            worksheet.Cells[row, 1].Value = row * 10.75m;
            worksheet.Cells[row, 2].Value = row % 2 == 0;
            worksheet.Cells[row, 3].Value = row % 3 == 0;
        }

        package.Save();
    }

    return checked((int)stream.Length);
}

static int WriteCellValueTemporal(int rowCount) {
    using var stream = new MemoryStream();
    using (var package = new ExcelPackage(stream)) {
        var worksheet = package.Workbook.Worksheets.Add("Temporal");
        var start = new DateTime(2026, 1, 1, 8, 30, 0, DateTimeKind.Unspecified);
        for (int row = 1; row <= rowCount; row++) {
            worksheet.Cells[row, 1].Value = start.AddDays(row);
            worksheet.Cells[row, 2].Value = TimeSpan.FromMinutes(row * 7);
            worksheet.Cells[row, 3].Value = start.AddHours(row % 24);
        }

        package.Save();
    }

    return checked((int)stream.Length);
}

static int WriteCellValueObjectMixed(int rowCount) {
    using var stream = new MemoryStream();
    using (var package = new ExcelPackage(stream)) {
        var worksheet = package.Workbook.Worksheets.Add("Objects");
        var start = new DateTime(2026, 1, 1, 8, 30, 0, DateTimeKind.Unspecified);
        for (int row = 1; row <= rowCount; row++) {
            object? name = "Item " + (row % 12).ToString(CultureInfo.InvariantCulture);
            object? amount = (double)row * 1.25d;
            object? active = row % 2 == 0;
            object? created = start.AddDays(row);
            worksheet.Cells[row, 1].Value = name;
            worksheet.Cells[row, 2].Value = amount;
            worksheet.Cells[row, 3].Value = active;
            worksheet.Cells[row, 4].Value = created;
        }

        package.Save();
    }

    return checked((int)stream.Length);
}

static int WriteCellFormula(int rowCount) {
    using var stream = new MemoryStream();
    using (var package = new ExcelPackage(stream)) {
        var worksheet = package.Workbook.Worksheets.Add("Formulas");
        for (int row = 1; row <= rowCount; row++) {
            worksheet.Cells[row, 1].Value = (double)row;
            worksheet.Cells[row, 2].Value = (double)(row % 17);
            worksheet.Cells[row, 3].Value = (double)(row % 29);
            worksheet.Cells[row, 4].Formula = $"SUM(A{row}:C{row})";
        }

        package.Save();
    }

    return checked((int)stream.Length);
}

static int ReadFormulaText(byte[] workbookBytes, int rowCount) {
    using var stream = new MemoryStream(workbookBytes, writable: false);
    using var package = new ExcelPackage(stream);
    var worksheet = package.Workbook.Worksheets["Formulas"];
    int metric = 0;
    for (int row = 2; row <= rowCount + 1; row++) {
        metric = AddStringMetric(metric, worksheet.Cells[row, 4].Formula);
    }

    return metric;
}

static int ReadSharedStrings(byte[] workbookBytes, int rowCount) {
    using var stream = new MemoryStream(workbookBytes, writable: false);
    using var package = new ExcelPackage(stream);
    var worksheet = package.Workbook.Worksheets["Strings"];
    int metric = 0;
    for (int row = 1; row <= rowCount; row++) {
        for (int col = 1; col <= 3; col++) {
            metric = AddStringMetric(metric, worksheet.Cells[row, col].Text);
        }
    }

    return metric;
}

static byte[] CreateWorkbookBytes(IReadOnlyList<SalesRecord> rows) {
    using var stream = new MemoryStream();
    using (var package = new ExcelPackage(stream)) {
        var worksheet = package.Workbook.Worksheets.Add("Data");
        PopulateWorksheet(worksheet, rows, includeTable: true, autoFit: true);
        package.Save();
    }

    return stream.ToArray();
}

static byte[] CreateFormulaWorkbookBytes(int rowCount) {
    using var stream = new MemoryStream();
    using (var package = new ExcelPackage(stream)) {
        var worksheet = package.Workbook.Worksheets.Add("Formulas");
        worksheet.Cells[1, 1].Value = "A";
        worksheet.Cells[1, 2].Value = "B";
        worksheet.Cells[1, 3].Value = "C";
        worksheet.Cells[1, 4].Value = "Total";
        for (int row = 2; row <= rowCount + 1; row++) {
            worksheet.Cells[row, 1].Value = (double)row;
            worksheet.Cells[row, 2].Value = (double)(row * 2);
            worksheet.Cells[row, 3].Value = (double)(row * 3);
            worksheet.Cells[row, 4].Formula = $"SUM(A{row}:C{row})";
        }

        package.Save();
    }

    return stream.ToArray();
}

static byte[] CreateSharedStringWorkbookBytes(int rowCount) {
    using var stream = new MemoryStream();
    using (var package = new ExcelPackage(stream)) {
        var worksheet = package.Workbook.Worksheets.Add("Strings");
        for (int row = 1; row <= rowCount; row++) {
            worksheet.Cells[row, 1].Value = "Repeated value " + (row % 12);
            worksheet.Cells[row, 2].Value = "Distinct value " + row.ToString(CultureInfo.InvariantCulture);
            worksheet.Cells[row, 3].Value = "Long segment " + new string((char)('A' + (row % 26)), 48);
        }

        package.Save();
    }

    return stream.ToArray();
}

static byte[] CreateSparseWorkbookBytes(int lastRow) {
    using var stream = new MemoryStream();
    using (var package = new ExcelPackage(stream)) {
        var worksheet = package.Workbook.Worksheets.Add("Data");
        worksheet.Cells[1, 1].Value = "Header";
        worksheet.Cells[lastRow, 1].Value = "Tail";
        package.Save();
    }

    return stream.ToArray();
}

static void PopulateWorksheet(ExcelWorksheet worksheet, IReadOnlyList<SalesRecord> rows, bool includeTable, bool autoFit) {
    WriteHeaders(worksheet);
    for (int i = 0; i < rows.Count; i++) {
        var row = rows[i];
        int r = i + 2;
        worksheet.Cells[r, 1].Value = row.Id;
        worksheet.Cells[r, 2].Value = row.Region;
        worksheet.Cells[r, 3].Value = row.Owner;
        worksheet.Cells[r, 4].Value = row.CreatedOn;
        worksheet.Cells[r, 5].Value = row.Amount;
        worksheet.Cells[r, 6].Value = row.Units;
        worksheet.Cells[r, 7].Value = row.Active;
        worksheet.Cells[r, 8].Value = row.Notes;
    }

    if (includeTable) {
        string tableName = string.Equals(worksheet.Name, "Data", StringComparison.OrdinalIgnoreCase) ? "SalesData" : worksheet.Name;
        var table = worksheet.Tables.Add(worksheet.Cells[1, 1, rows.Count + 1, 8], tableName);
        table.TableStyle = TableStyles.Medium2;
    }

    if (autoFit && worksheet.Dimension != null) {
        worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
    }
}

static void ApplyNavigation(ExcelWorksheet worksheet, int rowCount) {
    worksheet.View.FreezePanes(2, 2);
    worksheet.Cells[1, 1, rowCount + 1, 8].AutoFilter = true;
}

static void ApplyConditionalFormatting(ExcelWorksheet worksheet, int rowCount) {
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

static void ApplyDataValidation(ExcelWorksheet worksheet, int rowCount) {
    int lastRow = rowCount + 1;
    var validation = worksheet.DataValidations.AddIntegerValidation($"F2:F{lastRow.ToString(CultureInfo.InvariantCulture)}");
    validation.Operator = OfficeOpenXml.DataValidation.ExcelDataValidationOperator.between;
    validation.Formula.Value = 1;
    validation.Formula2.Value = 24;
}

static void WriteVariantRows(ExcelWorksheet worksheet, IReadOnlyList<SalesRecord> rows, IReadOnlyList<RealWorldColumnSpec> columns) {
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

static void ApplyVariantTable(ExcelWorksheet worksheet, int rowCount, int columnCount, bool autoFit) {
    var table = worksheet.Tables.Add(worksheet.Cells[1, 1, rowCount + 1, columnCount], "SalesData");
    table.TableStyle = TableStyles.Medium2;
    if (autoFit) {
        worksheet.Cells[1, 1, rowCount + 1, columnCount].AutoFitColumns();
    }
}

static void ApplyVariantNavigation(ExcelWorksheet worksheet, int rowCount, int columnCount) {
    worksheet.View.FreezePanes(2, 2);
    worksheet.Cells[1, 1, rowCount + 1, columnCount].AutoFilter = true;
}

static void ApplyVariantConditionalFormatting(ExcelWorksheet worksheet, int rowCount, IReadOnlyList<RealWorldColumnSpec> columns) {
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

static void ApplyVariantDataValidation(ExcelWorksheet worksheet, int rowCount, IReadOnlyList<RealWorldColumnSpec> columns) {
    int unitsColumn = GetColumnIndex(columns, "Units");
    int lastRow = rowCount + 1;
    var validation = worksheet.DataValidations.AddIntegerValidation(BuildColumnRange(unitsColumn, 2, lastRow));
    validation.Operator = OfficeOpenXml.DataValidation.ExcelDataValidationOperator.between;
    validation.Formula.Value = 1;
    validation.Formula2.Value = 24;
}

static void AddVariantPivotTable(ExcelPackage package, ExcelWorksheet dataWorksheet, int rowCount, int columnCount) {
    var pivotSheet = package.Workbook.Worksheets.Add("Pivot");
    var source = dataWorksheet.Cells[1, 1, rowCount + 1, columnCount];
    var pivot = pivotSheet.PivotTables.Add(pivotSheet.Cells["A3"], source, "SalesPivot");
    pivot.RowFields.Add(pivot.Fields["Region"]);
    pivot.ColumnFields.Add(pivot.Fields["Owner"]);
    var amount = pivot.DataFields.Add(pivot.Fields["Amount"]);
    amount.Function = DataFieldFunctions.Sum;
    amount.Name = "Total Amount";
}

static string BuildColumnRange(int columnIndex, int firstRow, int lastRow) {
    string column = GetColumnLetter(columnIndex);
    return column + firstRow.ToString(CultureInfo.InvariantCulture) + ":" + column + lastRow.ToString(CultureInfo.InvariantCulture);
}

static int GetColumnIndex(IReadOnlyList<RealWorldColumnSpec> columns, string header) {
    for (int i = 0; i < columns.Count; i++) {
        if (string.Equals(columns[i].Header, header, StringComparison.OrdinalIgnoreCase)) {
            return i + 1;
        }
    }

    throw new InvalidOperationException($"Column '{header}' was not found in the benchmark variant.");
}

static string GetColumnLetter(int columnIndex) {
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

static void AddPivotTable(ExcelPackage package, ExcelWorksheet dataWorksheet, int rowCount) {
    var pivotSheet = package.Workbook.Worksheets.Add("Pivot");
    var source = dataWorksheet.Cells[1, 1, rowCount + 1, 8];
    var pivot = pivotSheet.PivotTables.Add(pivotSheet.Cells["A3"], source, "SalesPivot");
    pivot.RowFields.Add(pivot.Fields["Region"]);
    pivot.ColumnFields.Add(pivot.Fields["Owner"]);
    var amount = pivot.DataFields.Add(pivot.Fields["Amount"]);
    amount.Function = DataFieldFunctions.Sum;
    amount.Name = "Total Amount";
}

static void AddRegionalChart(ExcelPackage package, IReadOnlyList<SalesRecord> rows) {
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

static void PopulateReportWorkbookData(ExcelWorksheet worksheet, IReadOnlyList<Dictionary<string, object?>> rows) {
    WriteReportWorkbookRows(worksheet, rows);
    var table = worksheet.Tables.Add(worksheet.Cells[1, 1, rows.Count + 1, 10], "Data");
    table.TableStyle = TableStyles.Medium2;
    if (worksheet.Dimension != null) {
        worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
    }
}

static void PopulateReportWorkbookDataTable(ExcelWorksheet worksheet, DataTable dataTable) {
    worksheet.Cells["A1"].LoadFromDataTable(dataTable, true);
    var table = worksheet.Tables.Add(worksheet.Cells[1, 1, dataTable.Rows.Count + 1, dataTable.Columns.Count], "Data");
    table.TableStyle = TableStyles.Medium2;
    if (worksheet.Dimension != null) {
        worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
    }
}

static void ApplyReportWorkbookCore(ExcelWorksheet worksheet, int rowCount) {
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

static void AddReportWorkbookChart(ExcelWorksheet worksheet, int rowCount) {
    int lastRow = rowCount + 1;
    var chart = worksheet.Drawings.AddChart("ReportScoreChart", eChartType.ColumnClustered);
    chart.Title.Text = "Score by Created";
    chart.SetPosition(1, 0, 11, 0);
    chart.SetSize(720, 320);
    chart.Series.Add(worksheet.Cells[2, 7, lastRow, 7], worksheet.Cells[2, 6, lastRow, 6]);
}

static void AddReportWorkbookPivotTable(ExcelWorksheet worksheet, int rowCount) {
    var source = worksheet.Cells[1, 1, rowCount + 1, 10];
    var pivot = worksheet.PivotTables.Add(worksheet.Cells["L24"], source, "ReportPivot");
    pivot.RowFields.Add(pivot.Fields["Region"]);
    pivot.ColumnFields.Add(pivot.Fields["Department"]);
    var score = pivot.DataFields.Add(pivot.Fields["Score"]);
    score.Function = DataFieldFunctions.Average;
    score.Name = "Average Score";
    var tickets = pivot.DataFields.Add(pivot.Fields["TicketCount"]);
    tickets.Function = DataFieldFunctions.Sum;
    tickets.Name = "Sum TicketCount";
}

static void WriteReportWorkbookRows(ExcelWorksheet worksheet, IReadOnlyList<Dictionary<string, object?>> rows) {
    string[] columns = ["Id", "Name", "Department", "Region", "IsEnabled", "Created", "Score", "Owner", "TicketCount", "Notes"];
    for (int i = 0; i < columns.Length; i++) {
        worksheet.Cells[1, i + 1].Value = columns[i];
    }

    for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
        Dictionary<string, object?> row = rows[rowIndex];
        int targetRow = rowIndex + 2;
        for (int columnIndex = 0; columnIndex < columns.Length; columnIndex++) {
            row.TryGetValue(columns[columnIndex], out object? value);
            worksheet.Cells[targetRow, columnIndex + 1].Value = value;
        }
    }
}

static IReadOnlyList<RegionSummary> BuildRegionSummaries(IReadOnlyList<SalesRecord> rows)
    => rows
        .GroupBy(static row => row.Region, StringComparer.Ordinal)
        .OrderBy(static group => group.Key, StringComparer.Ordinal)
        .Select(static group => new RegionSummary(
            group.Key,
            Math.Round(group.Sum(static row => row.Amount), 2),
            group.Sum(static row => row.Units)))
        .ToArray();

static void WriteHeaders(ExcelWorksheet worksheet) {
    worksheet.Cells[1, 1].Value = "Id";
    worksheet.Cells[1, 2].Value = "Region";
    worksheet.Cells[1, 3].Value = "Owner";
    worksheet.Cells[1, 4].Value = "CreatedOn";
    worksheet.Cells[1, 5].Value = "Amount";
    worksheet.Cells[1, 6].Value = "Units";
    worksheet.Cells[1, 7].Value = "Active";
    worksheet.Cells[1, 8].Value = "Notes";
}

static void WriteAppendHeaders(ExcelWorksheet worksheet) {
    worksheet.Cells[1, 1].Value = "Id";
    worksheet.Cells[1, 2].Value = "Region";
    worksheet.Cells[1, 3].Value = "Owner";
    worksheet.Cells[1, 4].Value = "Amount";
}

static IReadOnlyList<SalesRecord> CreateSalesRecords(int count) {
    string[] regions = ["North", "South", "East", "West", "Central"];
    string[] owners = ["Ava", "Noah", "Mia", "Liam", "Zoe", "Ethan", "Ivy", "Mason"];
    var records = new List<SalesRecord>(count);
    var start = new DateTime(2024, 1, 1, 8, 0, 0, DateTimeKind.Unspecified);

    for (int i = 0; i < count; i++) {
        records.Add(new SalesRecord {
            Id = i + 1,
            Region = regions[i % regions.Length],
            Owner = owners[i % owners.Length],
            CreatedOn = start.AddDays(i % 365).AddMinutes(i % 180),
            Amount = Math.Round(150 + ((i * 17.35) % 4500), 2),
            Units = 1 + (i % 24),
            Active = i % 3 != 0,
            Notes = $"Batch {(i % 12) + 1} / segment {(i % 7) + 1}"
        });
    }

    return records;
}

static IReadOnlyList<Dictionary<string, object?>> CreatePowerShellMixedRows(int count) {
    string[] regions = ["NA", "EU", "APAC", "LATAM"];
    var result = new List<Dictionary<string, object?>>(count);
    var start = new DateTime(2024, 1, 1, 8, 0, 0, DateTimeKind.Unspecified);
    for (int i = 0; i < count; i++) {
        int id = i + 1;
        result.Add(new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase) {
            ["Id"] = id,
            ["Name"] = "Server-" + id.ToString("D6", CultureInfo.InvariantCulture),
            ["Department"] = "Department-" + (id % 12 + 1).ToString(CultureInfo.InvariantCulture),
            ["Region"] = regions[i % regions.Length],
            ["IsEnabled"] = id % 4 != 0,
            ["Created"] = start.AddDays(i % 365).AddMinutes(i % 240),
            ["Score"] = Math.Round(100D + ((id * 17.456D) % 900D), 3),
            ["Owner"] = "owner" + (id % 128).ToString(CultureInfo.InvariantCulture) + "@example.test",
            ["TicketCount"] = id % 17,
            ["Notes"] = "Benchmark row " + id.ToString(CultureInfo.InvariantCulture)
        });
    }

    return result;
}

static DataTable CreatePowerShellMixedDataTable(IEnumerable<IReadOnlyDictionary<string, object?>> rows, string tableName) {
    string[] columns = ["Id", "Name", "Department", "Region", "IsEnabled", "Created", "Score", "Owner", "TicketCount", "Notes"];
    var table = new DataTable(tableName) { Locale = CultureInfo.InvariantCulture };
    table.Columns.Add("Id", typeof(int));
    table.Columns.Add("Name", typeof(string));
    table.Columns.Add("Department", typeof(string));
    table.Columns.Add("Region", typeof(string));
    table.Columns.Add("IsEnabled", typeof(bool));
    table.Columns.Add("Created", typeof(DateTime));
    table.Columns.Add("Score", typeof(double));
    table.Columns.Add("Owner", typeof(string));
    table.Columns.Add("TicketCount", typeof(int));
    table.Columns.Add("Notes", typeof(string));

    foreach (IReadOnlyDictionary<string, object?> sourceRow in rows) {
        object?[] values = new object?[columns.Length];
        for (int i = 0; i < values.Length; i++) {
            values[i] = sourceRow.TryGetValue(columns[i], out object? value)
                ? value
                : DBNull.Value;
        }

        table.Rows.Add(values);
    }

    return table;
}

static DateTime ReadDateCell(object? value)
    => value is DateTime dateTime
        ? dateTime
        : DateTime.FromOADate(Convert.ToDouble(value, CultureInfo.InvariantCulture));

static int AddSalesRecordMetric(int metric, ReadSalesRecord record) {
    metric = AddIntMetric(metric, record.Id);
    metric = AddStringMetric(metric, record.Region);
    metric = AddStringMetric(metric, record.Owner);
    metric = AddIntMetric(metric, record.CreatedOn.DayOfYear);
    metric = AddDoubleMetric(metric, record.Amount);
    metric = AddIntMetric(metric, record.Units);
    metric = AddIntMetric(metric, record.Active ? 1 : 0);
    return AddStringMetric(metric, record.Notes);
}

static int AddSalesHeadersMetric(int metric) {
    metric = AddStringMetric(metric, "Id");
    metric = AddStringMetric(metric, "Region");
    metric = AddStringMetric(metric, "Owner");
    metric = AddStringMetric(metric, "CreatedOn");
    metric = AddStringMetric(metric, "Amount");
    metric = AddStringMetric(metric, "Units");
    metric = AddStringMetric(metric, "Active");
    return AddStringMetric(metric, "Notes");
}

static int AddSalesRangeMetric(
    int metric,
    int id,
    string region,
    string owner,
    DateTime createdOn,
    double amount,
    int units,
    bool active,
    string notes) {
    metric = AddIntMetric(metric, id);
    metric = AddStringMetric(metric, region);
    metric = AddStringMetric(metric, owner);
    metric = AddIntMetric(metric, createdOn.DayOfYear);
    metric = AddDoubleMetric(metric, amount);
    metric = AddIntMetric(metric, units);
    metric = AddIntMetric(metric, active ? 1 : 0);
    return AddStringMetric(metric, notes);
}

static DataTable CreateSalesDataTable() {
    var table = new DataTable("Data") { Locale = CultureInfo.InvariantCulture };
    table.Columns.Add("Id", typeof(int));
    table.Columns.Add("Region", typeof(string));
    table.Columns.Add("Owner", typeof(string));
    table.Columns.Add("CreatedOn", typeof(DateTime));
    table.Columns.Add("Amount", typeof(double));
    table.Columns.Add("Units", typeof(int));
    table.Columns.Add("Active", typeof(bool));
    table.Columns.Add("Notes", typeof(string));
    return table;
}

static DataSet CreateSalesDataSet(IReadOnlyList<SalesRecord> rows) {
    var dataSet = new DataSet("Sales") { Locale = CultureInfo.InvariantCulture };
    dataSet.Tables.Add(CreateSalesDataTableFromRows(rows.Take(rows.Count / 2), "SalesA"));
    dataSet.Tables.Add(CreateSalesDataTableFromRows(rows.Skip(rows.Count / 2), "SalesB"));
    return dataSet;
}

static DataTable CreateSalesDataTableFromRows(IEnumerable<SalesRecord> rows, string tableName) {
    var table = CreateSalesDataTable();
    table.TableName = tableName;
    foreach (var row in rows) {
        table.Rows.Add(row.Id, row.Region, row.Owner, row.CreatedOn, row.Amount, row.Units, row.Active, row.Notes);
    }

    return table;
}

static int AddSalesDataTableMetric(DataTable table) {
    int metric = 0;
    foreach (DataColumn column in table.Columns) {
        metric = AddStringMetric(metric, column.ColumnName);
    }

    foreach (DataRow row in table.Rows) {
        metric = AddSalesRangeMetric(
            metric,
            Convert.ToInt32(row[0], CultureInfo.InvariantCulture),
            Convert.ToString(row[1], CultureInfo.InvariantCulture) ?? string.Empty,
            Convert.ToString(row[2], CultureInfo.InvariantCulture) ?? string.Empty,
            ReadDateCell(row[3]),
            Convert.ToDouble(row[4], CultureInfo.InvariantCulture),
            Convert.ToInt32(row[5], CultureInfo.InvariantCulture),
            Convert.ToBoolean(row[6], CultureInfo.InvariantCulture),
            Convert.ToString(row[7], CultureInfo.InvariantCulture) ?? string.Empty);
    }

    return metric;
}

static int AddIntMetric(int metric, int value) {
    unchecked {
        return (metric * 397) ^ value;
    }
}

static int AddDoubleMetric(int metric, double value) {
    unchecked {
        return AddIntMetric(metric, (int)Math.Round(value * 100, MidpointRounding.AwayFromZero));
    }
}

static int AddStringMetric(int metric, string? value) {
    unchecked {
        int result = metric;
        if (value == null) {
            return result * 397;
        }

        for (int i = 0; i < value.Length; i++) {
            result = (result * 397) ^ value[i];
        }

        return result;
    }
}

static int AddSparseMetric(int metric, int rowIndex, int expectedRows, object? value) {
    string? text = Convert.ToString(value, CultureInfo.InvariantCulture);
    if (string.IsNullOrEmpty(text)) {
        text = null;
    }

    if (rowIndex == 1) {
        if (!string.Equals(text, "Header", StringComparison.Ordinal)) {
            throw new InvalidOperationException("Sparse read did not return the first row value.");
        }

        return AddStringMetric(metric, text);
    }

    if (rowIndex == expectedRows) {
        if (!string.Equals(text, "Tail", StringComparison.Ordinal)) {
            throw new InvalidOperationException("Sparse read did not return the last row value.");
        }

        return AddStringMetric(metric, text);
    }

    if (text != null) {
        throw new InvalidOperationException($"Sparse read returned an unexpected value at row {rowIndex}.");
    }

    return AddStringMetric(metric, null);
}

static int ParseRowCount(string[] args) {
    for (int i = 0; i < args.Length; i++) {
        if (!string.Equals(args[i], "--rows", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(args[i], "--row-count", StringComparison.OrdinalIgnoreCase)) {
            continue;
        }

        if (i + 1 >= args.Length) {
            throw new ArgumentException("Missing value for --rows.");
        }

        string value = args[i + 1].Replace(",", string.Empty, StringComparison.Ordinal).Replace("_", string.Empty, StringComparison.Ordinal);
        if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int rowCount)
            || rowCount <= 0) {
            throw new ArgumentException("--rows must be a positive integer.");
        }

        return rowCount;
    }

    return DefaultRowCount;
}

static string? ParseOutputPath(string[] args)
    => ParseOptionValue(args, "--out", "--output", "--output-path")
       ?? (args.Length >= 1 && !args[0].StartsWith("-", StringComparison.Ordinal) ? args[0] : null);

static string? ParseOptionValue(string[] args, params string[] optionNames) {
    for (int i = 0; i < args.Length; i++) {
        if (!optionNames.Any(name => string.Equals(args[i], name, StringComparison.OrdinalIgnoreCase))) {
            continue;
        }

        if (i + 1 >= args.Length || args[i + 1].StartsWith("-", StringComparison.Ordinal)) {
            throw new ArgumentException($"Missing value for {args[i]}.");
        }

        return args[i + 1];
    }

    return null;
}

static string[] ParseOptionValues(string[] args, params string[] optionNames) {
    var values = new List<string>();
    for (int i = 0; i < args.Length; i++) {
        if (!optionNames.Any(name => string.Equals(args[i], name, StringComparison.OrdinalIgnoreCase))) {
            continue;
        }

        if (i + 1 >= args.Length || args[i + 1].StartsWith("-", StringComparison.Ordinal)) {
            throw new ArgumentException($"Missing value for {args[i]}.");
        }

        values.AddRange(args[i + 1]
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(value => value.Length > 0));
        i++;
    }

    return values.ToArray();
}

internal sealed class BenchmarkMeasurement {
    internal static BenchmarkMeasurementResult Measure(int warmupIterations, int measuredIterations, Func<int> action) {
        if (action == null) {
            throw new ArgumentNullException(nameof(action));
        }

        for (int i = 0; i < warmupIterations; i++) {
            action();
        }

        var elapsed = new List<double>(measuredIterations);
        int lastMetric = 0;

        for (int i = 0; i < measuredIterations; i++) {
            PrepareForMeasurement();

            var stopwatch = Stopwatch.StartNew();
            lastMetric = action();
            stopwatch.Stop();
            elapsed.Add(stopwatch.Elapsed.TotalMilliseconds);
        }

        return new BenchmarkMeasurementResult(lastMetric, elapsed);
    }

    private static void PrepareForMeasurement() {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
    }

    internal sealed class BenchmarkMeasurementResult {
        internal BenchmarkMeasurementResult(int outputMetric, IReadOnlyList<double> samplesMilliseconds) {
            OutputMetric = outputMetric;
            SamplesMilliseconds = samplesMilliseconds.ToArray();
        }

        internal int OutputMetric { get; }
        internal IReadOnlyList<double> SamplesMilliseconds { get; }
        internal double AverageMilliseconds => SamplesMilliseconds.Count == 0 ? 0 : SamplesMilliseconds.Average();
        internal double MedianMilliseconds {
            get {
                if (SamplesMilliseconds.Count == 0) {
                    return 0;
                }

                var ordered = SamplesMilliseconds.OrderBy(v => v).ToArray();
                int middle = ordered.Length / 2;
                if ((ordered.Length & 1) == 1) {
                    return ordered[middle];
                }

                return (ordered[middle - 1] + ordered[middle]) / 2.0;
            }
        }
    }
}

internal sealed class LegacyComparisonProfile {
    public DateTime GeneratedAtUtc { get; init; }
    public string Framework { get; init; } = string.Empty;
    public string MachineName { get; init; } = string.Empty;
    public int RowCount { get; init; }
    public int WarmupIterations { get; init; }
    public int MeasuredIterations { get; init; }
    public string Notes { get; init; } = string.Empty;
    public List<LegacyComparisonScenario> Scenarios { get; init; } = [];
}

internal sealed class LegacyComparisonScenario {
    public string Scenario { get; init; } = string.Empty;
    public string Library { get; init; } = string.Empty;
    public string Notes { get; init; } = string.Empty;
    public int OutputMetric { get; init; }
    public double AverageMilliseconds { get; init; }
    public double MedianMilliseconds { get; init; }
    public List<double> SamplesMilliseconds { get; init; } = [];
}

internal sealed class SalesRecord {
    public int Id { get; init; }
    public string Region { get; init; } = string.Empty;
    public string Owner { get; init; } = string.Empty;
    public DateTime CreatedOn { get; init; }
    public double Amount { get; init; }
    public int Units { get; init; }
    public bool Active { get; init; }
    public string Notes { get; init; } = string.Empty;
}

internal sealed record RegionSummary(string Region, double Amount, int Units);

internal sealed record RealWorldColumnSpec(string Header, Func<SalesRecord, object?> Selector);

internal sealed record RealWorldVariantOptions(bool AutoFit = true, bool ChartBeforePivot = false, bool PostMutation = false) {
    public static readonly RealWorldVariantOptions Default = new();
}

internal sealed class ReadSalesRecord {
    public int Id { get; set; }
    public string Region { get; set; } = string.Empty;
    public string Owner { get; set; } = string.Empty;
    public DateTime CreatedOn { get; set; }
    public double Amount { get; set; }
    public int Units { get; set; }
    public bool Active { get; set; }
    public string Notes { get; set; } = string.Empty;
}
