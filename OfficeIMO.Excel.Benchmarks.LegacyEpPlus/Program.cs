using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using OfficeOpenXml;
using OfficeOpenXml.Table;

const int DefaultRowCount = 2500;
const int SparseLastRow = 100_001;
const int WarmupIterations = 1;
const int MeasuredIterations = 3;
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
int topDataRows = Math.Min(rowCount, 100);
byte[] workbookBytes = CreateWorkbookBytes(rows);
byte[] formulaWorkbookBytes = CreateFormulaWorkbookBytes(rowCount);
byte[] sharedStringWorkbookBytes = CreateSharedStringWorkbookBytes(rowCount);
byte[] sparseWorkbookBytes = CreateSparseWorkbookBytes(SparseLastRow);
var scenarios = new List<LegacyComparisonScenario>();

AddScenario(scenarios, scenarioFilter, "write-bulk-report", LibraryName, "EPPlus 4.x manual row population, add table, autofit, save.", () => WriteBulkReport(rows), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "write-dataset-tables", LibraryName, "EPPlus 4.x import prepared DataTables as two styled worksheet tables and save.", () => WriteDataSetTables(salesDataSet), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "write-datatable-direct", LibraryName, "EPPlus 4.x import a prepared DataTable and save.", () => WriteDataTable(salesDataTable), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "write-datatable-table-direct", LibraryName, "EPPlus 4.x import a prepared DataTable as a styled worksheet table and save.", () => WriteDataTable(salesDataTable, includeTable: true), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "write-datareader-table-direct", LibraryName, "EPPlus 4.x import equivalent prepared data as a styled worksheet table and save.", () => WriteDataTable(salesDataTable, includeTable: true), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "write-cellvalues-rectangle-direct", LibraryName, "EPPlus 4.x write the same complete A1 rectangle and save.", () => WriteEquivalentSalesRows(rows), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "write-insertobjects-direct", LibraryName, "EPPlus 4.x import equivalent typed object data and save.", () => WriteDataTable(salesDataTable), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "write-fluent-rowsfrom-direct", LibraryName, "EPPlus 4.x import equivalent typed row data and save.", () => WriteDataTable(salesDataTable), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "append-plain-rows", LibraryName, "EPPlus 4.x append equivalent row/cell values.", () => AppendPlainRows(rows), warmupIterations, measuredIterations);
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
AddScenario(scenarios, scenarioFilter, "large-shared-strings", LibraryName, "EPPlus 4.x write repeated and distinct text-heavy cells.", () => WriteSharedStrings(rowCount), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "formula-heavy-read", LibraryName, "EPPlus 4.x read formula text from formula cells.", () => ReadFormulaText(formulaWorkbookBytes, rowCount), warmupIterations, measuredIterations);
AddScenario(scenarios, scenarioFilter, "shared-string-read", LibraryName, "EPPlus 4.x read repeated shared string payload.", () => ReadSharedStrings(sharedStringWorkbookBytes, rowCount), warmupIterations, measuredIterations);

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

static int WriteDataTable(DataTable dataTable, bool includeTable = false) {
    using var stream = new MemoryStream();
    using (var package = new ExcelPackage(stream)) {
        var worksheet = package.Workbook.Worksheets.Add("Data");
        worksheet.Cells["A1"].LoadFromDataTable(dataTable, true, includeTable ? TableStyles.Medium2 : TableStyles.None);
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
