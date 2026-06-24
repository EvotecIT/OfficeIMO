using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;

int rowCount = ParsePositiveOption(args, "--rows", "--row-count") ?? 2500;
int warmupIterations = ParsePositiveOption(args, "--warmup", "--warmups") ?? 1;
int measuredIterations = ParsePositiveOption(args, "--iterations", "--measured-iterations", "--samples") ?? 3;
string outputPath = ParseOptionValue(args, "--out", "--output", "--output-path")
    ?? Path.Combine("Docs", "benchmarks", "officeimo.excel.npoi-comparison.json");
string[] scenarioFilters = ParseOptionValues(args, "--scenario", "--scenarios");

if (HasSwitch(args, "--help") || HasSwitch(args, "-h") || HasSwitch(args, "/?")) {
    WriteUsage();
    return;
}

if (rowCount <= 0) {
    throw new ArgumentOutOfRangeException(nameof(rowCount));
}

if (warmupIterations <= 0) {
    throw new ArgumentOutOfRangeException(nameof(warmupIterations));
}

if (measuredIterations <= 0) {
    throw new ArgumentOutOfRangeException(nameof(measuredIterations));
}

var scenarioFilter = new HashSet<string>(scenarioFilters, StringComparer.OrdinalIgnoreCase);
bool IncludeScenario(string name) => scenarioFilter.Count == 0 || scenarioFilter.Contains(name);

var records = SalesRecord.Create(rowCount);
var npoiXlsx = new Lazy<byte[]>(() => WriteNpoiXlsx(records));
var npoiXls = new Lazy<byte[]>(() => WriteNpoiXls(records));
var npoiFormulaXls = new Lazy<byte[]>(() => WriteNpoiFormulaXls(records));

if ((IncludeScenario("xls-read-cellvalues") || IncludeScenario("xls-read-formulas")) && rowCount + 1 > 65_536) {
    throw new ArgumentOutOfRangeException(nameof(rowCount), "The xls scenarios cannot exceed the BIFF8 worksheet row limit.");
}

var measurements = new List<NpoiComparisonMeasurement>();

if (IncludeScenario("xlsx-write-cellvalues")) {
    AddScenario(measurements, "xlsx-write-cellvalues", "OfficeIMO.Excel", "Plain row/cell write to .xlsx through OfficeIMO CellValue.", () => WriteOfficeImoXlsx(records).Length, warmupIterations, measuredIterations);
    AddScenario(measurements, "xlsx-write-cellvalues", "NPOI XSSF", "Plain row/cell write to .xlsx through XSSFWorkbook.", () => WriteNpoiXlsx(records).Length, warmupIterations, measuredIterations);
}

if (IncludeScenario("xlsx-read-cellvalues")) {
    AddScenario(measurements, "xlsx-read-cellvalues", "OfficeIMO.Excel", "Plain row/cell read from an NPOI-generated .xlsx workbook.", () => ReadOfficeImoXlsx(npoiXlsx.Value, rowCount), warmupIterations, measuredIterations);
    AddScenario(measurements, "xlsx-read-cellvalues", "NPOI XSSF", "Plain row/cell read from the same NPOI-generated .xlsx workbook.", () => ReadNpoiWorkbook(npoiXlsx.Value, rowCount), warmupIterations, measuredIterations);
}

if (IncludeScenario("xls-read-cellvalues")) {
    AddScenario(measurements, "xls-read-cellvalues", "OfficeIMO.Excel Legacy XLS", "Read an HSSF-generated .xls workbook through the OfficeIMO legacy importer.", () => ReadOfficeImoXls(npoiXls.Value, rowCount), warmupIterations, measuredIterations);
    AddScenario(measurements, "xls-read-cellvalues", "NPOI HSSF", "Read the same HSSF-generated .xls workbook through HSSFWorkbook.", () => ReadNpoiWorkbook(npoiXls.Value, rowCount), warmupIterations, measuredIterations);
}

if (IncludeScenario("xls-read-formulas")) {
    AddScenario(measurements, "xls-read-formulas", "OfficeIMO.Excel Legacy XLS", "Read BIFF8 formula text and cached values from an HSSF-generated .xls workbook.", () => ReadOfficeImoXlsFormulas(npoiFormulaXls.Value, rowCount), warmupIterations, measuredIterations);
    AddScenario(measurements, "xls-read-formulas", "NPOI HSSF", "Read formula text and cached values from the same HSSF-generated .xls workbook.", () => ReadNpoiWorkbookFormulas(npoiFormulaXls.Value, rowCount), warmupIterations, measuredIterations);
}

if (measurements.Count == 0) {
    throw new ArgumentException("No NPOI comparison scenarios matched the requested filter.");
}

var result = new NpoiComparisonResult(
    DateTime.UtcNow,
    Environment.MachineName,
    System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
    rowCount,
    warmupIterations,
    measuredIterations,
    measurements);

string? outputDirectory = Path.GetDirectoryName(outputPath);
if (!string.IsNullOrWhiteSpace(outputDirectory)) {
    Directory.CreateDirectory(outputDirectory);
}

File.WriteAllText(outputPath, JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true }));
Console.WriteLine($"NPOI comparison written to '{outputPath}'.");

static void AddScenario(
    List<NpoiComparisonMeasurement> measurements,
    string scenario,
    string library,
    string description,
    Func<int> action,
    int warmupIterations,
    int measuredIterations) {
    for (int i = 0; i < warmupIterations; i++) {
        _ = action();
    }

    var elapsed = new double[measuredIterations];
    int metric = 0;
    for (int i = 0; i < measuredIterations; i++) {
        long start = Stopwatch.GetTimestamp();
        metric = action();
        elapsed[i] = Stopwatch.GetElapsedTime(start).TotalMilliseconds;
    }

    measurements.Add(new NpoiComparisonMeasurement(
        scenario,
        library,
        description,
        Math.Round(elapsed.Average(), 3),
        Math.Round(elapsed.Min(), 3),
        Math.Round(elapsed.Max(), 3),
        metric));
}

static byte[] WriteOfficeImoXlsx(IReadOnlyList<SalesRecord> records) {
    using var stream = new MemoryStream();
    using (var document = ExcelDocument.Create(stream, autoSave: false)) {
        ExcelSheet sheet = document.AddWorkSheet("Data");
        WriteOfficeImoRows(sheet, records);
        document.Save(stream);
    }

    return stream.ToArray();
}

static byte[] WriteNpoiXlsx(IReadOnlyList<SalesRecord> records) {
    using var stream = new MemoryStream();
    using var workbook = new XSSFWorkbook();
    ISheet sheet = workbook.CreateSheet("Data");
    WriteNpoiRows(sheet, records);
    workbook.Write(stream, leaveOpen: true);
    return stream.ToArray();
}

static byte[] WriteNpoiXls(IReadOnlyList<SalesRecord> records) {
    using var stream = new MemoryStream();
    using var workbook = new HSSFWorkbook();
    ISheet sheet = workbook.CreateSheet("Data");
    WriteNpoiRows(sheet, records);
    workbook.Write(stream, leaveOpen: true);
    return stream.ToArray();
}

static byte[] WriteNpoiFormulaXls(IReadOnlyList<SalesRecord> records) {
    using var stream = new MemoryStream();
    using var workbook = new HSSFWorkbook();
    ISheet sheet = workbook.CreateSheet("Data");
    IRow header = sheet.CreateRow(0);
    header.CreateCell(0).SetCellValue("Id");
    header.CreateCell(1).SetCellValue("Amount");
    header.CreateCell(2).SetCellValue("Rate");
    header.CreateCell(3).SetCellValue("Total");

    for (int i = 0; i < records.Count; i++) {
        int oneBasedRow = i + 2;
        IRow row = sheet.CreateRow(i + 1);
        SalesRecord record = records[i];
        row.CreateCell(0).SetCellValue(record.Id);
        row.CreateCell(1).SetCellValue(record.Amount);
        row.CreateCell(2).SetCellValue(record.Active ? 1.2d : 0.8d);
        row.CreateCell(3).SetCellFormula($"B{oneBasedRow}*C{oneBasedRow}");
    }

    HSSFFormulaEvaluator.EvaluateAllFormulaCells(workbook);
    workbook.Write(stream, leaveOpen: true);
    return stream.ToArray();
}

static void WriteOfficeImoRows(ExcelSheet sheet, IReadOnlyList<SalesRecord> records) {
    sheet.CellValue(1, 1, "Id");
    sheet.CellValue(1, 2, "Region");
    sheet.CellValue(1, 3, "Owner");
    sheet.CellValue(1, 4, "Amount");
    sheet.CellValue(1, 5, "Active");

    for (int i = 0; i < records.Count; i++) {
        int row = i + 2;
        SalesRecord record = records[i];
        sheet.CellValue(row, 1, record.Id);
        sheet.CellValue(row, 2, record.Region);
        sheet.CellValue(row, 3, record.Owner);
        sheet.CellValue(row, 4, record.Amount);
        sheet.CellValue(row, 5, record.Active);
    }
}

static void WriteNpoiRows(ISheet sheet, IReadOnlyList<SalesRecord> records) {
    IRow header = sheet.CreateRow(0);
    header.CreateCell(0).SetCellValue("Id");
    header.CreateCell(1).SetCellValue("Region");
    header.CreateCell(2).SetCellValue("Owner");
    header.CreateCell(3).SetCellValue("Amount");
    header.CreateCell(4).SetCellValue("Active");

    for (int i = 0; i < records.Count; i++) {
        IRow row = sheet.CreateRow(i + 1);
        SalesRecord record = records[i];
        row.CreateCell(0).SetCellValue(record.Id);
        row.CreateCell(1).SetCellValue(record.Region);
        row.CreateCell(2).SetCellValue(record.Owner);
        row.CreateCell(3).SetCellValue(record.Amount);
        row.CreateCell(4).SetCellValue(record.Active);
    }
}

static int ReadOfficeImoXlsx(byte[] workbookBytes, int rowCount) {
    using var reader = ExcelDocumentReader.Open(workbookBytes);
    object?[,] values = reader.GetSheet("Data").ReadRange($"A1:E{rowCount + 1}", ExecutionMode.Sequential);
    int metric = 0;
    for (int row = 0; row < values.GetLength(0); row++) {
        for (int column = 0; column < values.GetLength(1); column++) {
            metric = AddValueMetric(metric, values[row, column]);
        }
    }

    return metric;
}

static int ReadOfficeImoXls(byte[] workbookBytes, int rowCount) {
    LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(workbookBytes, new LegacyXlsImportOptions { ReportUnsupportedRecords = true });
    LegacyXlsWorksheet worksheet = workbook.Worksheets.Single(sheet => sheet.Name == "Data");
    int expectedCellCount = checked((rowCount + 1) * 5);
    if (worksheet.Cells.Count != expectedCellCount) {
        throw new InvalidOperationException($"Expected {expectedCellCount} cells, got {worksheet.Cells.Count}.");
    }

    int metric = 0;
    foreach (LegacyXlsCell cell in worksheet.Cells) {
        metric = AddValueMetric(metric, cell.Value);
    }

    return metric;
}

static int ReadOfficeImoXlsFormulas(byte[] workbookBytes, int rowCount) {
    LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(workbookBytes, new LegacyXlsImportOptions { ReportUnsupportedRecords = true });
    LegacyXlsWorksheet worksheet = workbook.Worksheets.Single(sheet => sheet.Name == "Data");
    List<LegacyXlsCell> formulaCells = worksheet.Cells
        .Where(cell => cell.IsFormula)
        .OrderBy(cell => cell.Row)
        .ThenBy(cell => cell.Column)
        .ToList();
    if (formulaCells.Count != rowCount) {
        throw new InvalidOperationException($"Expected {rowCount} formula cells, got {formulaCells.Count}.");
    }

    int metric = 0;
    foreach (LegacyXlsCell cell in formulaCells) {
        metric = AddValueMetric(metric, cell.FormulaText);
        metric = AddValueMetric(metric, cell.Value);
    }

    return metric;
}

static int ReadNpoiWorkbook(byte[] workbookBytes, int rowCount) {
    using var stream = new MemoryStream(workbookBytes, writable: false);
    using IWorkbook workbook = WorkbookFactory.Create(stream);
    ISheet sheet = workbook.GetSheet("Data");
    int metric = 0;
    for (int rowIndex = 0; rowIndex <= rowCount; rowIndex++) {
        IRow row = sheet.GetRow(rowIndex) ?? throw new InvalidOperationException($"Missing row {rowIndex + 1}.");
        for (int columnIndex = 0; columnIndex < 5; columnIndex++) {
            ICell cell = row.GetCell(columnIndex) ?? throw new InvalidOperationException($"Missing cell {rowIndex + 1},{columnIndex + 1}.");
            metric = AddValueMetric(metric, ReadNpoiCellValue(cell));
        }
    }

    return metric;
}

static int ReadNpoiWorkbookFormulas(byte[] workbookBytes, int rowCount) {
    using var stream = new MemoryStream(workbookBytes, writable: false);
    using IWorkbook workbook = WorkbookFactory.Create(stream);
    ISheet sheet = workbook.GetSheet("Data");
    int metric = 0;
    for (int rowIndex = 1; rowIndex <= rowCount; rowIndex++) {
        IRow row = sheet.GetRow(rowIndex) ?? throw new InvalidOperationException($"Missing row {rowIndex + 1}.");
        ICell formulaCell = row.GetCell(3) ?? throw new InvalidOperationException($"Missing formula cell {rowIndex + 1},4.");
        metric = AddValueMetric(metric, formulaCell.CellFormula);
        metric = AddValueMetric(metric, ReadNpoiFormulaCachedValue(formulaCell));
    }

    return metric;
}

static object? ReadNpoiCellValue(ICell cell) {
    return cell.CellType switch {
        CellType.String => cell.StringCellValue,
        CellType.Numeric => cell.NumericCellValue,
        CellType.Boolean => cell.BooleanCellValue,
        CellType.Blank => null,
        CellType.Error => cell.ErrorCellValue,
        CellType.Formula => cell.CellFormula,
        _ => cell.ToString()
    };
}

static object? ReadNpoiFormulaCachedValue(ICell cell) {
    if (cell.CellType != CellType.Formula) {
        return ReadNpoiCellValue(cell);
    }

    return cell.CachedFormulaResultType switch {
        CellType.String => cell.StringCellValue,
        CellType.Numeric => cell.NumericCellValue,
        CellType.Boolean => cell.BooleanCellValue,
        CellType.Blank => null,
        CellType.Error => cell.ErrorCellValue,
        _ => cell.ToString()
    };
}

static int AddValueMetric(int metric, object? value) {
    if (value == null) {
        return unchecked((metric * 397) ^ 17);
    }

    return AddStringMetric(metric, ToMetricText(value));
}

static string ToMetricText(object value) {
    return value switch {
        bool flag => flag ? "TRUE" : "FALSE",
        byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal
            => Convert.ToDecimal(value, CultureInfo.InvariantCulture).ToString("G29", CultureInfo.InvariantCulture),
        _ => Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty
    };
}

static int AddStringMetric(int metric, string value) {
    unchecked {
        int hash = metric;
        for (int i = 0; i < value.Length; i++) {
            hash = (hash * 397) ^ value[i];
        }

        return hash;
    }
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

static bool HasSwitch(string[] args, string optionName)
    => args.Any(arg => string.Equals(arg, optionName, StringComparison.OrdinalIgnoreCase));

static void WriteUsage() {
    Console.WriteLine("OfficeIMO.Excel NPOI opt-in comparison");
    Console.WriteLine();
    Console.WriteLine("Commands:");
    Console.WriteLine("  --rows N");
    Console.WriteLine("  --warmup N");
    Console.WriteLine("  --iterations N");
    Console.WriteLine("  --scenario name");
    Console.WriteLine("  --out path");
}

internal sealed record SalesRecord(int Id, string Region, string Owner, double Amount, bool Active) {
    private static readonly string[] Regions = ["North", "South", "East", "West", "Central"];
    private static readonly string[] Owners = ["Ava", "Noah", "Mia", "Liam", "Zoe", "Ethan", "Ivy", "Mason"];

    internal static IReadOnlyList<SalesRecord> Create(int count) {
        var records = new List<SalesRecord>(count);
        for (int i = 0; i < count; i++) {
            records.Add(new SalesRecord(
                i + 1,
                Regions[i % Regions.Length],
                Owners[i % Owners.Length],
                Math.Round(150 + ((i * 17.35) % 4500), 2),
                i % 3 != 0));
        }

        return records;
    }
}

internal sealed record NpoiComparisonResult(
    DateTime GeneratedAtUtc,
    string MachineName,
    string Framework,
    int RowCount,
    int WarmupIterations,
    int MeasuredIterations,
    IReadOnlyList<NpoiComparisonMeasurement> Measurements);

internal sealed record NpoiComparisonMeasurement(
    string Scenario,
    string Library,
    string Description,
    double AverageMilliseconds,
    double MinimumMilliseconds,
    double MaximumMilliseconds,
    int Metric);
