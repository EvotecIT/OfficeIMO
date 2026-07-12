using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
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
var npoiMetadataXls = new Lazy<byte[]>(() => WriteNpoiMetadataXls(records));
var npoiConditionalFormattingXls = new Lazy<byte[]>(() => WriteNpoiConditionalFormattingXls(records));
var npoiAutoFilterXls = new Lazy<byte[]>(() => WriteNpoiAutoFilterXls(records));
var npoiStylesXls = new Lazy<byte[]>(() => WriteNpoiStylesXls(records));
var npoiPicturesXls = new Lazy<byte[]>(() => NpoiPictureWorkbookFactory.WriteHssfPictureWorkbook(NpoiPictureWorkbookFactory.GetPictureCount(rowCount)));

if ((IncludeScenario("xls-read-cellvalues")
        || IncludeScenario("xls-read-formulas")
        || IncludeScenario("xls-read-metadata")
        || IncludeScenario("xls-read-conditional-formatting")
        || IncludeScenario("xls-read-autofilter-range")
        || IncludeScenario("xls-read-styles")) && rowCount + 1 > 65_536) {
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

if (IncludeScenario("xls-read-metadata")) {
    AddScenario(measurements, "xls-read-metadata", "OfficeIMO.Excel Legacy XLS", "Read BIFF8 comments, hyperlinks, merged ranges, and list validations from an HSSF-generated .xls workbook.", () => ReadOfficeImoXlsMetadata(npoiMetadataXls.Value, rowCount), warmupIterations, measuredIterations);
    AddScenario(measurements, "xls-read-metadata", "NPOI HSSF", "Read comments, hyperlinks, merged ranges, and list validations from the same HSSF-generated .xls workbook.", () => ReadNpoiWorkbookMetadata(npoiMetadataXls.Value, rowCount), warmupIterations, measuredIterations);
}

if (IncludeScenario("xls-read-conditional-formatting")) {
    AddScenario(measurements, "xls-read-conditional-formatting", "OfficeIMO.Excel Legacy XLS", "Read BIFF8 cell-is and formula conditional-formatting rules from an HSSF-generated .xls workbook.", () => ReadOfficeImoXlsConditionalFormatting(npoiConditionalFormattingXls.Value, rowCount), warmupIterations, measuredIterations);
    AddScenario(measurements, "xls-read-conditional-formatting", "NPOI HSSF", "Read cell-is and formula conditional-formatting rules from the same HSSF-generated .xls workbook.", () => ReadNpoiWorkbookConditionalFormatting(npoiConditionalFormattingXls.Value), warmupIterations, measuredIterations);
}

if (IncludeScenario("xls-read-autofilter-range")) {
    AddScenario(measurements, "xls-read-autofilter-range", "OfficeIMO.Excel Legacy XLS", "Read BIFF8 AutoFilter range/drop-down metadata from an HSSF-generated .xls workbook.", () => ReadOfficeImoXlsAutoFilterRange(npoiAutoFilterXls.Value, rowCount), warmupIterations, measuredIterations);
    AddScenario(measurements, "xls-read-autofilter-range", "NPOI HSSF", "Read the hidden AutoFilter range name from the same HSSF-generated .xls workbook.", () => ReadNpoiWorkbookAutoFilterRange(npoiAutoFilterXls.Value, rowCount), warmupIterations, measuredIterations);
}

if (IncludeScenario("xls-read-styles")) {
    AddScenario(measurements, "xls-read-styles", "OfficeIMO.Excel Legacy XLS", "Read BIFF8 font, fill, border, number-format, and alignment style signals from an HSSF-generated .xls workbook.", () => ReadOfficeImoXlsStyles(npoiStylesXls.Value, rowCount), warmupIterations, measuredIterations);
    AddScenario(measurements, "xls-read-styles", "NPOI HSSF", "Read font, fill, border, number-format, and alignment style signals from the same HSSF-generated .xls workbook.", () => ReadNpoiWorkbookStyles(npoiStylesXls.Value, rowCount), warmupIterations, measuredIterations);
}

if (IncludeScenario("xls-read-pictures")) {
    int pictureCount = NpoiPictureWorkbookFactory.GetPictureCount(rowCount);
    AddScenario(measurements, "xls-read-pictures", "OfficeIMO.Excel Legacy XLS", "Read BIFF8 embedded picture object and BLIP-store signals from an HSSF-generated .xls workbook.", () => NpoiPictureComparison.ReadOfficeImoXlsPictures(npoiPicturesXls.Value, pictureCount, AddValueMetric), warmupIterations, measuredIterations);
    AddScenario(measurements, "xls-read-pictures", "NPOI HSSF", "Read embedded picture bytes from the same HSSF-generated .xls workbook.", () => NpoiPictureComparison.ReadNpoiWorkbookPictures(npoiPicturesXls.Value, pictureCount, AddValueMetric), warmupIterations, measuredIterations);
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
    using (var document = ExcelDocument.Create(stream)) {
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

static byte[] WriteNpoiMetadataXls(IReadOnlyList<SalesRecord> records) {
    using var stream = new MemoryStream();
    using var workbook = new HSSFWorkbook();
    ISheet sheet = workbook.CreateSheet("Data");
    WriteNpoiRows(sheet, records);

    HSSFPatriarch drawing = (HSSFPatriarch)sheet.CreateDrawingPatriarch();
    int metadataRows = GetMetadataRowCount(records.Count);
    for (int i = 0; i < metadataRows; i++) {
        int rowIndex = i + 1;
        IRow row = sheet.GetRow(rowIndex) ?? throw new InvalidOperationException($"Missing row {rowIndex + 1}.");
        ICell ownerCell = row.GetCell(2) ?? throw new InvalidOperationException($"Missing owner cell {rowIndex + 1},3.");
        var anchor = new HSSFClientAnchor(0, 0, 0, 0, 3, rowIndex, 5, rowIndex + 3);
        IComment comment = drawing.CreateCellComment(anchor);
        comment.Author = "OfficeIMO";
        comment.String = new HSSFRichTextString($"Owner note {i + 1}");
        ownerCell.CellComment = comment;

        ICell regionCell = row.GetCell(1) ?? throw new InvalidOperationException($"Missing region cell {rowIndex + 1},2.");
        IHyperlink hyperlink = workbook.GetCreationHelper().CreateHyperlink(HyperlinkType.Url);
        hyperlink.Address = $"https://example.com/region/{records[i].Region.ToLowerInvariant()}";
        regionCell.Hyperlink = hyperlink;
    }

    int mergedRegionCount = GetMetadataMergedRegionCount(records.Count);
    for (int i = 0; i < mergedRegionCount; i++) {
        int rowIndex = (i * 2) + 1;
        sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 5, 6));
        IRow row = sheet.GetRow(rowIndex) ?? throw new InvalidOperationException($"Missing merged row {rowIndex + 1}.");
        row.CreateCell(5).SetCellValue($"Merged {i + 1}");
    }

    var validationRange = new CellRangeAddressList(1, records.Count, 4, 4);
    IDataValidationConstraint validationConstraint = DVConstraint.CreateExplicitListConstraint(["TRUE", "FALSE"]);
    var validation = new HSSFDataValidation(validationRange, validationConstraint) {
        SuppressDropDownArrow = false,
        EmptyCellAllowed = true,
        ShowErrorBox = true
    };
    sheet.AddValidationData(validation);

    workbook.Write(stream, leaveOpen: true);
    return stream.ToArray();
}

static byte[] WriteNpoiConditionalFormattingXls(IReadOnlyList<SalesRecord> records) {
    using var stream = new MemoryStream();
    using var workbook = new HSSFWorkbook();
    ISheet sheet = workbook.CreateSheet("Data");
    WriteNpoiRows(sheet, records);

    ISheetConditionalFormatting conditionalFormatting = sheet.SheetConditionalFormatting;
    IConditionalFormattingRule highAmountRule = conditionalFormatting.CreateConditionalFormattingRule(ComparisonOperator.GreaterThan, "1000");
    IPatternFormatting highAmountPattern = highAmountRule.CreatePatternFormatting();
    highAmountPattern.FillPattern = FillPattern.SolidForeground;
    highAmountPattern.FillForegroundColor = IndexedColors.Yellow.Index;
    conditionalFormatting.AddConditionalFormatting(
        [new CellRangeAddress(1, records.Count, 3, 3)],
        highAmountRule);

    IConditionalFormattingRule inactiveRule = conditionalFormatting.CreateConditionalFormattingRule("$E2=FALSE");
    IPatternFormatting inactivePattern = inactiveRule.CreatePatternFormatting();
    inactivePattern.FillPattern = FillPattern.SolidForeground;
    inactivePattern.FillForegroundColor = IndexedColors.Rose.Index;
    conditionalFormatting.AddConditionalFormatting(
        [new CellRangeAddress(1, records.Count, 4, 4)],
        inactiveRule);

    workbook.Write(stream, leaveOpen: true);
    return stream.ToArray();
}

static byte[] WriteNpoiAutoFilterXls(IReadOnlyList<SalesRecord> records) {
    using var stream = new MemoryStream();
    using var workbook = new HSSFWorkbook();
    ISheet sheet = workbook.CreateSheet("Data");
    WriteNpoiRows(sheet, records);
    sheet.SetAutoFilter(new CellRangeAddress(0, records.Count, 0, 4));
    workbook.Write(stream, leaveOpen: true);
    return stream.ToArray();
}

static byte[] WriteNpoiStylesXls(IReadOnlyList<SalesRecord> records) {
    using var stream = new MemoryStream();
    using var workbook = new HSSFWorkbook();
    ISheet sheet = workbook.CreateSheet("Data");
    IDataFormat dataFormat = workbook.CreateDataFormat();

    IFont headerFont = workbook.CreateFont();
    headerFont.IsBold = true;
    headerFont.Color = IndexedColors.White.Index;

    ICellStyle headerStyle = workbook.CreateCellStyle();
    headerStyle.SetFont(headerFont);
    headerStyle.FillForegroundColor = IndexedColors.DarkBlue.Index;
    headerStyle.FillPattern = FillPattern.SolidForeground;
    headerStyle.BorderBottom = BorderStyle.Medium;

    ICellStyle amountStyle = workbook.CreateCellStyle();
    amountStyle.DataFormat = dataFormat.GetFormat("$#,##0.00");
    amountStyle.Alignment = HorizontalAlignment.Right;

    ICellStyle ownerStyle = workbook.CreateCellStyle();
    ownerStyle.Alignment = HorizontalAlignment.Center;
    ownerStyle.WrapText = true;

    ICellStyle inactiveStyle = workbook.CreateCellStyle();
    inactiveStyle.FillForegroundColor = IndexedColors.Rose.Index;
    inactiveStyle.FillPattern = FillPattern.SolidForeground;

    IRow header = sheet.CreateRow(0);
    string[] headings = ["Id", "Region", "Owner", "Amount", "Active"];
    for (int columnIndex = 0; columnIndex < headings.Length; columnIndex++) {
        ICell cell = header.CreateCell(columnIndex);
        cell.SetCellValue(headings[columnIndex]);
        cell.CellStyle = headerStyle;
    }

    for (int i = 0; i < records.Count; i++) {
        IRow row = sheet.CreateRow(i + 1);
        SalesRecord record = records[i];
        row.CreateCell(0).SetCellValue(record.Id);
        row.CreateCell(1).SetCellValue(record.Region);

        ICell ownerCell = row.CreateCell(2);
        ownerCell.SetCellValue(record.Owner);
        ownerCell.CellStyle = ownerStyle;

        ICell amountCell = row.CreateCell(3);
        amountCell.SetCellValue(record.Amount);
        amountCell.CellStyle = amountStyle;

        ICell activeCell = row.CreateCell(4);
        activeCell.SetCellValue(record.Active);
        if (!record.Active) {
            activeCell.CellStyle = inactiveStyle;
        }
    }

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
    LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(workbookBytes, new LegacyXlsImportOptions { ReportUnsupportedContent = true });
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
    LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(workbookBytes, new LegacyXlsImportOptions { ReportUnsupportedContent = true });
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

static int ReadOfficeImoXlsMetadata(byte[] workbookBytes, int rowCount) {
    LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(workbookBytes, new LegacyXlsImportOptions { ReportUnsupportedContent = true });
    LegacyXlsWorksheet worksheet = workbook.Worksheets.Single(sheet => sheet.Name == "Data");
    int expectedCommentCount = GetMetadataRowCount(rowCount);
    int expectedHyperlinkCount = GetMetadataRowCount(rowCount);
    int expectedMergedRangeCount = GetMetadataMergedRegionCount(rowCount);
    int expectedValidationCount = 1;
    ValidateMetadataCounts(
        worksheet.Comments.Count,
        worksheet.Hyperlinks.Count,
        worksheet.MergedRanges.Count,
        worksheet.DataValidations.Count,
        expectedCommentCount,
        expectedHyperlinkCount,
        expectedMergedRangeCount,
        expectedValidationCount,
        "OfficeIMO legacy XLS");

    int metric = 0;
    foreach (LegacyXlsComment comment in worksheet.Comments.OrderBy(comment => comment.Row).ThenBy(comment => comment.Column)) {
        metric = AddValueMetric(metric, comment.Text);
        metric = AddValueMetric(metric, comment.Author);
    }

    foreach (LegacyXlsHyperlink hyperlink in worksheet.Hyperlinks.OrderBy(hyperlink => hyperlink.StartRow).ThenBy(hyperlink => hyperlink.StartColumn)) {
        metric = AddValueMetric(metric, hyperlink.Target);
    }

    metric = AddValueMetric(metric, worksheet.MergedRanges.Count);
    metric = AddValueMetric(metric, worksheet.DataValidations.Count);
    return metric;
}

static int ReadOfficeImoXlsConditionalFormatting(byte[] workbookBytes, int rowCount) {
    LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(workbookBytes, new LegacyXlsImportOptions { ReportUnsupportedContent = true });
    LegacyXlsWorksheet worksheet = workbook.Worksheets.Single(sheet => sheet.Name == "Data");
    ValidateConditionalFormattingCounts(worksheet.ConditionalFormattings.Count, "OfficeIMO legacy XLS");

    int metric = 0;
    foreach (LegacyXlsConditionalFormatting formatting in worksheet.ConditionalFormattings.OrderBy(formatting => formatting.Ranges.FirstOrDefault(), StringComparer.Ordinal)) {
        metric = AddValueMetric(metric, formatting.Type.ToString());
        metric = AddValueMetric(metric, formatting.Operator?.ToString());
        metric = AddValueMetric(metric, formatting.Formula1);
        metric = AddValueMetric(metric, formatting.Formula2);
        metric = AddValueMetric(metric, string.Join(";", formatting.Ranges));
    }

    return metric;
}

static int ReadOfficeImoXlsAutoFilterRange(byte[] workbookBytes, int rowCount) {
    LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(workbookBytes, new LegacyXlsImportOptions { ReportUnsupportedContent = true });
    LegacyXlsWorksheet worksheet = workbook.Worksheets.Single(sheet => sheet.Name == "Data");
    LegacyXlsDefinedName filterName = workbook.DefinedNames.Single(name =>
        string.Equals(name.Name, "_FilterDatabase", StringComparison.OrdinalIgnoreCase)
        && name.LocalSheetIndex == 0);
    ValidateAutoFilterRange(filterName.Reference, rowCount, "OfficeIMO legacy XLS");

    int metric = 0;
    metric = AddValueMetric(metric, filterName.Name);
    metric = AddValueMetric(metric, NormalizeSheetQuote(filterName.Reference));
    return metric;
}

static int ReadOfficeImoXlsStyles(byte[] workbookBytes, int rowCount) {
    LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(workbookBytes, new LegacyXlsImportOptions { ReportUnsupportedContent = true });
    LegacyXlsWorksheet worksheet = workbook.Worksheets.Single(sheet => sheet.Name == "Data");
    ValidateStyleCellCount(worksheet.Cells.Count, rowCount, "OfficeIMO legacy XLS");

    LegacyXlsCellFormat headerFormat = GetOfficeImoCellFormat(workbook, worksheet, 1, 1);
    LegacyXlsFont headerFont = GetOfficeImoFont(workbook, headerFormat.FontIndex);
    if (!headerFont.Bold || headerFormat.FillPattern == 0 || headerFormat.Border?.BottomStyle != 2) {
        throw new InvalidOperationException("OfficeIMO legacy XLS did not preserve the expected header font/fill/border style signals.");
    }

    LegacyXlsCellFormat ownerFormat = GetOfficeImoCellFormat(workbook, worksheet, 2, 3);
    LegacyXlsCellFormat amountFormat = GetOfficeImoCellFormat(workbook, worksheet, 2, 4);
    LegacyXlsCellFormat inactiveFormat = GetOfficeImoCellFormat(workbook, worksheet, 2, 5);
    if (ownerFormat.HorizontalAlignment != 2 || !ownerFormat.WrapText) {
        throw new InvalidOperationException("OfficeIMO legacy XLS did not preserve the expected owner alignment style signals.");
    }

    if (amountFormat.NumberFormatCode == null || amountFormat.NumberFormatCode.IndexOf('$', StringComparison.Ordinal) < 0) {
        throw new InvalidOperationException("OfficeIMO legacy XLS did not preserve the expected currency number format signal.");
    }

    if (inactiveFormat.FillPattern == 0) {
        throw new InvalidOperationException("OfficeIMO legacy XLS did not preserve the expected inactive-row fill style signal.");
    }

    return BuildStyleMetric(rowCount);
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

static int ReadNpoiWorkbookMetadata(byte[] workbookBytes, int rowCount) {
    using var stream = new MemoryStream(workbookBytes, writable: false);
    using IWorkbook workbook = WorkbookFactory.Create(stream);
    ISheet sheet = workbook.GetSheet("Data");
    int commentCount = 0;
    int hyperlinkCount = 0;
    int metric = 0;
    for (int rowIndex = 1; rowIndex <= rowCount; rowIndex++) {
        IRow? row = sheet.GetRow(rowIndex);
        if (row == null) {
            continue;
        }

        foreach (ICell cell in row.Cells) {
            if (cell.CellComment != null) {
                commentCount++;
                metric = AddValueMetric(metric, cell.CellComment.String.String);
                metric = AddValueMetric(metric, cell.CellComment.Author);
            }

            if (cell.Hyperlink != null) {
                hyperlinkCount++;
                metric = AddValueMetric(metric, cell.Hyperlink.Address);
            }
        }
    }

    int mergedRangeCount = sheet.NumMergedRegions;
    int validationCount = sheet.GetDataValidations().Count;
    ValidateMetadataCounts(
        commentCount,
        hyperlinkCount,
        mergedRangeCount,
        validationCount,
        GetMetadataRowCount(rowCount),
        GetMetadataRowCount(rowCount),
        GetMetadataMergedRegionCount(rowCount),
        1,
        "NPOI HSSF");

    metric = AddValueMetric(metric, mergedRangeCount);
    metric = AddValueMetric(metric, validationCount);
    return metric;
}

static int ReadNpoiWorkbookConditionalFormatting(byte[] workbookBytes) {
    using var stream = new MemoryStream(workbookBytes, writable: false);
    using IWorkbook workbook = WorkbookFactory.Create(stream);
    ISheet sheet = workbook.GetSheet("Data");
    ISheetConditionalFormatting conditionalFormatting = sheet.SheetConditionalFormatting;
    ValidateConditionalFormattingCounts(conditionalFormatting.NumConditionalFormattings, "NPOI HSSF");

    int metric = 0;
    for (int formattingIndex = 0; formattingIndex < conditionalFormatting.NumConditionalFormattings; formattingIndex++) {
        IConditionalFormatting formatting = conditionalFormatting.GetConditionalFormattingAt(formattingIndex);
        for (int ruleIndex = 0; ruleIndex < formatting.NumberOfRules; ruleIndex++) {
            IConditionalFormattingRule rule = formatting.GetRule(ruleIndex);
            metric = AddValueMetric(metric, NormalizeConditionalFormattingType(rule.ConditionType.ToString()));
            metric = AddValueMetric(metric, NormalizeConditionalFormattingOperator(rule.ComparisonOperation.ToString()));
            metric = AddValueMetric(metric, rule.Formula1);
            metric = AddValueMetric(metric, rule.Formula2);
        }

        foreach (CellRangeAddress range in formatting.GetFormattingRanges()) {
            metric = AddValueMetric(metric, range.FormatAsString());
        }
    }

    return metric;
}

static int ReadNpoiWorkbookAutoFilterRange(byte[] workbookBytes, int rowCount) {
    using var stream = new MemoryStream(workbookBytes, writable: false);
    using IWorkbook workbook = WorkbookFactory.Create(stream);
    IName filterName = Enumerable.Range(0, workbook.NumberOfNames)
        .Select(workbook.GetNameAt)
        .Single(name => string.Equals(name.NameName, "_FilterDatabase", StringComparison.OrdinalIgnoreCase)
            && name.SheetIndex == 0);
    ValidateAutoFilterRange(filterName.RefersToFormula, rowCount, "NPOI HSSF");

    int metric = 0;
    metric = AddValueMetric(metric, filterName.NameName);
    metric = AddValueMetric(metric, NormalizeSheetQuote(filterName.RefersToFormula));
    return metric;
}

static int ReadNpoiWorkbookStyles(byte[] workbookBytes, int rowCount) {
    using var stream = new MemoryStream(workbookBytes, writable: false);
    using IWorkbook workbook = WorkbookFactory.Create(stream);
    ISheet sheet = workbook.GetSheet("Data");
    ValidateStyleCellCount(CountNpoiCells(sheet, rowCount), rowCount, "NPOI HSSF");

    ICellStyle headerStyle = GetNpoiCell(sheet, 0, 0).CellStyle;
    IFont headerFont = workbook.GetFontAt(headerStyle.FontIndex);
    if (!headerFont.IsBold || headerStyle.FillPattern != FillPattern.SolidForeground || headerStyle.BorderBottom != BorderStyle.Medium) {
        throw new InvalidOperationException("NPOI HSSF did not preserve the expected header font/fill/border style signals.");
    }

    ICellStyle ownerStyle = GetNpoiCell(sheet, 1, 2).CellStyle;
    ICellStyle amountStyle = GetNpoiCell(sheet, 1, 3).CellStyle;
    ICellStyle inactiveStyle = GetNpoiCell(sheet, 1, 4).CellStyle;
    if (ownerStyle.Alignment != HorizontalAlignment.Center || !ownerStyle.WrapText) {
        throw new InvalidOperationException("NPOI HSSF did not preserve the expected owner alignment style signals.");
    }

    if (amountStyle.GetDataFormatString().IndexOf('$', StringComparison.Ordinal) < 0) {
        throw new InvalidOperationException("NPOI HSSF did not preserve the expected currency number format signal.");
    }

    if (inactiveStyle.FillPattern != FillPattern.SolidForeground) {
        throw new InvalidOperationException("NPOI HSSF did not preserve the expected inactive-row fill style signal.");
    }

    return BuildStyleMetric(rowCount);
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
        string text => text.TrimEnd('\0'),
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

static int GetMetadataRowCount(int rowCount) => Math.Min(rowCount, 64);

static int GetMetadataMergedRegionCount(int rowCount) => Math.Min(Math.Max(rowCount / 2, 1), 16);

static void ValidateMetadataCounts(
    int commentCount,
    int hyperlinkCount,
    int mergedRangeCount,
    int validationCount,
    int expectedCommentCount,
    int expectedHyperlinkCount,
    int expectedMergedRangeCount,
    int expectedValidationCount,
    string libraryName) {
    if (commentCount != expectedCommentCount
        || hyperlinkCount != expectedHyperlinkCount
        || mergedRangeCount != expectedMergedRangeCount
        || validationCount != expectedValidationCount) {
        throw new InvalidOperationException(
            $"{libraryName} metadata counts did not match. "
            + $"Comments {commentCount}/{expectedCommentCount}, "
            + $"Hyperlinks {hyperlinkCount}/{expectedHyperlinkCount}, "
            + $"MergedRanges {mergedRangeCount}/{expectedMergedRangeCount}, "
            + $"DataValidations {validationCount}/{expectedValidationCount}.");
    }
}

static void ValidateConditionalFormattingCounts(int conditionalFormattingCount, string libraryName) {
    const int ExpectedConditionalFormattingCount = 2;
    if (conditionalFormattingCount != ExpectedConditionalFormattingCount) {
        throw new InvalidOperationException(
            $"{libraryName} conditional formatting count did not match. "
            + $"ConditionalFormatting {conditionalFormattingCount}/{ExpectedConditionalFormattingCount}.");
    }
}

static void ValidateAutoFilterRange(string? reference, int rowCount, string libraryName) {
    string expectedReference = $"'Data'!$A$1:$E${rowCount + 1}";
    if (!string.Equals(NormalizeSheetQuote(reference), NormalizeSheetQuote(expectedReference), StringComparison.OrdinalIgnoreCase)) {
        throw new InvalidOperationException(
            $"{libraryName} AutoFilter range did not match. "
            + $"Reference {reference ?? "<null>"}/{expectedReference}.");
    }
}

static void ValidateStyleCellCount(int actualCellCount, int rowCount, string libraryName) {
    int expectedCellCount = checked((rowCount + 1) * 5);
    if (actualCellCount != expectedCellCount) {
        throw new InvalidOperationException($"{libraryName} style workbook cell count did not match. Cells {actualCellCount}/{expectedCellCount}.");
    }
}

static LegacyXlsCellFormat GetOfficeImoCellFormat(LegacyXlsWorkbook workbook, LegacyXlsWorksheet worksheet, int row, int column) {
    LegacyXlsCell cell = worksheet.Cells.Single(cell => cell.Row == row && cell.Column == column);
    if (cell.StyleIndex >= workbook.CellFormats.Count) {
        throw new InvalidOperationException($"OfficeIMO legacy XLS style index {cell.StyleIndex} is outside the parsed XF table.");
    }

    return workbook.CellFormats[cell.StyleIndex];
}

static LegacyXlsFont GetOfficeImoFont(LegacyXlsWorkbook workbook, ushort fontIndex) {
    int index = fontIndex < 4 ? fontIndex : fontIndex > 4 ? fontIndex - 1 : -1;
    if (index < 0 || index >= workbook.Fonts.Count) {
        throw new InvalidOperationException($"OfficeIMO legacy XLS font index {fontIndex} is outside the parsed font table.");
    }

    return workbook.Fonts[index];
}

static int CountNpoiCells(ISheet sheet, int rowCount) {
    int cellCount = 0;
    for (int rowIndex = 0; rowIndex <= rowCount; rowIndex++) {
        IRow row = sheet.GetRow(rowIndex) ?? throw new InvalidOperationException($"Missing row {rowIndex + 1}.");
        cellCount += row.Cells.Count;
    }

    return cellCount;
}

static ICell GetNpoiCell(ISheet sheet, int rowIndex, int columnIndex) {
    IRow row = sheet.GetRow(rowIndex) ?? throw new InvalidOperationException($"Missing row {rowIndex + 1}.");
    return row.GetCell(columnIndex) ?? throw new InvalidOperationException($"Missing cell {rowIndex + 1},{columnIndex + 1}.");
}

static int BuildStyleMetric(int rowCount) {
    int metric = AddValueMetric(0, rowCount);
    metric = AddValueMetric(metric, "Header:BoldFillBottomBorder");
    metric = AddValueMetric(metric, "Owner:CenteredWrapped");
    metric = AddValueMetric(metric, "Amount:Currency");
    metric = AddValueMetric(metric, "Inactive:Fill");
    return metric;
}

static string NormalizeSheetQuote(string? reference) {
    return (reference ?? string.Empty).Replace("'Data'!", "Data!", StringComparison.OrdinalIgnoreCase);
}

static string? NormalizeConditionalFormattingType(string? value) {
    return value switch {
        null => null,
        "Formula" => "Expression",
        _ => value
    };
}

static string? NormalizeConditionalFormattingOperator(string? value) {
    return value switch {
        null => null,
        "NoComparison" => null,
        _ => value
    };
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
    Console.WriteLine();
    Console.WriteLine("Scenarios:");
    Console.WriteLine("  xlsx-write-cellvalues");
    Console.WriteLine("  xlsx-read-cellvalues");
    Console.WriteLine("  xls-read-cellvalues");
    Console.WriteLine("  xls-read-formulas");
    Console.WriteLine("  xls-read-metadata");
    Console.WriteLine("  xls-read-conditional-formatting");
    Console.WriteLine("  xls-read-autofilter-range");
    Console.WriteLine("  xls-read-styles");
    Console.WriteLine("  xls-read-pictures");
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
