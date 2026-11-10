using BenchmarkDotNet.Attributes;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Opens a workbook written by another producer, adds one cell below its data and saves it. XLBench's edit
/// shapes only open OfficeIMO-generated input, so this lane covers ClosedXML output and an Excel-written file.
/// Every output is checked through the Open XML SDK: the new cell holds its text and no row was lost.
/// Workbooks from other producers (EPPlus, LibreOffice, Google Sheets) plug in without a code change: every .xlsx in
/// the folder named by <c>OFFICEIMO_EXCEL_PRODUCER_FIXTURES</c> becomes one more case, named after the file.
/// </summary>
[MemoryDiagnoser]
[BenchmarkCategory("Performance", "EditableRoundTrip")]
public class ExcelProducerRoundTripBenchmarks {
    private const string EditedText = "Edited";
    private const string FixturesVariable = "OFFICEIMO_EXCEL_PRODUCER_FIXTURES";
    private byte[] _input = [];
    private int _lastRow;
    private int _rowCount;
    private string? _closedXmlProblem;

    [ParamsSource(nameof(Producers))]
    public string Producer { get; set; } = "OfficeIMO";

    public static IEnumerable<string> Producers() {
        yield return "OfficeIMO";
        yield return "ClosedXML";
        yield return "Excel";
        foreach (string file in FixtureFiles()) yield return Path.GetFileName(file);
    }

    private static IEnumerable<string> FixtureFiles() {
        string? folder = Environment.GetEnvironmentVariable(FixturesVariable);
        return string.IsNullOrEmpty(folder) || !Directory.Exists(folder)
            ? []
            : Directory.EnumerateFiles(folder, "*.xlsx").OrderBy(static file => file, StringComparer.Ordinal);
    }

    [GlobalSetup]
    public void Setup() {
        var records = ExcelBenchmarkScenarioFactory.CreateSalesRecords(10_000);
        _input = Producer switch {
            "OfficeIMO" => ExcelBenchmarkScenarioFactory.CreateWorkbookBytes(records),
            "ClosedXML" => CreateClosedXmlWorkbook(records),
            "Excel" => ReadExcelFixture(),
            _ => File.ReadAllBytes(FixtureFiles().SingleOrDefault(file => Path.GetFileName(file) == Producer)
                ?? throw new ArgumentOutOfRangeException(nameof(Producer), Producer, $"No such workbook in {FixturesVariable}."))
        };
        _lastRow = LastRowIndex(_input);
        _rowCount = RowCount(_input);
        Validate(OfficeIMO_Open_Edit_Save());
        // A competitor that cannot open the file, or whose output fails the same check, reports that error
        // instead of an unvalidated time.
        try { Validate(ClosedXML_Open_Edit_Save()); } catch (Exception exception) { _closedXmlProblem = exception.GetType().Name + ": " + exception.Message; }
    }

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMO_Open_Edit_Save() {
        using var output = new MemoryStream();
        using (var document = ExcelDocument.Load(new MemoryStream(_input, writable: false))) {
            document.Sheets[0].Cell(_lastRow + 1, 1).SetValue(EditedText);
            document.Save(output);
        }
        return output.ToArray();
    }

    [Benchmark]
    public byte[] ClosedXML_Open_Edit_Save() {
        if (_closedXmlProblem != null) throw new InvalidDataException("ClosedXML failed this workbook in setup: " + _closedXmlProblem);
        using var output = new MemoryStream();
        using (var workbook = new XLWorkbook(new MemoryStream(_input, writable: false))) {
            workbook.Worksheet(1).Cell(_lastRow + 1, 1).Value = EditedText;
            workbook.SaveAs(output);
        }
        return output.ToArray();
    }

    private static byte[] CreateClosedXmlWorkbook(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> records) {
        using var output = new MemoryStream();
        using (var workbook = new XLWorkbook()) {
            ExcelBenchmarkScenarioFactory.PopulateClosedXmlWorksheet(workbook.Worksheets.Add("Data"), records);
            workbook.SaveAs(output);
        }
        return output.ToArray();
    }

    private static byte[] ReadExcelFixture() {
        MarkPflug65KFixture.EnsureAuthentic(MarkPflug65KFixture.XlsxFileName);
        return File.ReadAllBytes(MarkPflug65KFixture.XlsxPath);
    }

    private static int RowCount(byte[] package) {
        using var document = SpreadsheetDocument.Open(new MemoryStream(package, writable: false), false);
        return FirstWorksheet(document).Descendants<Row>().Count();
    }

    private static int LastRowIndex(byte[] package) {
        using var document = SpreadsheetDocument.Open(new MemoryStream(package, writable: false), false);
        int last = 0;
        foreach (Row row in FirstWorksheet(document).Descendants<Row>()) last = row.RowIndex?.Value is uint index ? (int)index : last + 1;
        return last;
    }

    private void Validate(byte[] package) {
        using var document = SpreadsheetDocument.Open(new MemoryStream(package, writable: false), false);
        if (LastRowIndex(package) != _lastRow + 1) throw new InvalidDataException($"Expected the edit in row {_lastRow + 1}.");
        if (RowCount(package) != _rowCount + 1) throw new InvalidDataException("The save lost or added rows.");
        Cell cell = FirstWorksheet(document).Descendants<Row>().Last().Elements<Cell>().Single();
        string text = cell.DataType?.Value == CellValues.SharedString
            ? document.WorkbookPart!.SharedStringTablePart!.SharedStringTable!.Elements<SharedStringItem>()
                .ElementAt(int.Parse(cell.CellValue!.Text)).InnerText
            : cell.InlineString?.InnerText ?? cell.CellValue?.Text ?? string.Empty;
        if (text != EditedText) throw new InvalidDataException($"The edited cell holds '{text}'.");
    }

    private static Worksheet FirstWorksheet(SpreadsheetDocument document) {
        WorkbookPart workbook = document.WorkbookPart!;
        Sheet sheet = workbook.Workbook!.Sheets!.Elements<Sheet>().First();
        return ((WorksheetPart)workbook.GetPartById(sheet.Id!.Value!)).Worksheet!;
    }
}
