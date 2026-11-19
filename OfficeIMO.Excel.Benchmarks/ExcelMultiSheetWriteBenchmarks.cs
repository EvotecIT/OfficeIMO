#if NET10_0_OR_GREATER
using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Writer;
using LargeXlsx;
using SpreadCheetah;
using MiniExcelApi = MiniExcelLibs.MiniExcel;
using GeneratedRecord = OfficeIMO.Excel.Benchmarks.ExcelReaderGeneratedBenchmarks.GeneratedRecord;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Complete in-memory XLSX workbooks of many equally sized sheets (name, id, date, amount), each library on
/// its default options. Setup reopens every output with ExcelDataReader and checks the sheet and row counts, the
/// headers and every row's name, id, date and amount.
/// </summary>
[MemoryDiagnoser]
public class ExcelMultiSheetWriteBenchmarks {
    private static readonly string[] Headers = ["Name", "Id", "CreatedOn", "Amount"];
    private static readonly ExcelWriterOptions NoTable = new() { CreateTable = false };
    private static readonly XlsxStyle LargeXlsxDateStyle = XlsxStyle.Default.With(new XlsxNumberFormat("yyyy-mm-dd"));

    [Params(10, 50)]
    public int SheetCount { get; set; }

    [Params(5_000)]
    public int RowsPerSheet { get; set; }

    private GeneratedRecord[] _rows = [];

    [GlobalSetup]
    public void Setup() {
        _rows = new GeneratedRecord[RowsPerSheet];
        for (int index = 0; index < _rows.Length; index++) _rows[index] = ExcelReaderGeneratedBenchmarks.CreateRecord(index);
        string detail = $"SheetCount={SheetCount}; RowsPerSheet={RowsPerSheet}; ExcelDataReader reread every sheet's headers and every row's name, id, date and amount";
        ExcelOutputValidation.Check("OfficeIMO", "Typed", detail, () => Validate(nameof(OfficeIMO_Typed), OfficeIMO_Typed()));
        ExcelOutputValidation.Check("OfficeIMO", "TypedNoTable", detail, () => Validate(nameof(OfficeIMO_TypedNoTable), OfficeIMO_TypedNoTable()));
        ExcelOutputValidation.Check("OfficeIMO", "Editable", detail, () => Validate(nameof(OfficeIMO_Editable), OfficeIMO_Editable()));
        ExcelOutputValidation.Check("SpreadCheetah", "Default", detail, () => Validate(nameof(SpreadCheetah), SpreadCheetah()));
        ExcelOutputValidation.Check("ExcelReader.NET", "Default", detail, () => Validate(nameof(ExcelReaderNet), ExcelReaderNet()));
        ExcelOutputValidation.Check("MiniExcel", "Default", detail, () => Validate(nameof(MiniExcel), MiniExcel()));
        ExcelOutputValidation.Check("LargeXlsx", "Default", detail, () => Validate(nameof(LargeXlsx), LargeXlsx()));
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("case:Typed")]
    public byte[] OfficeIMO_Typed() {
        using var output = new MemoryStream();
        using (var writer = ExcelDocument.CreateWriter(ExcelFileFormat.Xlsx, output, leaveOpen: true)) {
            for (int sheet = 0; sheet < SheetCount; sheet++) {
                writer.WriteTable(SheetName(sheet), _rows, Headers,
                    static (row, value) => row.Write(value.Name).Write(value.Id).Write(value.CreatedOn).Write(value.Amount));
            }
        }
        return output.ToArray();
    }

    /// <summary>The row writer without its default native table, the same parts the other writers produce.</summary>
    [Benchmark]
    [BenchmarkCategory("case:TypedNoTable")]
    public byte[] OfficeIMO_TypedNoTable() {
        using var output = new MemoryStream();
        using (var writer = ExcelDocument.CreateWriter(ExcelFileFormat.Xlsx, output, NoTable, leaveOpen: true)) {
            for (int sheet = 0; sheet < SheetCount; sheet++) {
                writer.WriteTable(SheetName(sheet), _rows, Headers,
                    static (row, value) => row.Write(value.Name).Write(value.Id).Write(value.CreatedOn).Write(value.Amount));
            }
        }
        return output.ToArray();
    }

    [Benchmark]
    [BenchmarkCategory("case:Editable")]
    public byte[] OfficeIMO_Editable() {
        using var output = new MemoryStream();
        using (var document = ExcelDocument.Create()) {
            for (int sheet = 0; sheet < SheetCount; sheet++) {
                document.AddSheet(SheetName(sheet)).WriteData(_rows, ("Name", static row => row.Name), ("Id", static row => row.Id),
                    ("CreatedOn", static row => row.CreatedOn), ("Amount", static row => row.Amount));
            }
            document.Save(output);
        }
        return output.ToArray();
    }

    [Benchmark]
    [BenchmarkCategory("case:Default")]
    public byte[] SpreadCheetah() => SpreadCheetahAsync().GetAwaiter().GetResult();

    private async Task<byte[]> SpreadCheetahAsync() {
        using var output = new MemoryStream();
        await using (var spreadsheet = await global::SpreadCheetah.Spreadsheet.CreateNewAsync(output).ConfigureAwait(false)) {
            var cells = new DataCell[4];
            for (int sheet = 0; sheet < SheetCount; sheet++) {
                await spreadsheet.StartWorksheetAsync(SheetName(sheet)).ConfigureAwait(false);
                for (int column = 0; column < cells.Length; column++) cells[column] = new DataCell(Headers[column]);
                await spreadsheet.AddRowAsync(cells).ConfigureAwait(false);
                foreach (var row in _rows) {
                    cells[0] = new DataCell(row.Name);
                    cells[1] = new DataCell(row.Id);
                    cells[2] = new DataCell(row.CreatedOn);
                    cells[3] = new DataCell(row.Amount);
                    await spreadsheet.AddRowAsync(cells).ConfigureAwait(false);
                }
            }
            await spreadsheet.FinishAsync().ConfigureAwait(false);
        }
        return output.ToArray();
    }

    [Benchmark]
    [BenchmarkCategory("case:Default")]
    public byte[] ExcelReaderNet() {
        using var output = new MemoryStream();
        using (var workbook = XlsxWorkbookWriter.Create(output, leaveOpen: true)) {
            for (int sheet = 0; sheet < SheetCount; sheet++) {
                using var worksheet = workbook.AddSheet(SheetName(sheet));
                using (var header = worksheet.StartRow()) {
                    foreach (string name in Headers) header.Write(name);
                }
                foreach (var row in _rows) {
                    using var cells = worksheet.StartRow();
                    cells.Write(row.Name); cells.Write(row.Id); cells.Write(row.CreatedOn); cells.Write(row.Amount);
                }
                worksheet.End();
            }
            workbook.End();
        }
        return output.ToArray();
    }

    [Benchmark]
    [BenchmarkCategory("case:Default")]
    public byte[] MiniExcel() {
        using var output = new MemoryStream();
        var sheets = new Dictionary<string, object>(SheetCount);
        for (int sheet = 0; sheet < SheetCount; sheet++) sheets.Add(SheetName(sheet), _rows);
        MiniExcelApi.SaveAs(output, sheets, excelType: MiniExcelLibs.ExcelType.XLSX);
        return output.ToArray();
    }

    [Benchmark]
    [BenchmarkCategory("case:Default")]
    public byte[] LargeXlsx() {
        using var output = new MemoryStream();
        using (var writer = new XlsxWriter(output)) {
            for (int sheet = 0; sheet < SheetCount; sheet++) {
                writer.BeginWorksheet(SheetName(sheet));
                writer.BeginRow();
                foreach (string name in Headers) writer.Write(name);
                foreach (var row in _rows) {
                    writer.BeginRow();
                    writer.Write(row.Name).Write(row.Id).Write(row.CreatedOn, LargeXlsxDateStyle).Write(row.Amount);
                }
            }
        }
        return output.ToArray();
    }

    private static string SheetName(int index) => "Sheet" + (index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);

    private void Validate(string lane, byte[] workbook) {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        using var stream = new MemoryStream(workbook);
        using var reader = global::ExcelDataReader.ExcelReaderFactory.CreateReader(stream);
        int sheets = 0;
        do {
            if (!reader.Read() || reader.FieldCount != Headers.Length) throw new InvalidDataException($"{lane}: sheet {sheets + 1} has no header row.");
            for (int column = 0; column < Headers.Length; column++) {
                if (!Equals(reader.GetValue(column), Headers[column])) throw new InvalidDataException($"{lane}: sheet {sheets + 1} header {column + 1} is '{reader.GetValue(column)}'.");
            }
            int rows = 0;
            while (reader.Read()) {
                if (rows >= _rows.Length) throw new InvalidDataException($"{lane}: sheet {sheets + 1} has more than {_rows.Length} data rows.");
                CheckRow(lane, sheets, rows, _rows[rows], reader);
                rows++;
            }
            if (rows != RowsPerSheet) throw new InvalidDataException($"{lane}: sheet {sheets + 1} has {rows} data rows, expected {RowsPerSheet}.");
            sheets++;
        } while (reader.NextResult());
        if (sheets != SheetCount) throw new InvalidDataException($"{lane}: {sheets} sheets, expected {SheetCount}.");
    }

    // Writers store the date as a styled serial number; ExcelDataReader returns it as a DateTime or, unstyled, as the serial.
    private static void CheckRow(string lane, int sheet, int row, GeneratedRecord expected, System.Data.IDataRecord reader) {
        object name = reader.GetValue(0), id = reader.GetValue(1), date = reader.GetValue(2), amount = reader.GetValue(3);
        bool dateMatches = date switch {
            DateTime value => value == expected.CreatedOn,
            double serial => serial == expected.CreatedOn.ToOADate(),
            _ => false
        };
        if (!Equals(name, expected.Name) || id is not double idValue || idValue != expected.Id || !dateMatches
            || amount is not double amountValue || amountValue != expected.Amount) {
            throw new InvalidDataException($"{lane}: sheet {sheet + 1} data row {row + 1} holds {name}/{id}/{date}/{amount}.");
        }
    }
}
#endif
