#if NET10_0_OR_GREATER
using System.Globalization;
using System.Text;
using BenchmarkDotNet.Attributes;
using OfficeIMO.CSV;
using OfficeIMO.Excel.Csv;
using MiniExcelApi = MiniExcelLibs.MiniExcel;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// A CSV file (name, id, ISO date, amount) converted to a complete in-memory XLSX workbook. Setup reopens every
/// output with ExcelDataReader and checks the headers, the row count and every row's name, id and amount.
/// </summary>
[MemoryDiagnoser]
public class CsvToXlsxConversionBenchmarks {
    [Params(50_000)]
    public int Rows { get; set; }

    private byte[] _csv = [];

    [GlobalSetup]
    public void Setup() {
        _csv = ExcelCsvConversionData.CreateCsv(Rows);
        string detail = $"Rows={Rows}; ExcelDataReader reread the headers and every row's name, id and amount";
        ExcelOutputValidation.Check("OfficeIMO", "ImportCsv", detail,
            () => ExcelCsvConversionData.ValidateWorkbook(nameof(OfficeIMO_ImportCsv), OfficeIMO_ImportCsv(), Rows));
        ExcelOutputValidation.Check("OfficeIMO", "StreamingText", detail,
            () => ExcelCsvConversionData.ValidateWorkbook(nameof(OfficeIMO_StreamingText), OfficeIMO_StreamingText(), Rows));
        ExcelOutputValidation.Check("MiniExcel", "Default", detail,
            () => ExcelCsvConversionData.ValidateWorkbook(nameof(MiniExcel), MiniExcel(), Rows));
        ExcelOutputValidation.Check("Sylvan.Data.Excel", "Default", detail,
            () => ExcelCsvConversionData.ValidateWorkbook(nameof(Sylvan), Sylvan(), Rows));
    }

    /// <summary>The editable import on its defaults: inferred column types and a native table.</summary>
    [Benchmark(Baseline = true)]
    [BenchmarkCategory("case:ImportCsv")]
    public byte[] OfficeIMO_ImportCsv() {
        using var input = new MemoryStream(_csv, writable: false);
        using var output = new MemoryStream();
        using (var document = ExcelDocument.Create()) {
            document.AddSheet("Sheet1").ImportCsv(input);
            document.Save(output);
        }
        return output.ToArray();
    }

    /// <summary>The CSV reader streamed into the row writer as text, the same work as the other lanes.</summary>
    [Benchmark]
    [BenchmarkCategory("case:StreamingText")]
    public byte[] OfficeIMO_StreamingText() {
        using var input = new MemoryStream(_csv, writable: false);
        using var output = new MemoryStream();
        using (var reader = CsvDocument.OpenReader(input))
        using (var writer = ExcelDocument.CreateWriter(output, new ExcelWriterOptions { CreateTable = false }, leaveOpen: true)) {
            writer.WriteTable("Sheet1", reader);
        }
        return output.ToArray();
    }

    [Benchmark]
    [BenchmarkCategory("case:Default")]
    public byte[] MiniExcel() {
        using var input = new MemoryStream(_csv, writable: false);
        using var output = new MemoryStream();
        MiniExcelApi.ConvertCsvToXlsx(input, output);
        return output.ToArray();
    }

    [Benchmark]
    [BenchmarkCategory("case:Default")]
    public byte[] Sylvan() {
        using var text = new StreamReader(new MemoryStream(_csv, writable: false), Encoding.UTF8);
        using var output = new MemoryStream();
        using (var reader = global::Sylvan.Data.Csv.CsvDataReader.Create(text))
        using (var writer = global::Sylvan.Data.Excel.ExcelDataWriter.Create(output, global::Sylvan.Data.Excel.ExcelWorkbookType.ExcelXml)) {
            writer.Write(reader);
        }
        return output.ToArray();
    }
}

/// <summary>
/// A one-sheet XLSX workbook (name, id, date, amount) converted to in-memory CSV. Setup checks the header, the line
/// count and every line's name and id in every output.
/// </summary>
[MemoryDiagnoser]
public class XlsxToCsvConversionBenchmarks {
    [Params(50_000)]
    public int Rows { get; set; }

    private byte[] _xlsx = [];

    [GlobalSetup]
    public void Setup() {
        _xlsx = ExcelCsvConversionData.CreateWorkbook(Rows);
        string detail = $"Rows={Rows}; the header, the line count and every line's name and id checked";
        ExcelOutputValidation.Check("OfficeIMO", "SaveAsCsv", detail,
            () => ExcelCsvConversionData.ValidateCsv(nameof(OfficeIMO_SaveAsCsv), OfficeIMO_SaveAsCsv(), Rows));
        ExcelOutputValidation.Check("OfficeIMO", "Reader", detail,
            () => ExcelCsvConversionData.ValidateCsv(nameof(OfficeIMO_Reader), OfficeIMO_Reader(), Rows));
        ExcelOutputValidation.Check("MiniExcel", "Default", detail,
            () => ExcelCsvConversionData.ValidateCsv(nameof(MiniExcel), MiniExcel(), Rows));
        ExcelOutputValidation.Check("Sylvan.Data.Excel", "Default", detail,
            () => ExcelCsvConversionData.ValidateCsv(nameof(Sylvan), Sylvan(), Rows));
    }

    /// <summary>The editable route: load the workbook and save its sheet as CSV.</summary>
    [Benchmark(Baseline = true)]
    [BenchmarkCategory("case:SaveAsCsv")]
    public byte[] OfficeIMO_SaveAsCsv() {
        using var input = new MemoryStream(_xlsx, writable: false);
        using var output = new MemoryStream();
        using (var document = ExcelDocument.Load(input)) {
            document.Sheets[0].SaveAsCsv(output);
        }
        return output.ToArray();
    }

    /// <summary>The streaming route: the workbook reader written straight to the CSV writer.</summary>
    [Benchmark]
    [BenchmarkCategory("case:Reader")]
    public byte[] OfficeIMO_Reader() {
        using var input = new MemoryStream(_xlsx, writable: false);
        using var output = new MemoryStream();
        using (var reader = ExcelDocument.OpenReader(input)) {
            CsvDocument.WriteDataReader(output, reader);
        }
        return output.ToArray();
    }

    [Benchmark]
    [BenchmarkCategory("case:Default")]
    public byte[] MiniExcel() {
        using var input = new MemoryStream(_xlsx, writable: false);
        using var output = new MemoryStream();
        MiniExcelApi.ConvertXlsxToCsv(input, output);
        return output.ToArray();
    }

    [Benchmark]
    [BenchmarkCategory("case:Default")]
    public byte[] Sylvan() {
        using var input = new MemoryStream(_xlsx, writable: false);
        using var output = new MemoryStream();
        using (var reader = global::Sylvan.Data.Excel.ExcelDataReader.Create(input, global::Sylvan.Data.Excel.ExcelWorkbookType.ExcelXml))
        using (var text = new StreamWriter(output, new UTF8Encoding(false), leaveOpen: true))
        using (var writer = global::Sylvan.Data.Csv.CsvDataWriter.Create(text)) {
            writer.Write(reader);
        }
        return output.ToArray();
    }
}

internal static class ExcelCsvConversionData {
    private static readonly string[] Headers = ["Name", "Id", "CreatedOn", "Amount"];

    internal static byte[] CreateCsv(int rows) {
        var builder = new StringBuilder(rows * 40);
        builder.Append("Name,Id,CreatedOn,Amount\n");
        for (int index = 0; index < rows; index++) {
            var record = ExcelReaderGeneratedBenchmarks.CreateRecord(index);
            builder.Append(record.Name).Append(',')
                .Append(record.Id.ToString(CultureInfo.InvariantCulture)).Append(',')
                .Append(record.CreatedOn.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)).Append(',')
                .Append(record.Amount.ToString(CultureInfo.InvariantCulture)).Append('\n');
        }
        return Encoding.UTF8.GetBytes(builder.ToString());
    }

    internal static byte[] CreateWorkbook(int rows) {
        var records = new ExcelReaderGeneratedBenchmarks.GeneratedRecord[rows];
        for (int index = 0; index < rows; index++) records[index] = ExcelReaderGeneratedBenchmarks.CreateRecord(index);
        using var output = new MemoryStream();
        using (var writer = ExcelDocument.CreateWriter(output, new ExcelWriterOptions { CreateTable = false }, leaveOpen: true)) {
            writer.WriteTable("Sheet1", records, Headers,
                static (row, value) => row.Write(value.Name).Write(value.Id).Write(value.CreatedOn).Write(value.Amount));
        }
        return output.ToArray();
    }

    internal static void ValidateWorkbook(string lane, byte[] workbook, int rows) {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        using var stream = new MemoryStream(workbook);
        using var reader = global::ExcelDataReader.ExcelReaderFactory.CreateReader(stream);
        if (!reader.Read() || reader.FieldCount != Headers.Length) throw new InvalidDataException($"{lane}: missing header row.");
        for (int column = 0; column < Headers.Length; column++) {
            if (!string.Equals(Convert.ToString(reader.GetValue(column), CultureInfo.InvariantCulture), Headers[column], StringComparison.Ordinal))
                throw new InvalidDataException($"{lane}: header {column + 1} is '{reader.GetValue(column)}'.");
        }
        int count = 0;
        while (reader.Read()) {
            if (count >= rows) throw new InvalidDataException($"{lane}: more than {rows} data rows.");
            var expected = ExcelReaderGeneratedBenchmarks.CreateRecord(count);
            object? name = reader.GetValue(0), id = reader.GetValue(1), amount = reader.GetValue(3);
            // Text and inferred-number cells are both accepted; the value must match.
            if (!string.Equals(Convert.ToString(name, CultureInfo.InvariantCulture), expected.Name, StringComparison.Ordinal)
                || !string.Equals(Convert.ToString(id, CultureInfo.InvariantCulture), expected.Id.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal)
                || !double.TryParse(Convert.ToString(amount, CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
                || parsed != expected.Amount) {
                throw new InvalidDataException($"{lane}: data row {count + 1} is {name}/{id}/{amount}.");
            }
            count++;
        }
        if (count != rows) throw new InvalidDataException($"{lane}: {count} data rows, expected {rows}.");
    }

    internal static void ValidateCsv(string lane, byte[] csv, int rows) {
        string[] lines = Encoding.UTF8.GetString(csv).TrimEnd('\r', '\n').Split('\n');
        if (lines.Length != rows + 1) throw new InvalidDataException($"{lane}: {lines.Length} lines, expected {rows + 1}.");
        if (!Normalize(lines[0]).StartsWith(string.Join(",", Headers), StringComparison.Ordinal))
            throw new InvalidDataException($"{lane}: header line {lines[0]}.");
        for (int index = 0; index < rows; index++) {
            var record = ExcelReaderGeneratedBenchmarks.CreateRecord(index);
            string expected = record.Name + "," + record.Id.ToString(CultureInfo.InvariantCulture) + ",";
            if (!Normalize(lines[index + 1]).StartsWith(expected, StringComparison.Ordinal)) {
                throw new InvalidDataException($"{lane}: line {index + 2} is {lines[index + 1]}.");
            }
        }

        static string Normalize(string line) => line.TrimStart('﻿').TrimEnd('\r').Replace("\"", string.Empty);
    }
}
#endif
