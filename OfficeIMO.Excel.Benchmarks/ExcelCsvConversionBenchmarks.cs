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
/// output with ExcelDataReader and checks the row count and the last row's name and id.
/// </summary>
[MemoryDiagnoser]
public class CsvToXlsxConversionBenchmarks {
    [Params(50_000)]
    public int Rows { get; set; }

    private byte[] _csv = [];

    [GlobalSetup]
    public void Setup() {
        _csv = ExcelCsvConversionData.CreateCsv(Rows);
        ExcelCsvConversionData.ValidateWorkbook(nameof(OfficeIMO_ImportCsv), OfficeIMO_ImportCsv(), Rows);
        ExcelCsvConversionData.ValidateWorkbook(nameof(OfficeIMO_StreamingText), OfficeIMO_StreamingText(), Rows);
        ExcelCsvConversionData.ValidateWorkbook(nameof(MiniExcel), MiniExcel(), Rows);
        ExcelCsvConversionData.ValidateWorkbook(nameof(Sylvan), Sylvan(), Rows);
    }

    /// <summary>The editable import on its defaults: inferred column types and a native table.</summary>
    [Benchmark(Baseline = true)]
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
    public byte[] MiniExcel() {
        using var input = new MemoryStream(_csv, writable: false);
        using var output = new MemoryStream();
        MiniExcelApi.ConvertCsvToXlsx(input, output);
        return output.ToArray();
    }

    [Benchmark]
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
/// A one-sheet XLSX workbook (name, id, date, amount) converted to in-memory CSV. Setup checks the line count and
/// the last line's name and id in every output.
/// </summary>
[MemoryDiagnoser]
public class XlsxToCsvConversionBenchmarks {
    [Params(50_000)]
    public int Rows { get; set; }

    private byte[] _xlsx = [];

    [GlobalSetup]
    public void Setup() {
        _xlsx = ExcelCsvConversionData.CreateWorkbook(Rows);
        ExcelCsvConversionData.ValidateCsv(nameof(OfficeIMO_SaveAsCsv), OfficeIMO_SaveAsCsv(), Rows);
        ExcelCsvConversionData.ValidateCsv(nameof(OfficeIMO_Reader), OfficeIMO_Reader(), Rows);
        ExcelCsvConversionData.ValidateCsv(nameof(MiniExcel), MiniExcel(), Rows);
        ExcelCsvConversionData.ValidateCsv(nameof(Sylvan), Sylvan(), Rows);
    }

    /// <summary>The editable route: load the workbook and save its sheet as CSV.</summary>
    [Benchmark(Baseline = true)]
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
    public byte[] OfficeIMO_Reader() {
        using var input = new MemoryStream(_xlsx, writable: false);
        using var output = new MemoryStream();
        using (var reader = ExcelDocument.OpenReader(input)) {
            CsvDocument.WriteDataReader(output, reader);
        }
        return output.ToArray();
    }

    [Benchmark]
    public byte[] MiniExcel() {
        using var input = new MemoryStream(_xlsx, writable: false);
        using var output = new MemoryStream();
        MiniExcelApi.ConvertXlsxToCsv(input, output);
        return output.ToArray();
    }

    [Benchmark]
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
        int count = 0;
        object? name = null, id = null;
        while (reader.Read()) {
            count++;
            name = reader.GetValue(0);
            id = reader.GetValue(1);
        }
        var last = ExcelReaderGeneratedBenchmarks.CreateRecord(rows - 1);
        if (count != rows + 1
            || !string.Equals(Convert.ToString(name, CultureInfo.InvariantCulture), last.Name, StringComparison.Ordinal)
            || !string.Equals(Convert.ToString(id, CultureInfo.InvariantCulture), last.Id.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal)) {
            throw new InvalidDataException($"{lane}: {count} rows, last row {name}/{id}.");
        }
    }

    internal static void ValidateCsv(string lane, byte[] csv, int rows) {
        string[] lines = Encoding.UTF8.GetString(csv).TrimEnd('\r', '\n').Split('\n');
        var last = ExcelReaderGeneratedBenchmarks.CreateRecord(rows - 1);
        string expected = last.Name + "," + last.Id.ToString(CultureInfo.InvariantCulture) + ",";
        string lastLine = lines[^1].TrimStart('﻿').Replace("\"", string.Empty);
        if (lines.Length != rows + 1 || !lastLine.StartsWith(expected, StringComparison.Ordinal)) {
            throw new InvalidDataException($"{lane}: {lines.Length} lines, last line {lines[^1]}.");
        }
    }
}
#endif
