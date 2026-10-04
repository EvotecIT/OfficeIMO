using System.Data.Common;
using System.Globalization;
using System.Text;
using System.Xml;
using BenchmarkDotNet.Attributes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Benchmarks;
using Sylvan.Data.Excel;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>Reads small string tables whose final item requires full XML decoding.</summary>
[MemoryDiagnoser]
public class ExcelSharedStringReadBenchmarks {
    private const int RowCount = 256;
    private byte[] _workbook = [];
    private string[] _expected = [];

    [Params("Ascii", "UnicodeTail", "EntityTail")]
    public string Shape { get; set; } = "Ascii";

    [GlobalSetup]
    public void Setup() {
        string? priority = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_PROCESS_PRIORITY");
        if (!string.IsNullOrEmpty(priority)) BenchmarkProcessorAffinity.ApplyPriority(priority);
        _expected = Enumerable.Range(0, RowCount)
            .Select(index => "Label" + index.ToString("D4", CultureInfo.InvariantCulture)).ToArray();
        _expected[^1] = Shape switch {
            "Ascii" => _expected[^1],
            "UnicodeTail" => "Łódź 東京 😀",
            "EntityTail" => "A&B <last>",
            _ => throw new InvalidOperationException("Unknown shared-string shape.")
        };
        using var stream = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook)) {
            WorkbookPart workbook = document.AddWorkbookPart();
            WorksheetPart worksheet = workbook.AddNewPart<WorksheetPart>();
            WriteWorksheet(worksheet);
            workbook.Workbook = new Workbook(new Sheets(new Sheet {
                Id = workbook.GetIdOfPart(worksheet), SheetId = 1, Name = "Data"
            }));
            SharedStringTablePart strings = workbook.AddNewPart<SharedStringTablePart>();
            // Keep the ordinary Excel declaration and element layout, with the
            // uncommon text at the end so late fallback remains measurable.
            using var writer = new StreamWriter(strings.GetStream(), new UTF8Encoding(false));
            writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            writer.Write("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"256\" uniqueCount=\"256\">");
            foreach (string value in _expected) {
                writer.Write("<si><t>");
                writer.Write(System.Security.SecurityElement.Escape(value));
                writer.Write("</t></si>");
            }
            writer.Write("</sst>");
        }
        _workbook = stream.ToArray();
        using (DbDataReader reader = OpenOfficeIMO()) Validate(reader);
        using var peerStream = new MemoryStream(_workbook, writable: false);
        using DbDataReader peer = OpenSylvan(peerStream);
        Validate(peer);
    }

    private static void WriteWorksheet(WorksheetPart part) {
        const string ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        using var writer = XmlWriter.Create(part.GetStream(), new XmlWriterSettings { Encoding = new UTF8Encoding(false) });
        writer.WriteStartElement("worksheet", ns);
        writer.WriteStartElement("dimension", ns);
        writer.WriteAttributeString("ref", "A1:A256");
        writer.WriteEndElement();
        writer.WriteStartElement("sheetData", ns);
        for (int index = 0; index < RowCount; index++) {
            string row = (index + 1).ToString(CultureInfo.InvariantCulture);
            writer.WriteStartElement("row", ns);
            writer.WriteAttributeString("r", row);
            writer.WriteStartElement("c", ns);
            writer.WriteAttributeString("r", "A" + row);
            writer.WriteAttributeString("t", "s");
            writer.WriteElementString("v", ns, index.ToString(CultureInfo.InvariantCulture));
            writer.WriteEndElement();
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    [Benchmark(Baseline = true)]
    public long OfficeIMO() {
        using DbDataReader reader = OpenOfficeIMO();
        return Observe(reader);
    }

    [Benchmark]
    public long Sylvan() {
        using var stream = new MemoryStream(_workbook, writable: false);
        using DbDataReader reader = OpenSylvan(stream);
        return Observe(reader);
    }

    private DbDataReader OpenOfficeIMO() => ExcelDocument.OpenDataReader(
        _workbook, new ExcelReadOptions { HasHeaderRow = false });

    private static DbDataReader OpenSylvan(Stream stream) => global::Sylvan.Data.Excel.ExcelDataReader.Create(
        stream, ExcelWorkbookType.ExcelXml, new ExcelDataReaderOptions { Schema = ExcelSchema.NoHeaders });

    private void Validate(DbDataReader reader) {
        foreach (string expected in _expected) {
            if (!reader.Read() || reader.FieldCount != 1 || reader.GetString(0) != expected)
                throw new InvalidDataException("Shared-string read changed a row value or count.");
        }
        if (reader.Read()) throw new InvalidDataException("Shared-string read returned extra rows.");
    }

    private static long Observe(DbDataReader reader) {
        long signature = 0;
        int rows = 0;
        while (reader.Read()) {
            string value = reader.GetString(0);
            foreach (char character in value) signature = unchecked(signature * 31 + character);
            rows++;
        }
        if (rows != RowCount) throw new InvalidDataException("Shared-string row count changed.");
        return signature;
    }
}
