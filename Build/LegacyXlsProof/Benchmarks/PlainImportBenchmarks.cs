using System.Security.Cryptography;
using BenchmarkDotNet.Attributes;
using OfficeIMO.Excel;

[MemoryDiagnoser]
public class LegacyXlsPlainImportBenchmarks {
    private static readonly ExcelSaveOptions NativeSave = new() { Format = ExcelFileFormat.Xls };
    private static readonly ExcelSaveOptions OpenXmlSave = new() { Format = ExcelFileFormat.Xlsx };
    private byte[] _source = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        string root = Environment.GetEnvironmentVariable("OFFICEIMO_LEGACY_XLS_PROOF_OUTPUT")
            ?? throw new InvalidOperationException("Point OFFICEIMO_LEGACY_XLS_PROOF_OUTPUT to the master capture.");
        _source = LoadFixture(root);

        if (ReadPlainCells() != 6) {
            throw new InvalidDataException("The source workbook does not contain Name in A1 and 42 in B2.");
        }
        using ExcelDocument saved = Open(ConvertToXlsx());
        using ExcelDocument edited = Open(EditNativeXls());
        if (saved.SourceFormat != ExcelFileFormat.Xlsx || !HasCell(saved, 1, 1, "Name")
            || !HasCell(saved, 2, 2, "42") || edited.SourceFormat != ExcelFileFormat.Xls
            || !HasCell(edited, 1, 1, "Changed") || !HasCell(edited, 2, 2, "42")) {
            throw new InvalidDataException("The benchmark did not preserve the expected cell values and output formats.");
        }
        Console.WriteLine($"Input XLS {Convert.ToHexString(SHA256.HashData(_source))}");
    }

    [Benchmark]
    public int ReadPlainCells() => ReadPlainCells(_source);

    internal static byte[] LoadFixture(string root) =>
        File.ReadAllBytes(Path.Combine(root, "test-files", "LegacyXls_NormalLoad_SaveAsyncPathProducesValidXlsx", "001.xls"));

    internal static int ReadPlainCells(byte[] bytes) {
        using ExcelDocument document = Open(bytes);
        if (document.SourceFormat != ExcelFileFormat.Xls || document.Sheets.Count != 1) {
            throw new InvalidDataException("The benchmark requires an XLS workbook with one sheet.");
        }
        ExcelSheet sheet = document.Sheets[0];
        if (!sheet.TryGetCellText(1, 1, out string? first) || first != "Name"
            || !sheet.TryGetCellText(2, 2, out string? second) || second != "42") {
            throw new InvalidDataException("The benchmark workbook must contain Name in A1 and 42 in B2.");
        }
        return first.Length + second.Length;
    }

    [Benchmark]
    public byte[] ConvertToXlsx() {
        using ExcelDocument document = Open(_source);
        return document.ToBytes(OpenXmlSave);
    }

    [Benchmark]
    public byte[] EditNativeXls() {
        using ExcelDocument document = Open(_source);
        document.Sheets[0].Cell(1, 1).SetValue("Changed");
        return document.ToBytes(NativeSave);
    }

    private static ExcelDocument Open(byte[] bytes) {
        using var stream = new MemoryStream(bytes, writable: false);
        return ExcelDocument.Load(stream);
    }

    private static bool HasCell(ExcelDocument document, int row, int column, string expected) =>
        document.Sheets[0].TryGetCellText(row, column, out string? value) && value == expected;
}
