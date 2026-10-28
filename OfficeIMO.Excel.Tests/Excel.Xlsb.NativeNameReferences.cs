using OfficeIMO.Excel;
using OfficeIMO.Excel.Xlsb;
using OfficeIMO.Excel.Xlsb.Biff12;
using System.IO.Compression;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelXlsbNativeNameReferencesTests {
    [Fact]
    public void Deleting_local_names_does_not_retarget_validation_condition_or_unknown_future_formulas() {
        using var seed = ExcelDocument.Create();
        var first = seed.AddSheet("First"); first.CellValue(1, 1, 10);
        var retained = seed.AddSheet("Retained"); retained.CellValue(1, 1, 20);
        seed.SetNamedRange("RemovedLocal", "A1", first, save: false);
        seed.SetNamedRange("RetainedValue", "'Retained'!A1", save: false);
        byte[] original = seed.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
        var source = XlsbWorkbookReader.Load(original);
        uint nameId = checked((uint)source.DefinedNames.ToList().FindIndex(name => name.Name == "RetainedValue") + 1);
        byte[] expression = new byte[] { 0x43 }.Concat(BitConverter.GetBytes(nameId)).Concat(new byte[] { 0x1E, 0, 0, 0x0D }).ToArray();
        byte[] validation = Validation(expression);
        byte[] condition = Condition(expression);
        byte[] future = Future(expression);
        byte[] worksheet;
        using (var archive = new ZipArchive(new MemoryStream(original), ZipArchiveMode.Read)) {
            using var input = archive.GetEntry("xl/worksheets/sheet2.bin")!.Open(); var records = XlsbRecordReader.ReadAll(input);
            using var output = new MemoryStream();
            foreach (var record in records) {
                if (record.Type == 130) {
                    XlsbRecordWriter.Write(output, 64, validation);
                    XlsbRecordWriter.Write(output, 463, condition);
                    XlsbRecordWriter.Write(output, 464);
                    XlsbRecordWriter.Write(output, 35, new byte[4]);
                    XlsbRecordWriter.Write(output, 0x3FFE, future);
                    XlsbRecordWriter.Write(output, 36);
                }
                XlsbRecordWriter.Write(output, record.Type, record.Data);
            }
            worksheet = output.ToArray();
        }
        byte[] modified;
        using (var output = new MemoryStream()) {
            output.Write(original, 0, original.Length);
            using (var archive = new ZipArchive(output, ZipArchiveMode.Update, leaveOpen: true)) {
                archive.GetEntry("xl/worksheets/sheet2.bin")!.Delete();
                using var target = archive.CreateEntry("xl/worksheets/sheet2.bin").Open(); target.Write(worksheet, 0, worksheet.Length);
            }
            modified = output.ToArray();
        }
        using var document = ExcelDocument.Load(new MemoryStream(modified));
        document.RemoveWorksheet("First");
        byte[] saved = document.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
        var loaded = XlsbWorkbookReader.Load(saved);
        Assert.Single(loaded.DefinedNames);
        using var result = new ZipArchive(new MemoryStream(saved), ZipArchiveMode.Read);
        using var resultPart = result.GetEntry("xl/worksheets/sheet2.bin")!.Open();
        var actual = XlsbRecordReader.ReadAll(resultPart);
        foreach (var item in new[] { (Type: 64, Start: 44), (Type: 463, Start: 50), (Type: 0x3FFE, Start: 20) }) {
            byte[] bytes = actual.Single(record => record.Type == item.Type).Data;
            byte[] tokens = bytes.Skip(item.Start).Take(expression.Length).ToArray();
            Assert.True(XlsbFormulaTextReader.TryRead(tokens, out string? text, loaded));
            Assert.Equal("RetainedValue>0", text);
        }
        Assert.Equal(new byte[] { 0xAB, 0xCD }, actual.Single(record => record.Type == 0x3FFE).Data.Skip(actual.Single(record => record.Type == 0x3FFE).Data.Length - 2));
    }

    private static void Write(Stream output, byte[] bytes) => output.Write(bytes, 0, bytes.Length);

    private static byte[] Parsed(byte[] tokens) => BitConverter.GetBytes(tokens.Length).Concat(tokens).Concat(new byte[4]).ToArray();
    private static byte[] Validation(byte[] tokens) {
        using var output = new MemoryStream();
        Write(output, BitConverter.GetBytes(7)); Write(output, BitConverter.GetBytes(1)); Write(output, new byte[16]);
        for (int index = 0; index < 4; index++) Write(output, BitConverter.GetBytes(uint.MaxValue));
        Write(output, Parsed(tokens)); Write(output, Parsed(Array.Empty<byte>())); return output.ToArray();
    }
    private static byte[] Condition(byte[] tokens) {
        using var output = new MemoryStream(); byte[] header = new byte[42];
        BitConverter.GetBytes(2).CopyTo(header, 0); BitConverter.GetBytes(1).CopyTo(header, 12); BitConverter.GetBytes(tokens.Length).CopyTo(header, 30);
        Write(output, header); Write(output, BitConverter.GetBytes(uint.MaxValue)); Write(output, Parsed(tokens)); return output.ToArray();
    }
    private static byte[] Future(byte[] tokens) {
        using var output = new MemoryStream();
        Write(output, BitConverter.GetBytes(4)); Write(output, BitConverter.GetBytes(1)); Write(output, BitConverter.GetBytes(2));
        Write(output, BitConverter.GetBytes(tokens.Length)); Write(output, new byte[4]); Write(output, tokens); Write(output, new byte[] { 0xAB, 0xCD });
        return output.ToArray();
    }
}
