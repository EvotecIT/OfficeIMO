using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Excel;

namespace OfficeIMO.LegacyImport.Tests;

public sealed partial class EarlyBiffImportTests {
    [Fact]
    public void WorkbookRetainsEveryCellFindingAndDeduplicatesGlobalFindingsAcrossSheets() {
        using var bytes = new MemoryStream();
        using (var writer = new BinaryWriter(bytes, Encoding.UTF8, leaveOpen: true)) {
            Record(writer, BofCode(4), record => WriteBof(record, 4, 0x0100));
            Record(writer, 0x008E, record => record.Write(0));
            string[] names = { "First", "Second", "Third" };
            foreach (string name in names) Record(writer, 0x0085, record =>
                WriteString(record, Encoding.ASCII.GetBytes(name), shortLength: true));
            long firstSheet = bytes.Position;
            bytes.Position = 14;
            writer.Write((int)firstSheet);
            bytes.Position = firstSheet;
            for (int index = 0; index < names.Length; index++) {
                using var sheetBytes = new MemoryStream();
                using (var sheet = new BinaryWriter(sheetBytes, Encoding.UTF8, leaveOpen: true)) {
                    Record(sheet, BofCode(4), record => WriteBof(record, 4, 0x10));
                    if (index != 0) for (int style = 0; style < 3; style++)
                        Record(sheet, 0x0443, record => record.Write(new byte[] { 0, 1, 0xF1, 0xFF, 0, 0xFC, 0, 0, 0, 0, 0, 0 }));
                    int cells = index == 0 ? 2 : 1;
                    for (ushort row = 0; row < cells; row++) Record(sheet, 0x0203, record => {
                        record.Write(row); record.Write((ushort)0); record.Write((ushort)(index == 0 ? 1 : 0));
                        record.Write(index == 0 ? row + 1d : index + 2d);
                    });
                    Record(sheet, 0x000A, _ => { });
                }
                Record(writer, 0x008F, record => {
                    record.Write((int)sheetBytes.Length);
                    WriteString(record, Encoding.ASCII.GetBytes(names[index]), shortLength: true);
                });
                writer.Write(sheetBytes.ToArray());
            }
            Record(writer, 0x000A, _ => { });
        }
        bytes.Position = 0;
        using ExcelDocument document = ExcelDocument.Load(bytes, null, out OfficeConversionReport report);
        Assert.Equal(new[] { "early-biff-cell-style", "early-biff-cell-style", "early-biff-number-format" },
            report.Diagnostics.Select(finding => finding.Code));
        Assert.Equal(new string?[] { "'First'!A1", "'First'!A2", null }, report.Diagnostics.Select(finding => finding.Location));
        using var output = new MemoryStream();
        document.Save(output, new ExcelSaveOptions { AllowLoss = true });
        output.Position = 0;
        using var package = new ZipArchive(output, ZipArchiveMode.Read);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        string[][] expected = { new[] { "A1=1", "A2=2" }, new[] { "A1=3" }, new[] { "A1=4" } };
        for (int index = 0; index < expected.Length; index++) {
            using Stream worksheet = package.GetEntry($"xl/worksheets/sheet{index + 1}.xml")!.Open();
            Assert.Equal(expected[index], XDocument.Load(worksheet).Descendants(ns + "c")
                .Select(cell => (string?)cell.Attribute("r") + "=" + cell.Element(ns + "v")!.Value));
        }
    }
}
