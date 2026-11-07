using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Excel;

namespace OfficeIMO.LegacyImport.Tests;

public sealed partial class EarlyBiffImportTests {
    [Theory]
    [InlineData("matched", "")]
    [InlineData("size", "Second")]
    [InlineData("two-later", "Second|Third")]
    [InlineData("return-to-first", "Second")]
    [InlineData("face", "Second")]
    [InlineData("color-only", "")]
    [InlineData("family-case", "")]
    [InlineData("fontless-matched", "")]
    [InlineData("fontless-mismatch", "Third")]
    [InlineData("ignored-nonworksheets", "")]
    public void WorkbookReportsEachWorksheetWhoseDefaultFontChangesGeometry(string scenario, string locations) {
        using var source = new MemoryStream(BuildDefaultFontWorkbook(scenario));
        using ExcelDocument document = ExcelDocument.Load(source, null, out OfficeConversionReport report);
        OfficeDiagnostic[] fontFindings = report.Diagnostics.Where(finding => finding.Code == "early-biff-default-font-metrics").ToArray();
        string[] expectedLocations = locations.Length == 0 ? Array.Empty<string>() : locations.Split('|');
        Assert.Equal(expectedLocations, fontFindings.Select(finding => finding.Location));
        Assert.All(fontFindings, finding => {
            Assert.Equal(OfficeDiagnosticSeverity.Warning, finding.Severity);
            Assert.Equal(OfficeConversionLossKind.Approximation, finding.LossKind);
        });
        bool hasOtherLoss = scenario == "ignored-nonworksheets";
        using var output = new MemoryStream();
        if (expectedLocations.Length != 0 || hasOtherLoss) {
            Assert.Throws<NotSupportedException>(() => document.Save(output, new ExcelSaveOptions { AllowLoss = false }));
            Assert.Equal(0, output.Length);
        } else document.Save(output, new ExcelSaveOptions { AllowLoss = false, ValidateOpenXml = true });
        document.Save(output, new ExcelSaveOptions { AllowLoss = true, ValidateOpenXml = true });
        output.Position = 0;
        using var package = new ZipArchive(output, ZipArchiveMode.Read);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        using Stream workbookPart = package.GetEntry("xl/workbook.xml")!.Open();
        string[] expectedNames = scenario is "size" or "face" or "color-only" or "family-case"
            ? new[] { "First", "Second" } : new[] { "First", "Second", "Third" };
        Assert.Equal(expectedNames, XDocument.Load(workbookPart).Descendants(ns + "sheet").Select(sheet => (string?)sheet.Attribute("name")));
        for (int index = 0; index < expectedNames.Length; index++) {
            using Stream worksheet = package.GetEntry($"xl/worksheets/sheet{index + 1}.xml")!.Open();
            XElement cell = Assert.Single(XDocument.Load(worksheet).Descendants(ns + "c"));
            Assert.Equal("A1", (string?)cell.Attribute("r"));
            Assert.Equal("n", (string?)cell.Attribute("t") ?? "n");
            Assert.Equal((index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture), cell.Element(ns + "v")!.Value);
            if (scenario == "color-only" && index == 1) {
                using Stream stylesPart = package.GetEntry("xl/styles.xml")!.Open();
                XElement styles = XDocument.Load(stylesPart).Root!;
                XElement format = styles.Element(ns + "cellXfs")!.Elements(ns + "xf").ElementAt((int?)cell.Attribute("s") ?? 0);
                XElement[] fonts = styles.Element(ns + "fonts")!.Elements(ns + "font").ToArray();
                Assert.Equal("FF000000", (string?)fonts[0].Element(ns + "color")!.Attribute("rgb"));
                Assert.Equal("FFFF0000", (string?)fonts[(int)format.Attribute("fontId")!].Element(ns + "color")!.Attribute("rgb"));
            }
        }
    }

    private static byte[] BuildDefaultFontWorkbook(string scenario) {
        var fonts = new (string? Family, ushort Height, ushort Flags, ushort Color)[] {
            ("Arial", 220, 0, 8), ("Arial", 220, 0, 8), ("Arial", 220, 0, 8)
        };
        switch (scenario) {
            case "size": fonts = fonts.Take(2).ToArray(); fonts[1].Height = 440; break;
            case "two-later": fonts[1].Height = 440; fonts[2].Height = 440; break;
            case "return-to-first": fonts[1].Family = "Courier New"; break;
            case "face": fonts = fonts.Take(2).ToArray(); fonts[1].Flags = 3; break;
            case "color-only": fonts = fonts.Take(2).ToArray(); fonts[1].Color = 10; break;
            case "family-case": fonts = fonts.Take(2).ToArray(); fonts[1].Family = "aRiAl"; break;
            case "fontless-matched":
            case "ignored-nonworksheets": fonts[0].Family = null; break;
            case "fontless-mismatch": fonts[0].Family = null; fonts[2].Height = 440; break;
        }
        using var bytes = new MemoryStream();
        using var writer = new BinaryWriter(bytes);
        Record(writer, BofCode(4), record => WriteBof(record, 4, 0x0100));
        Record(writer, 0x008E, record => record.Write(0));
        string[] names = { "First", "Second", "Third" };
        bool ignored = scenario == "ignored-nonworksheets";
        foreach (string name in (ignored ? new[] { "Macro", "Chart" } : Array.Empty<string>()).Concat(names.Take(fonts.Length))) {
            Record(writer, 0x0085, record => WriteString(record, Encoding.ASCII.GetBytes(name), shortLength: true));
        }
        long firstSheet = bytes.Position;
        bytes.Position = 14;
        writer.Write((int)firstSheet);
        bytes.Position = firstSheet;
        if (ignored) {
            WriteSheet("Macro", 0x0040, ("Courier New", 440, 0, 8), 0);
            WriteSheet("Chart", 0x0020, ("Times New Roman", 660, 0, 8), 0);
        }
        for (int index = 0; index < fonts.Length; index++) WriteSheet(names[index], 0x0010, fonts[index], index + 1);
        Record(writer, 0x000A, _ => { });
        return bytes.ToArray();

        void WriteSheet(string name, ushort kind, (string? Family, ushort Height, ushort Flags, ushort Color) font, int value) {
            using var sheetBytes = new MemoryStream();
            using (var sheet = new BinaryWriter(sheetBytes, Encoding.UTF8, leaveOpen: true)) {
                Record(sheet, BofCode(4), record => WriteBof(record, 4, kind));
                if (font.Family != null) Record(sheet, 0x0231, record => {
                    record.Write(font.Height); record.Write(font.Flags); record.Write(font.Color);
                    WriteString(record, Encoding.ASCII.GetBytes(font.Family), shortLength: true);
                });
                Record(sheet, 0x0443, record => record.Write(new byte[] { 0, 0, 0xF1, 0xFF, 0, 0xFC, 0, 0, 0, 0, 0, 0 }));
                Record(sheet, 0x0055, record => record.Write((ushort)10));
                Record(sheet, 0x0099, record => record.Write((ushort)3072));
                Record(sheet, 0x007D, record => {
                    record.Write((ushort)0); record.Write((ushort)0); record.Write((ushort)3328);
                    record.Write((ushort)0); record.Write((ushort)0); record.Write((ushort)0);
                });
                Record(sheet, 0x0203, record => {
                    record.Write((ushort)0); record.Write((ushort)0); record.Write((ushort)0); record.Write((double)value);
                });
                Record(sheet, 0x000A, _ => { });
            }
            Record(writer, 0x008F, record => {
                record.Write((int)sheetBytes.Length);
                WriteString(record, Encoding.ASCII.GetBytes(name), shortLength: true);
            });
            writer.Write(sheetBytes.ToArray());
        }
    }

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
