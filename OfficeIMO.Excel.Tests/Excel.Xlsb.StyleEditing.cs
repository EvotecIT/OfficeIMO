using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Xlsb.Biff12;
using System.IO.Compression;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelXlsbStyleEditingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Added_styles_preserve_native_style_prefixes_and_chart_VBA_parts(bool encryptedSource) {
        string fixture = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", "ExcelPreservationCorpus", "excel-populated-macro.xlsb");
        byte[] original = File.ReadAllBytes(fixture);
        using var source = ExcelDocument.Load(new MemoryStream(original));
        using var encrypted = new MemoryStream();
        if (encryptedSource) source.SaveEncrypted(encrypted, "test-password", ExcelFileFormat.Xlsb);
        using var document = encryptedSource ? ExcelDocument.LoadEncrypted(encrypted, "test-password") : ExcelDocument.Load(new MemoryStream(original));
        var styles = document.WorkbookPartRoot.WorkbookStylesPart!.Stylesheet!;
        uint font = checked((uint)styles.Fonts!.ChildElements.Count);
        styles.Fonts.Append(new Font(new Bold(), new Color { Rgb = "FF123456" }));
        uint fill = checked((uint)styles.Fills!.ChildElements.Count);
        styles.Fills.Append(new Fill(new PatternFill(new ForegroundColor { Rgb = "FFDAE8FC" }, new BackgroundColor { Indexed = 64 }) { PatternType = PatternValues.Solid }));
        uint border = checked((uint)styles.Borders!.ChildElements.Count);
        styles.Borders.Append(new Border(new LeftBorder(new Color { Rgb = "FF123456" }) { Style = BorderStyleValues.Double }));
        styles.NumberingFormats ??= new NumberingFormats();
        uint numberFormat = Math.Max(164U, styles.NumberingFormats.Elements<NumberingFormat>().Select(item => item.NumberFormatId!.Value).DefaultIfEmpty(163U).Max() + 1);
        styles.NumberingFormats.Append(new NumberingFormat { NumberFormatId = numberFormat, FormatCode = "0.00000" });
        uint format = checked((uint)styles.CellFormats!.ChildElements.Count);
        styles.CellFormats.Append(new CellFormat { FontId = font, FillId = fill, BorderId = border, NumberFormatId = numberFormat, FormatId = 0, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyNumberFormat = true });
        document.Sheets[0].CellValue(200, 3, 1.23456);
        document.Sheets[0].SetCellStyleIndex(200, 3, format, save: false);
        byte[] saved = document.ToBytes(ExcelFileFormat.Xlsb);
        using var reloaded = ExcelDocument.Load(new MemoryStream(saved));
        var snapshot = reloaded.Sheets[0].GetCellStyle(200, 3);
        Assert.True(snapshot.Bold);
        Assert.Equal("0.00000", snapshot.NumberFormatCode);
        Assert.Equal("solid", snapshot.FillPatternType);
        var loadedStyles = reloaded.WorkbookPartRoot.WorkbookStylesPart!.Stylesheet!;
        Assert.Equal(BorderStyleValues.Double, loadedStyles.Borders!.Elements<Border>().ElementAt((int)border).LeftBorder!.Style!.Value);
        using var before = new ZipArchive(new MemoryStream(original), ZipArchiveMode.Read);
        using var after = new ZipArchive(new MemoryStream(saved), ZipArchiveMode.Read);
        foreach (var entry in before.Entries.Where(item => item.FullName != "xl/worksheets/sheet1.bin" && item.FullName != "xl/styles.bin"))
            Assert.Equal(Read(entry), Read(after.GetEntry(entry.FullName)!));
        var oldRecords = XlsbRecordReader.ReadAll(new MemoryStream(Read(before.GetEntry("xl/styles.bin")!)));
        var newRecords = XlsbRecordReader.ReadAll(new MemoryStream(Read(after.GetEntry("xl/styles.bin")!)));
        foreach (int type in new[] { 43, 45, 46, 47, 48 }) {
            var prefix = oldRecords.Where(record => record.Type == type).Select(record => record.Data).ToArray();
            Assert.Equal(prefix, newRecords.Where(record => record.Type == type).Take(prefix.Length).Select(record => record.Data));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Native_locale_number_format_records_survive_appended_styles(bool encrypted) {
        using var seed = ExcelDocument.Create();
        var sheet = seed.AddWorksheet("Data"); sheet.CellValue(1, 1, 12.5D); sheet.CellItalic(1, 1);
        var styles = seed.WorkbookPartRoot.WorkbookStylesPart!.Stylesheet!;
        styles.NumberingFormats ??= new NumberingFormats();
        styles.NumberingFormats.Append(new NumberingFormat { NumberFormatId = 5, FormatCode = "\"PLN\" #,##0.00" });
        styles.NumberingFormats.Count = (uint)styles.NumberingFormats.ChildElements.Count;
        byte[] original = seed.ToBytes(ExcelFileFormat.Xlsb);
        using var document = ExcelDocument.Load(new MemoryStream(original));
        document.Sheets[0].CellBold(1, 1);
        using var output = new MemoryStream();
        if (encrypted) document.SaveEncrypted(output, "test-password", ExcelFileFormat.Xlsb);
        else document.Save(output, ExcelFileFormat.Xlsb);
        using var copy = encrypted ? ExcelDocument.LoadEncrypted(output, "test-password") : ExcelDocument.Load(output);
        Assert.True(copy.Sheets[0].CellAt(1, 1).GetStyle().Bold);
        var retained = Assert.Single(copy.WorkbookPartRoot.WorkbookStylesPart!.Stylesheet!.NumberingFormats!.Elements<NumberingFormat>());
        Assert.Equal(5U, retained.NumberFormatId!.Value);
        Assert.Equal("\"PLN\" #,##0.00", retained.FormatCode!.Value);
    }

    [Fact]
    public void Projected_native_chartsheets_have_valid_sheet_views() {
        string fixture = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", "ExcelPreservationCorpus", "excel-populated-macro.xlsb");
        using var document = ExcelDocument.Load(fixture);
        var parts = document.WorkbookPartRoot.ChartsheetParts.ToArray();
        Assert.NotEmpty(parts);
        var validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator();
        foreach (var part in parts) {
            Assert.Equal(0U, Assert.Single(part.Chartsheet!.GetFirstChild<ChartSheetViews>()!.Elements<ChartSheetView>()).WorkbookViewId!.Value);
            Assert.Empty(validator.Validate(part.Chartsheet));
        }
    }

    private static byte[] Read(ZipArchiveEntry entry) {
        using var input = entry.Open(); using var output = new MemoryStream(); input.CopyTo(output); return output.ToArray();
    }
}
