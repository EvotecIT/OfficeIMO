using System.Xml.Linq;
using System.IO.Compression;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public class ExcelTextPersistenceContractTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExplicitFontFamilyReplacesThemeSchemeWithoutChangingSiblingFont(bool major) {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Fonts");
        sheet.Cell(1, 1).SetValue("Change").SetBold();
        sheet.Cell(1, 2).SetValue("Keep").SetBold();
        using var source = new MemoryStream();
        document.Save(source);
        source.Position = 0;
        using (var package = SpreadsheetDocument.Open(source, true)) {
            var styles = package.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            var cell = package.WorkbookPart.WorksheetParts.Single().Worksheet.Descendants<Cell>().First();
            var format = styles.CellFormats!.Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value);
            var font = styles.Fonts!.Elements<Font>().ElementAt((int)format.FontId!.Value);
            font.RemoveAllChildren<FontScheme>();
            font.AddChild(new FontScheme { Val = major ? FontSchemeValues.Major : FontSchemeValues.Minor }, true);
        }
        source.Position = 0;
        using var loaded = ExcelDocument.Load(source);
        loaded.Sheets[0].Cell(1, 1).SetFontFamily("Arial");
        using var output = SpreadsheetDocument.Open(new MemoryStream(loaded.ToBytes()), false);
        var workbook = output.WorkbookPart!;
        var writtenStyles = workbook.WorkbookStylesPart!.Stylesheet!;
        var cells = workbook.WorksheetParts.Single().Worksheet.Descendants<Cell>().ToArray();
        Font Resolve(Cell cell) => writtenStyles.Fonts!.Elements<Font>().ElementAt((int)writtenStyles.CellFormats!.Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value).FontId!.Value);
        Font changed = Resolve(cells[0]);
        Assert.Equal("Arial", changed.FontName!.Val!.Value);
        Assert.Null(changed.FontScheme);
        Assert.NotNull(changed.Bold);
        Assert.Equal(major ? FontSchemeValues.Major : FontSchemeValues.Minor, Resolve(cells[1]).FontScheme!.Val!.Value);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DomSharedStringsPersistEdgeWhitespaceWithPreserveSpace(bool standardWriter) {
        string[] values = { "text\n", "text\r\n", " text ", "\ttext\t", "Łódź 😀\n" };
        using var source = new MemoryStream();
        using (var initial = ExcelDocument.Create()) {
            initial.AddSheet("Text").Cell(1, 1).SetValue("template");
            initial.Save(source);
        }
        source.Position = 0;
        using var document = ExcelDocument.Load(source);
        var sheet = document.Sheets[0];
        for (int index = 0; index < values.Length; index++) sheet.Cell(index + 1, 1).SetValue(values[index]);
        using var saved = new MemoryStream();
        document.Save(saved, new ExcelSaveOptions { DisableFastPackageWriter = standardWriter });
        saved.Position = 0;
        using var zip = new ZipArchive(saved, ZipArchiveMode.Read);
        using var xmlStream = zip.GetEntry("xl/sharedStrings.xml")!.Open();
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var texts = XDocument.Load(xmlStream, LoadOptions.PreserveWhitespace).Descendants(ns + "t").ToArray();
        foreach (string value in values) {
            var text = Assert.Single(texts, item => item.Value == value);
            Assert.Equal("preserve", (string?)text.Attribute(XNamespace.Xml + "space"));
        }
    }
}
