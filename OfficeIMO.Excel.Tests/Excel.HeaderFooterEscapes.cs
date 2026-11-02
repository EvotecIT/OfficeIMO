using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Theory]
    [InlineData(true, "R&&L", "&LR&&L&CTitle&R&G")]
    [InlineData(false, "R&&C", "&LR&&C&CTitle&R&G")]
    [InlineData(true, "R&&R", "&LR&&R&CTitle&R&G")]
    [InlineData(false, "&\"Acme &L,Regular\"R", "&L&\"Acme &L,Regular\"R&CTitle&R&G")]
    [InlineData(false, "&\"Acme & Sons,Regular\"R", "&L&\"Acme & Sons,Regular\"R&CTitle&R&G")]
    [InlineData(true, "First&LSecond", "&LFirstSecond&CTitle&R&G")]
    public void HeaderFooterSectionEditsAndImagesPreserveEscapedAndRepeatedSectionContent(bool header, string left, string expected) {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Report");
        byte[] image = HeaderFooterEscapesTestImage();
        if (header) {
            sheet.SetHeader(left: left);
            sheet.SetHeader(center: "Title");
            sheet.SetHeaderImage(ExcelHeaderFooterPosition.Right, image, width: OfficeLength.Points(12), height: OfficeLength.Points(6));
        } else {
            sheet.SetFooter(left: left);
            sheet.SetFooter(center: "Title");
            sheet.SetFooterImage(ExcelHeaderFooterPosition.Right, image, width: OfficeLength.Points(12), height: OfficeLength.Points(6));
        }
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XElement value = ReadSavedHeaderFooter(document);
        Assert.Equal(expected, value.Element(ns + (header ? "oddHeader" : "oddFooter"))!.Value);
    }

    [Fact]
    public void HeaderFooterPictureMarkersDistinguishLiteralAmpersandsAndReleaseRemovedImages() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Report");
        sheet.SetHeader(left: "R&&G");
        Assert.False(sheet.GetHeaderFooter().HeaderHasPicturePlaceholder);
        byte[] image = HeaderFooterEscapesTestImage();
        sheet.SetHeaderImage(ExcelHeaderFooterPosition.Left, image, width: OfficeLength.Points(12), height: OfficeLength.Points(6));
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        Assert.Equal("&L&GR&&G", ReadSavedHeaderFooter(document).Element(ns + "oddHeader")!.Value);
        Assert.Equal(image, sheet.GetHeaderFooter().HeaderLeftImage!.Bytes);
        sheet.SetHeader(left: "R&&G");
        var info = sheet.GetHeaderFooter();
        Assert.False(info.HeaderHasPicturePlaceholder);
        Assert.Null(info.HeaderLeftImage);
        using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
        using var zip = new ZipArchive(saved, ZipArchiveMode.Read);
        Assert.DoesNotContain(zip.Entries, entry => entry.FullName.StartsWith("xl/drawings/", StringComparison.Ordinal)
            || entry.FullName.StartsWith("xl/media/", StringComparison.Ordinal));
        using var stream = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        XElement root = XDocument.Load(stream).Root!;
        Assert.Null(root.Element(ns + "legacyDrawingHF"));
        Assert.Equal("&LR&&G", root.Element(ns + "headerFooter")!.Element(ns + "oddHeader")!.Value);
    }

    private static byte[] HeaderFooterEscapesTestImage() => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

    private static XElement ReadSavedHeaderFooter(ExcelDocument document) {
        using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
        using var zip = new ZipArchive(saved, ZipArchiveMode.Read);
        using var stream = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        return XDocument.Load(stream).Root!.Element(ns + "headerFooter")!;
    }
}
