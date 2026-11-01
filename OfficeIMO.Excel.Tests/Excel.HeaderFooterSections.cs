using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void HeaderFooterSectionEditsPreserveTheOtherHalfAndAuthoredPageVariants() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Report");
        sheet.SetHeaderFooter(headerLeft: "Left", headerCenter: "Title", headerRight: "&P",
            footerLeft: "Confidential", footerRight: "Page &P of &N", differentFirstPage: true,
            differentOddEven: true, alignWithMargins: false, scaleWithDoc: false);
        sheet.SetFirstPageHeaderFooter(headerCenter: "First title");
        sheet.SetEvenPageHeaderFooter(footerCenter: "Even footer");
        sheet.SetHeader(left: string.Empty, center: "Updated title");
        sheet.SetFooter(left: "Reviewed");
        sheet.SetHeader(differentFirstPage: true, differentOddEven: true);
        using var saved = new MemoryStream();
        document.Save(saved); saved.Position = 0;
        using (var reopened = ExcelDocument.Load(saved)) {
            var info = reopened.Sheets[0].GetHeaderFooter();
            Assert.Equal(string.Empty, info.HeaderLeft);
            Assert.Equal("Updated title", info.HeaderCenter);
            Assert.Equal("&P", info.HeaderRight);
            Assert.Equal("Reviewed", info.FooterLeft);
            Assert.Equal("Page &P of &N", info.FooterRight);
            Assert.Equal("First title", info.FirstHeaderCenter);
            Assert.Empty(info.FirstFooterLeft); Assert.Empty(info.FirstFooterCenter); Assert.Empty(info.FirstFooterRight);
            Assert.Empty(info.EvenHeaderLeft); Assert.Empty(info.EvenHeaderCenter); Assert.Empty(info.EvenHeaderRight);
            Assert.Equal("Even footer", info.EvenFooterCenter);
            Assert.True(info.DifferentFirstPage); Assert.True(info.DifferentOddEven);
        }
        saved.Position = 0;
        using (var zip = new ZipArchive(saved, ZipArchiveMode.Read, leaveOpen: true)) {
            using var stream = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open();
            var xml = XDocument.Load(stream);
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var runningText = xml.Root!.Element(ns + "headerFooter")!;
            Assert.Equal("0", (string?) runningText.Attribute("alignWithMargins"));
            Assert.Equal("0", (string?) runningText.Attribute("scaleWithDoc"));
        }
        sheet.SetHeader(differentFirstPage: false, alignWithMargins: true);
        using var disabled = new MemoryStream();
        document.Save(disabled); disabled.Position = 0;
        using var final = ExcelDocument.Load(disabled);
        var disabledInfo = final.Sheets[0].GetHeaderFooter();
        Assert.False(disabledInfo.DifferentFirstPage);
        Assert.True(disabledInfo.DifferentOddEven);
        Assert.Equal("First title", disabledInfo.FirstHeaderCenter);
        Assert.Empty(disabledInfo.FirstFooterLeft); Assert.Empty(disabledInfo.FirstFooterCenter); Assert.Empty(disabledInfo.FirstFooterRight);
        Assert.Empty(disabledInfo.EvenHeaderLeft); Assert.Empty(disabledInfo.EvenHeaderCenter); Assert.Empty(disabledInfo.EvenHeaderRight);
        Assert.Equal("Even footer", disabledInfo.EvenFooterCenter);
        Assert.Equal("Page &P of &N", disabledInfo.FooterRight);
        sheet.SetFooter(differentFirstPage: true);
        using var enabled = new MemoryStream();
        document.Save(enabled); enabled.Position = 0;
        using var enabledDocument = ExcelDocument.Load(enabled);
        var enabledInfo = enabledDocument.Sheets[0].GetHeaderFooter();
        Assert.True(enabledInfo.DifferentFirstPage);
        Assert.Equal("First title", enabledInfo.FirstHeaderCenter);
        Assert.Equal("Reviewed", enabledInfo.FirstFooterLeft);
        Assert.Equal("Page &P of &N", enabledInfo.FirstFooterRight);
        Assert.Empty(enabledInfo.EvenHeaderCenter);
    }
}
