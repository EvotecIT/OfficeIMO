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
        sheet.PageSetup.Header.Left.Text = "Left";
        sheet.PageSetup.Header.Center.Text = "Title";
        sheet.PageSetup.Header.Right.Text = "&P";
        sheet.PageSetup.Footer.Left.Text = "Confidential";
        sheet.PageSetup.Footer.Right.Text = "Page &P of &N";
        sheet.PageSetup.AlignHeaderFooterWithMargins = false;
        sheet.PageSetup.ScaleHeaderFooterWithDocument = false;
        // Writing variant text turns the variant on without copying the running text into it.
        sheet.PageSetup.FirstPageHeader.Center.Text = "First title";
        sheet.PageSetup.EvenPageFooter.Center.Text = "Even footer";
        sheet.PageSetup.Header.Left.Text = string.Empty;
        sheet.PageSetup.Header.Center.Text = "Updated title";
        sheet.PageSetup.Footer.Left.Text = "Reviewed";
        sheet.PageSetup.DifferentFirstPage = true;
        sheet.PageSetup.DifferentOddEven = true;
        using var saved = new MemoryStream();
        document.Save(saved); saved.Position = 0;
        using (var reopened = ExcelDocument.Load(saved)) {
            var info = reopened.Sheets[0].PageSetup;
            Assert.Equal(string.Empty, info.Header.Left.Text);
            Assert.Equal("Updated title", info.Header.Center.Text);
            Assert.Equal("&P", info.Header.Right.Text);
            Assert.Equal("Reviewed", info.Footer.Left.Text);
            Assert.Equal("Page &P of &N", info.Footer.Right.Text);
            Assert.Equal("First title", info.FirstPageHeader.Center.Text);
            Assert.Empty(info.FirstPageFooter.Left.Text); Assert.Empty(info.FirstPageFooter.Center.Text); Assert.Empty(info.FirstPageFooter.Right.Text);
            Assert.Empty(info.EvenPageHeader.Left.Text); Assert.Empty(info.EvenPageHeader.Center.Text); Assert.Empty(info.EvenPageHeader.Right.Text);
            Assert.Equal("Even footer", info.EvenPageFooter.Center.Text);
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
        sheet.PageSetup.DifferentFirstPage = false;
        sheet.PageSetup.AlignHeaderFooterWithMargins = true;
        using var disabled = new MemoryStream();
        document.Save(disabled); disabled.Position = 0;
        using var final = ExcelDocument.Load(disabled);
        var disabledInfo = final.Sheets[0].PageSetup;
        Assert.False(disabledInfo.DifferentFirstPage);
        Assert.True(disabledInfo.DifferentOddEven);
        Assert.Equal("First title", disabledInfo.FirstPageHeader.Center.Text);
        Assert.Empty(disabledInfo.FirstPageFooter.Left.Text); Assert.Empty(disabledInfo.FirstPageFooter.Center.Text); Assert.Empty(disabledInfo.FirstPageFooter.Right.Text);
        Assert.Empty(disabledInfo.EvenPageHeader.Left.Text); Assert.Empty(disabledInfo.EvenPageHeader.Center.Text); Assert.Empty(disabledInfo.EvenPageHeader.Right.Text);
        Assert.Equal("Even footer", disabledInfo.EvenPageFooter.Center.Text);
        Assert.Equal("Page &P of &N", disabledInfo.Footer.Right.Text);
        sheet.PageSetup.DifferentFirstPage = true;
        using var enabled = new MemoryStream();
        document.Save(enabled); enabled.Position = 0;
        using var enabledDocument = ExcelDocument.Load(enabled);
        var enabledInfo = enabledDocument.Sheets[0].PageSetup;
        Assert.True(enabledInfo.DifferentFirstPage);
        Assert.Equal("First title", enabledInfo.FirstPageHeader.Center.Text);
        Assert.Equal("Reviewed", enabledInfo.FirstPageFooter.Left.Text);
        Assert.Equal("Page &P of &N", enabledInfo.FirstPageFooter.Right.Text);
        Assert.Empty(enabledInfo.EvenPageHeader.Center.Text);
    }
}
