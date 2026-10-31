using System.Xml.Linq;
using OfficeIMO.Excel;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class OfficeQuickStartAuthoringTests {
    [Theory]
    [InlineData(1, WordParagraphStyles.Heading1)]
    [InlineData(9, WordParagraphStyles.Heading9)]
    public void AddedHeadingRetainsItsStyleAfterReload(int level, WordParagraphStyles style) {
        using var bytes = new MemoryStream();
        using (var document = WordDocument.Create()) {
            var heading = document.AddHeading("Quarterly report", level);
            Assert.Equal(style, heading.Style);
            Assert.Throws<ArgumentOutOfRangeException>(() => document.AddHeading("Invalid", 10));
            Assert.Single(document.Paragraphs);
            document.Save(bytes);
        }
        bytes.Position = 0;
        using var loaded = WordDocument.Load(bytes);
        var paragraph = Assert.Single(loaded.Paragraphs);
        Assert.Equal("Quarterly report", paragraph.Text);
        Assert.Equal(style, paragraph.Style);
    }

    [Fact]
    public void ExcelCreateThemePersistsItsPaletteAndFonts() {
        OfficeTheme theme = OfficeTheme.Modern.WithName("Review theme");
        using var bytes = new MemoryStream();
        using (var book = ExcelDocument.Create(new ExcelCreateOptions { Theme = theme })) {
            book.AddSheet("Orders");
            book.Save(bytes);
        }
        bytes.Position = 0;
        using var loaded = ExcelDocument.Load(bytes);
        var info = loaded.GetWorkbookTheme(includeXml: true);
        Assert.Equal(theme.Name, info.Name);
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var xml = XDocument.Parse(info.Xml!);
        Assert.Equal(theme.Palette.Accent1.ToRgbHex(), xml.Descendants(a + "accent1").Single().Element(a + "srgbClr")!.Attribute("val")!.Value);
        Assert.Equal(theme.Fonts.Body.Family, xml.Descendants(a + "minorFont").Single().Element(a + "latin")!.Attribute("typeface")!.Value);
    }
}
