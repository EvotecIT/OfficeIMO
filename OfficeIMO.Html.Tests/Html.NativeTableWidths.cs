using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;
using System.Xml.Linq;

namespace OfficeIMO.Tests;

public class HtmlNativeTableWidths {
    private const string FiveColumns = "<tr><td>First</td><td>Second</td><td>Third</td><td>Fourth</td><td>Fifth column is visible</td></tr>";

    [Fact]
    public void HtmlToWord_DefaultFiveColumnTableFitsSectionAfterSaveAndReopen() {
        using WordDocument word = HtmlConversionDocument.Parse("<table>" + FiveColumns + "</table>").ToWordDocument();
        using MemoryStream saved = word.ToStream();
        using WordDocument reopened = WordDocument.Load(saved);
        WordTable table = Assert.Single(reopened.Tables);
        WordSection section = reopened.Sections[0];
        int textWidth = (int)section.PageSettings.Width!.Value - (int)section.Margins.Left - (int)section.Margins.Right;
        Assert.InRange(table.GridColumnWidth.Sum(), 1, textWidth);
        Assert.Equal(5, table.Rows[0].Cells.Count);
        Assert.Contains("Fifth column is visible", table.Rows[0].Cells[4].Paragraphs.Select(p => p.Text));
    }

    [Theory]
    [InlineData("style='width:600px'", "", 9000)]
    [InlineData("", "<colgroup><col span='5' style='width:100px'></colgroup>", 7500)]
    public void HtmlToWord_AuthoredTableAndColumnWidthsArePreserved(string attributes, string columns, int width) {
        using WordDocument word = HtmlConversionDocument.Parse("<table " + attributes + ">" + columns + FiveColumns + "</table>").ToWordDocument();
        WordTable table = Assert.Single(word.Tables);
        Assert.Equal(width, table.GridColumnWidth.Sum());
    }

    [Fact]
    public void HtmlToWord_DefaultNestedTableFitsItsContainingCell() {
        using WordDocument word = HtmlConversionDocument.Parse("<table><tr><td>Outer</td><td><table>" + FiveColumns + "</table></td></tr></table>").ToWordDocument();
        WordTable outer = Assert.Single(word.Tables);
        WordTableCell containingCell = outer.Rows[0].Cells[1];
        WordTable nested = Assert.Single(containingCell.NestedTables);
        Assert.InRange(nested.GridColumnWidth.Sum(), 1, containingCell.Width!.Value);
    }
    [Theory]
    [InlineData("", "")]
    [InlineData("style='width:100%'", "")]
    [InlineData("", "<colgroup><col span='5' style='width:20%'></colgroup>")]
    public void HtmlToWord_ListItemNestedTableFitsItsActualHostCell(string attributes, string columns) {
        using WordDocument word = HtmlConversionDocument.Parse("<table><tr><td>Outer</td><td><ul><li><table " + attributes + ">" + columns + FiveColumns + "</table></li></ul></td></tr></table>").ToWordDocument();
        WordTable outer = Assert.Single(word.Tables);
        WordTableCell containingCell = outer.Rows[0].Cells[1];
        using MemoryStream saved = word.ToStream();
        using var archive = new System.IO.Compression.ZipArchive(saved, System.IO.Compression.ZipArchiveMode.Read);
        using var xml = archive.GetEntry("word/document.xml")!.Open();
        var document = System.Xml.Linq.XDocument.Load(xml);
        System.Xml.Linq.XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        Assert.Equal(new[] { 2400, 2400 }, outer.GridColumnWidth);
        var nested = Assert.Single(document.Descendants(w + "sdtContent").Elements(w + "tbl"));
        int width = nested.Element(w + "tblGrid")!.Elements(w + "gridCol").Sum(col => (int)col.Attribute(w + "w")!);
        Assert.InRange(width, 1, containingCell.Width!.Value - 216);
    }

    [Theory]
    [InlineData("body", "")]
    [InlineData("header", "")]
    [InlineData("footer", "")]
    [InlineData("body", "style='width:100%'")]
    [InlineData("header", "style='width:100%'")]
    [InlineData("footer", "style='width:100%'")]
    public void HtmlToWord_WrappedTableUsesTheTargetSectionWidth(string scope, string attributes) {
        using WordDocument word = WordDocument.Create();
        word.AddParagraph("First section");
        WordSection target = word.AddSection();
        target.PageSettings.Width = 7000;
        target.Margins.Left = 1000;
        target.Margins.Right = 1000;
        HtmlConversionDocument html = HtmlConversionDocument.Parse("<ul><li><table " + attributes + ">" + FiveColumns + "</table></li></ul>");
        if (scope == "header") word.AddHtmlToHeader(html);
        else if (scope == "footer") word.AddHtmlToFooter(html);
        else word.AddHtmlToBody(html);
        using MemoryStream saved = word.ToStream();
        using var archive = new System.IO.Compression.ZipArchive(saved, System.IO.Compression.ZipArchiveMode.Read);
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        int[] widths = archive.Entries.Where(entry => entry.FullName.StartsWith("word/", StringComparison.Ordinal) && entry.FullName.EndsWith(".xml", StringComparison.Ordinal))
            .SelectMany(entry => { using var xml = entry.Open(); return XDocument.Load(xml).Descendants(w + "tblGrid").Select(grid => grid.Elements(w + "gridCol").Sum(col => (int)col.Attribute(w + "w")!)).ToArray(); }).ToArray();
        Assert.Equal(5000, Assert.Single(widths));
    }

    [Theory]
    [InlineData(false, 100, 5000, false)]
    [InlineData(true, 100, 5000, false)]
    [InlineData(false, 50, 2500, false)]
    [InlineData(true, 50, 2500, false)]
    [InlineData(false, 100, 5000, true)]
    [InlineData(true, 100, 5000, true)]
    [InlineData(false, 50, 2500, true)]
    [InlineData(true, 50, 2500, true)]
    public void HtmlToWord_SharedStoryTableUsesTheSelectedLayoutSection(bool footer, int percent, int expectedWidth, bool inherited) {
        using WordDocument word = WordDocument.Create();
        WordSection first = word.Sections[0];
        first.PageSettings.Width = 12000;
        first.Margins.Left = 1500;
        first.Margins.Right = 1500;
        if (footer) first.GetOrCreateFooter(WordHeaderFooterType.Default).AddParagraph("Shared footer");
        else first.GetOrCreateHeader(WordHeaderFooterType.Default).AddParagraph("Shared header");
        WordSection last = word.AddSection();
        last.PageSettings.Width = 7000;
        last.Margins.Left = 1000;
        last.Margins.Right = 1000;
        if (footer) {
            var reference = first._sectionProperties.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.FooterReference>()!;
            last._sectionProperties.RemoveAllChildren<DocumentFormat.OpenXml.Wordprocessing.FooterReference>();
            if (!inherited) last._sectionProperties.PrependChild((DocumentFormat.OpenXml.Wordprocessing.FooterReference)reference.CloneNode(true));
            last.Footer.Default = first.Footer.Default;
        } else {
            var reference = first._sectionProperties.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.HeaderReference>()!;
            last._sectionProperties.RemoveAllChildren<DocumentFormat.OpenXml.Wordprocessing.HeaderReference>();
            if (!inherited) last._sectionProperties.PrependChild((DocumentFormat.OpenXml.Wordprocessing.HeaderReference)reference.CloneNode(true));
            last.Header.Default = first.Header.Default;
        }
        HtmlConversionDocument html = HtmlConversionDocument.Parse("<ul><li><table style='width:" + percent + "%'>" + FiveColumns + "</table></li></ul>");
        if (footer) word.AddHtmlToFooter(html); else word.AddHtmlToHeader(html);
        using MemoryStream saved = word.ToStream();
        using WordDocument reopened = WordDocument.Load(saved);
        using MemoryStream resaved = reopened.ToStream();
        using var archive = new System.IO.Compression.ZipArchive(resaved, System.IO.Compression.ZipArchiveMode.Read);
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        int[] widths = archive.Entries.Where(entry => entry.FullName.StartsWith(footer ? "word/footer" : "word/header", StringComparison.Ordinal) && entry.FullName.EndsWith(".xml", StringComparison.Ordinal))
            .SelectMany(entry => { using var xml = entry.Open(); return XDocument.Load(xml).Descendants(w + "tblGrid").Select(grid => grid.Elements(w + "gridCol").Sum(col => (int)col.Attribute(w + "w")!)).ToArray(); }).ToArray();
        Assert.Equal(expectedWidth, Assert.Single(widths));
    }

    [Fact]
    public void HtmlToWord_MovedPercentageTableUsesItsCurrentBodySection() {
        using WordDocument word = WordDocument.Create();
        WordSection first = word.Sections[0];
        first.PageSettings.Width = 7000;
        first.Margins.Left = 1000;
        first.Margins.Right = 1000;
        word.AddHtmlToBody(HtmlConversionDocument.Parse("<table style='width:100%'>" + FiveColumns + "</table>"));
        WordTable table = Assert.Single(word.Tables);
        Assert.Equal(5000, table.GridColumnWidth.Sum());
        WordSection last = word.AddSection();
        last.PageSettings.Width = 12000;
        last.Margins.Left = 1500;
        last.Margins.Right = 1500;
        WordParagraph anchor = word.AddParagraph("Second section anchor");
        table.Remove();
        word.InsertTableAfter(anchor, table);
        using MemoryStream saved = word.ToStream();
        using WordDocument reopened = WordDocument.Load(saved);
        Assert.Equal(9000, Assert.Single(reopened.Tables).GridColumnWidth.Sum());
    }

}
