using System.Linq;
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
        using WordDocument word = HtmlConversionDocument.Parse("<table>" + FiveColumns + "</table>").ToWord();
        using MemoryStream saved = new System.IO.MemoryStream(word.ToBytes());
        using WordDocument reopened = WordDocument.Load(saved);
        WordTable table = Assert.Single(reopened.Tables);
        WordSection section = reopened.Sections[0];
        int textWidth = (int)checked((uint?)(section.PageSettings.Width?.ToTwips()))!.Value - (int)checked((uint)(section.Margins.Left.GetValueOrDefault().ToTwips())) - (int)checked((uint)(section.Margins.Right.GetValueOrDefault().ToTwips()));
        Assert.InRange(table.GridColumnWidth.Select(width => checked((int)width.ToTwips())).ToList().Sum(), 1, textWidth);
        Assert.Equal(5, table.Rows[0].Cells.Count);
        Assert.Contains("Fifth column is visible", table.Rows[0].Cells[4].Paragraphs.Select(p => p.Text));
    }

    [Theory]
    [InlineData("style='width:600px'", "", 9000)]
    [InlineData("", "<colgroup><col span='5' style='width:100px'></colgroup>", 7500)]
    public void HtmlToWord_AuthoredTableAndColumnWidthsArePreserved(string attributes, string columns, int width) {
        using WordDocument word = HtmlConversionDocument.Parse("<table " + attributes + ">" + columns + FiveColumns + "</table>").ToWord();
        WordTable table = Assert.Single(word.Tables);
        Assert.Equal(width, table.GridColumnWidth.Select(width => checked((int)width.ToTwips())).ToList().Sum());
    }

    [Fact]
    public void HtmlToWord_DefaultNestedTableFitsItsContainingCell() {
        using WordDocument word = HtmlConversionDocument.Parse("<table><tr><td>Outer</td><td><table>" + FiveColumns + "</table></td></tr></table>").ToWord();
        WordTable outer = Assert.Single(word.Tables);
        WordTableCell containingCell = outer.Rows[0].Cells[1];
        WordTable nested = Assert.Single(containingCell.NestedTables);
        Assert.InRange(nested.GridColumnWidth.Select(width => checked((int)width.ToTwips())).ToList().Sum(), 1, containingCell.Width!.Value.ToNativeValue());
    }
    [Theory]
    [InlineData("", "")]
    [InlineData("style='width:100%'", "")]
    [InlineData("", "<colgroup><col span='5' style='width:20%'></colgroup>")]
    public void HtmlToWord_ListItemNestedTableFitsItsActualHostCell(string attributes, string columns) {
        using WordDocument word = HtmlConversionDocument.Parse("<table><tr><td>Outer</td><td><ul><li><table " + attributes + ">" + columns + FiveColumns + "</table></li></ul></td></tr></table>").ToWord();
        WordTable outer = Assert.Single(word.Tables);
        WordTableCell containingCell = outer.Rows[0].Cells[1];
        using MemoryStream saved = new System.IO.MemoryStream(word.ToBytes());
        using var archive = new System.IO.Compression.ZipArchive(saved, System.IO.Compression.ZipArchiveMode.Read);
        using var xml = archive.GetEntry("word/document.xml")!.Open();
        var document = System.Xml.Linq.XDocument.Load(xml);
        System.Xml.Linq.XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        Assert.Equal(new[] { 2400, 2400 }, outer.GridColumnWidth.Select(width => checked((int)width.ToTwips())).ToList());
        var nested = Assert.Single(document.Descendants(w + "sdtContent").Elements(w + "tbl"));
        int width = nested.Element(w + "tblGrid")!.Elements(w + "gridCol").Sum(col => (int)col.Attribute(w + "w")!);
        Assert.InRange(width, 1, containingCell.Width!.Value.ToNativeValue() - 216);
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
        target.PageSettings.Width = global::OfficeIMO.OfficeLength.Points(350D);
        target.Margins.Left = global::OfficeIMO.OfficeLength.Points(50D);
        target.Margins.Right = global::OfficeIMO.OfficeLength.Points(50D);
        HtmlConversionDocument html = HtmlConversionDocument.Parse("<ul><li><table " + attributes + ">" + FiveColumns + "</table></li></ul>");
        if (scope == "header") target.GetOrCreateHeader(WordHeaderFooterType.Default).AddHtml(html);
        else if (scope == "footer") target.GetOrCreateFooter(WordHeaderFooterType.Default).AddHtml(html);
        else word.AddHtml(html);
        using MemoryStream saved = new System.IO.MemoryStream(word.ToBytes());
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
    public async System.Threading.Tasks.Task HtmlToWord_SharedStoryImageAndTableUseTheSelectedLayoutSection(bool footer, int percent, int expectedWidth, bool inherited) {
        using WordDocument word = WordDocument.Create();
        WordSection first = word.Sections[0];
        first.PageSettings.Width = global::OfficeIMO.OfficeLength.Points(600D);
        first.Margins.Left = global::OfficeIMO.OfficeLength.Points(75D);
        first.Margins.Right = global::OfficeIMO.OfficeLength.Points(75D);
        if (footer) first.GetOrCreateFooter(WordHeaderFooterType.Default).AddParagraph("Shared footer");
        else first.GetOrCreateHeader(WordHeaderFooterType.Default).AddParagraph("Shared header");
        WordSection last = word.AddSection();
        last.PageSettings.Width = global::OfficeIMO.OfficeLength.Points(350D);
        last.Margins.Left = global::OfficeIMO.OfficeLength.Points(50D);
        last.Margins.Right = global::OfficeIMO.OfficeLength.Points(50D);
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
        const string pixel = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP4/w8AAv8B/h10yjMAAAAASUVORK5CYII=";
        HtmlConversionDocument html = HtmlConversionDocument.Parse("<ul><li><table style='width:" + percent + "%'>" + FiveColumns +
            "</table></li></ul><p><img src='data:image/png;base64," + pixel + "' width='1200' height='1200' style='max-width:100%' alt='Shared story image'></p>");
        WordHeaderFooter story = footer ? first.Footer.Default! : first.Header.Default!;
        if (percent == 100 && !inherited) {
            using WordDocument foreign = WordDocument.Create();
            string before = (story._header as DocumentFormat.OpenXml.OpenXmlElement ?? story._footer!).OuterXml;
            if (footer) {
                Assert.Throws<ArgumentException>(() => first.Footer.Default!.AddHtml(html, section: foreign.Sections[0]));
                WordFooter unrelatedStory = last.GetOrCreateFooter(WordHeaderFooterType.First);
                string unrelatedBefore = unrelatedStory._footer!.OuterXml;
                await Assert.ThrowsAsync<ArgumentException>(() => unrelatedStory.AddHtmlAsync(html, section: first));
                Assert.Equal(unrelatedBefore, unrelatedStory._footer.OuterXml);
            } else {
                await Assert.ThrowsAsync<ArgumentException>(() => first.Header.Default!.AddHtmlAsync(html, section: foreign.Sections[0]));
                WordHeader unrelatedStory = last.GetOrCreateHeader(WordHeaderFooterType.First);
                string unrelatedBefore = unrelatedStory._header!.OuterXml;
                Assert.Throws<ArgumentException>(() => unrelatedStory.AddHtml(html, section: first));
                Assert.Equal(unrelatedBefore, unrelatedStory._header.OuterXml);
            }
            Assert.Equal(before, (story._header as DocumentFormat.OpenXml.OpenXmlElement ?? story._footer!).OuterXml);
        }
        var created = footer
            ? await first.Footer.Default!.AddHtmlAsync(html, section: last)
            : first.Header.Default!.AddHtml(html, section: last);
        Assert.NotEmpty(created);
        Assert.DoesNotContain(created.OfType<WordParagraph>(), paragraph => paragraph.Text.Contains("Shared footer", StringComparison.Ordinal) || paragraph.Text.Contains("Shared header", StringComparison.Ordinal));
        WordParagraph importedImage = Assert.Single(created.OfType<WordParagraph>(), paragraph => paragraph.Image != null);
        Assert.Equal(250D, importedImage.Image!.Width!.Value.ToPoints(), 3);
        importedImage.AddText("Imported image");
        if (footer) first.Footer.Default!.AddHtml(html);
        else await first.Header.Default!.AddHtmlAsync(html);
        using MemoryStream saved = new System.IO.MemoryStream(word.ToBytes());
        using WordDocument reopened = WordDocument.Load(saved);
        using MemoryStream resaved = new System.IO.MemoryStream(reopened.ToBytes());
        using var archive = new System.IO.Compression.ZipArchive(resaved, System.IO.Compression.ZipArchiveMode.Read);
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        int[] widths = archive.Entries.Where(entry => entry.FullName.StartsWith(footer ? "word/footer" : "word/header", StringComparison.Ordinal) && entry.FullName.EndsWith(".xml", StringComparison.Ordinal))
            .SelectMany(entry => { using var xml = entry.Open(); return XDocument.Load(xml).Descendants(w + "tblGrid").Select(grid => grid.Elements(w + "gridCol").Sum(col => (int)col.Attribute(w + "w")!)).ToArray(); }).ToArray();
        Assert.Equal(new[] { expectedWidth, 9000 * percent / 100 }, widths);
        XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        long[] imageWidths = archive.Entries.Where(entry => entry.FullName.StartsWith(footer ? "word/footer" : "word/header", StringComparison.Ordinal) && entry.FullName.EndsWith(".xml", StringComparison.Ordinal))
            .SelectMany(entry => { using var xml = entry.Open(); return XDocument.Load(xml).Descendants(wp + "extent").Select(extent => (long)extent.Attribute("cx")!).ToArray(); }).ToArray();
        Assert.Equal(new[] { 5000L * 635, 9000L * 635 }, imageWidths);
        Assert.Contains("Imported image", footer ? reopened.Sections[0].Footer.Default!.Paragraphs.Select(paragraph => paragraph.Text) : reopened.Sections[0].Header.Default!.Paragraphs.Select(paragraph => paragraph.Text));
    }

    [Fact]
    public void HtmlToWord_MovedPercentageTableUsesItsCurrentBodySection() {
        using WordDocument word = WordDocument.Create();
        WordSection first = word.Sections[0];
        first.PageSettings.Width = global::OfficeIMO.OfficeLength.Points(350D);
        first.Margins.Left = global::OfficeIMO.OfficeLength.Points(50D);
        first.Margins.Right = global::OfficeIMO.OfficeLength.Points(50D);
        word.AddHtml(HtmlConversionDocument.Parse("<table style='width:100%'>" + FiveColumns + "</table>"));
        WordTable table = Assert.Single(word.Tables);
        Assert.Equal(5000, table.GridColumnWidth.Select(width => checked((int)width.ToTwips())).ToList().Sum());
        WordSection last = word.AddSection();
        last.PageSettings.Width = global::OfficeIMO.OfficeLength.Points(600D);
        last.Margins.Left = global::OfficeIMO.OfficeLength.Points(75D);
        last.Margins.Right = global::OfficeIMO.OfficeLength.Points(75D);
        WordParagraph anchor = word.AddParagraph("Second section anchor");
        table.Remove();
        word.InsertTableAfter(anchor, table);
        using MemoryStream saved = new System.IO.MemoryStream(word.ToBytes());
        using WordDocument reopened = WordDocument.Load(saved);
        Assert.Equal(9000, Assert.Single(reopened.Tables).GridColumnWidth.Select(width => checked((int)width.ToTwips())).ToList().Sum());
    }

}
