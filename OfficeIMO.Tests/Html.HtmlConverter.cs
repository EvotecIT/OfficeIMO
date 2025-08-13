using OfficeIMO.Word.Html;
using OfficeIMO.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    private static void RemoveCustomStyle(string styleId) {
        var field = typeof(WordParagraphStyle).GetField("_customStyles", BindingFlags.NonPublic | BindingFlags.Static);
        var dict = (IDictionary<string, Style>)field!.GetValue(null);
        dict.Remove(styleId);
    }
    [Fact(Skip = "TODO: Implement HTML to Word conversion - currently only stub implementation")]
    public void Test_Html_RoundTrip() {
        string html = "<p>Hello <b>world</b> and <i>universe</i>.</p>";
        
        var doc = html.LoadFromHtml(new HtmlToWordOptions { FontFamily = "Calibri" });
        string roundTrip = doc.ToHtml(new WordToHtmlOptions { IncludeFontStyles = true });

        Assert.Contains("<b>", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("</b>", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("world", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<i>", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("</i>", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("universe", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains($"font-family:{FontResolver.Resolve("Calibri")}", roundTrip, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Test_Html_Headings_RoundTrip() {
        string html = "<h1>Heading 1</h1><h2>Heading 2</h2><h3>Heading 3</h3><h4>Heading 4</h4><h5>Heading 5</h5><h6>Heading 6</h6>";
        
        var doc = html.LoadFromHtml(new HtmlToWordOptions { FontFamily = "Calibri" });
        string roundTrip = doc.ToHtml(new WordToHtmlOptions { IncludeFontStyles = true });

        for (int i = 1; i <= 6; i++) {
            string tag = $"h{i}";
            Assert.Contains("<" + tag + ">", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains($"Heading {i}", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("</" + tag + ">", roundTrip, StringComparison.OrdinalIgnoreCase);
        }
    }

    [Fact]
    public void Test_Html_Lists_RoundTrip() {
        string html = "<ul><li>Item 1<ul><li>Sub 1</li><li>Sub 2</li></ul></li><li>Item 2</li></ul><ol><li>First</li><li>Second</li></ol>";
        
        var doc = html.LoadFromHtml(new HtmlToWordOptions());
        string roundTrip = doc.ToHtml(new WordToHtmlOptions { IncludeListStyles = true });

        Assert.Contains("<ul", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<ol", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Sub 1", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Second", roundTrip, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Test_Html_Table_RoundTrip() {
        string html = "<table><tr><td>A</td><td>B</td></tr><tr><td>C</td><td>D</td></tr></table>";
        
        var doc = html.LoadFromHtml(new HtmlToWordOptions());
        string roundTrip = doc.ToHtml(new WordToHtmlOptions());

        Assert.Contains("<table>", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("A", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("D", roundTrip, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Test_Html_NestedTable_RoundTrip() {
        string html = "<table><tr><td>Outer</td><td><table><tr><td>Inner</td></tr></table></td></tr></table>";

        var doc = html.LoadFromHtml(new HtmlToWordOptions());

        Assert.Single(doc.Sections[0].Tables);
        var outer = doc.Sections[0].Tables[0];
        Assert.Equal(2, outer.Rows[0].Cells.Count);
        var innerCell = outer.Rows[0].Cells[1];
        Assert.True(innerCell.HasNestedTables);
        var inner = innerCell.NestedTables[0];
        Assert.Equal(1, inner.Rows.Count);
        Assert.Equal(1, inner.Rows[0].Cells.Count);

        string roundTrip = doc.ToHtml(new WordToHtmlOptions());
        int tableCount = System.Text.RegularExpressions.Regex.Matches(roundTrip, "<table", System.Text.RegularExpressions.RegexOptions.IgnoreCase).Count;
        Assert.True(tableCount >= 2);
        Assert.Contains("Outer", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Inner", roundTrip, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Test_Html_Image_Base64_RoundTrip() {
        string assetPath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png");
        byte[] imageBytes = File.ReadAllBytes(assetPath);
        string base64 = Convert.ToBase64String(imageBytes);
        string html = $"<p><img src=\"data:image/png;base64,{base64}\" /></p>";
        
        var doc = html.LoadFromHtml(new HtmlToWordOptions());
        string roundTrip = doc.ToHtml(new WordToHtmlOptions());

        Assert.Contains("<img", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("data:image/png;base64", roundTrip, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Test_Html_Image_File_RoundTrip() {
        string assetPath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png");
        string uri = new Uri(assetPath).AbsoluteUri;
        string html = $"<p><img src=\"{uri}\" /></p>";
        
        var doc = html.LoadFromHtml(new HtmlToWordOptions());
        string roundTrip = doc.ToHtml(new WordToHtmlOptions());

        Assert.Contains("<img", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("data:image/png;base64", roundTrip, StringComparison.OrdinalIgnoreCase);
    }

    [Fact(Skip = "TODO: Implement font family mapping and CSS font-family support")]
    public void Test_Html_FontResolver() {
        string html = "<p>Hello</p>";
        
        var doc = html.LoadFromHtml(new HtmlToWordOptions { FontFamily = "monospace" });
        string roundTrip = doc.ToHtml(new WordToHtmlOptions { IncludeFontStyles = true });
        
        Assert.Contains($"font-family:{FontResolver.Resolve("monospace")}", roundTrip, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Test_Html_Urls_CreateHyperlinks() {
        string html = "<p>Visit http://example.com</p>";
        using MemoryStream ms = new MemoryStream();
        
        var doc = html.LoadFromHtml(new HtmlToWordOptions());
        doc.Save(ms);

        ms.Position = 0;
        using WordprocessingDocument docx = WordprocessingDocument.Open(ms, false);
        var hyperlink = docx.MainDocumentPart!.Document.Body!.Descendants<Hyperlink>().FirstOrDefault();
        Assert.NotNull(hyperlink);
        var rel = docx.MainDocumentPart.HyperlinkRelationships.First();
        Assert.StartsWith("http://example.com", rel.Uri.ToString());
    }

    [Fact]
    public void Test_Html_InlineStyles_ParagraphStyle() {
        string html = "<p style=\"font-weight:bold;font-size:32px\">Styled</p>";
        using MemoryStream ms = new MemoryStream();
        
        var doc = html.LoadFromHtml(new HtmlToWordOptions());
        doc.Save(ms);

        ms.Position = 0;
        using WordprocessingDocument docx = WordprocessingDocument.Open(ms, false);
        Paragraph p = docx.MainDocumentPart!.Document.Body!.Elements<Paragraph>().First();
        string styleId = p.ParagraphProperties?.ParagraphStyleId?.Val;
        Assert.Equal(WordParagraphStyles.Heading1.ToString(), styleId);
    }

    [Fact(Skip = "Top bookmark shifts paragraph order")]
    public void Test_Html_Headings() {
        string html = "<h1>Heading 1</h1><h2>Heading 2</h2>";

        var doc = html.LoadFromHtml(new HtmlToWordOptions());

        Assert.Equal(WordParagraphStyles.Heading1, doc.Paragraphs[1].Style);
        Assert.Equal("Heading 1", doc.Paragraphs[1].Text);
        Assert.Equal(WordParagraphStyles.Heading2, doc.Paragraphs[2].Style);
    }

    [Fact]
    public void Test_Html_Blockquote_RoundTrip() {
        string html = "<blockquote>Quoted text</blockquote>";

        var doc = html.LoadFromHtml(new HtmlToWordOptions());
        Assert.Equal("Quoted text", doc.Paragraphs[0].Text);
        Assert.True(doc.Paragraphs[0].IndentationBefore > 0);

        string roundTrip = doc.ToHtml();
        Assert.Contains("<blockquote>", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Quoted text", roundTrip, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Test_Html_Blockquote_WithoutQuoteStyle() {
        RemoveCustomStyle("Quote");
        string html = "<blockquote>Quoted text</blockquote>";

        var doc = html.LoadFromHtml(new HtmlToWordOptions());

        Assert.False(doc.StyleExists("Quote"));
        Assert.Equal("Quoted text", doc.Paragraphs[0].Text);
        Assert.True(doc.Paragraphs[0].IndentationBefore > 0);
        Assert.Null(doc.Paragraphs[0].Style);
    }

    [Fact]
    public void Test_Html_Blockquote_WithQuoteStyle() {
        RemoveCustomStyle("Quote");
        var quote = WordParagraphStyle.CreateFontStyle("Quote", "Arial");
        WordParagraphStyle.RegisterCustomStyle("Quote", quote);

        string html = "<blockquote>Quoted text</blockquote>";
        var doc = html.LoadFromHtml(new HtmlToWordOptions());

        Assert.True(doc.StyleExists("Quote"));
        Assert.Equal("Quoted text", doc.Paragraphs[0].Text);
        Assert.Equal(WordParagraphStyles.Custom, doc.Paragraphs[0].Style);

        RemoveCustomStyle("Quote");
    }

    [Fact]
    public void Test_Html_Q_RoundTrip() {
        string html = "<p>Before <q>quoted</q> after</p>";

        var doc = html.LoadFromHtml(new HtmlToWordOptions());

        var text = string.Concat(doc.Paragraphs[0].GetRuns().Select(r => r.Text));
        Assert.Equal($"Before \u201Cquoted\u201D after", text);

        string roundTrip = doc.ToHtml();
        Assert.Contains("<q>quoted</q>", roundTrip, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Test_Html_Q_CustomCharacters() {
        var options = new HtmlToWordOptions { QuotePrefix = "«", QuoteSuffix = "»" };
        string html = "<p>Before <q>quoted</q> after</p>";

        var doc = html.LoadFromHtml(options);

        var text = string.Concat(doc.Paragraphs[0].GetRuns().Select(r => r.Text));
        Assert.Equal("Before «quoted» after", text);

        string roundTrip = doc.ToHtml();
        Assert.Contains("<q>quoted</q>", roundTrip, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Test_Html_Lists_Structure() {
        string html = "<ul><li>Item 1<ul><li>Sub 1</li></ul></li><li>Item 2</li></ul>";

        var doc = html.LoadFromHtml(new HtmlToWordOptions());

        Assert.True(doc.Lists.Count > 0);
    }

    [Fact]
    public void Test_Html_OrderedList_StartAndType() {
        string html = "<ol start=\"5\" type=\"a\"><li>First</li><li>Second</li></ol>";

        var doc = html.LoadFromHtml(new HtmlToWordOptions());

        Assert.Equal(5, doc.Lists[0].Numbering.Levels[0].StartNumberingValue);
        Assert.Equal(NumberFormatValues.LowerLetter, doc.Lists[0].Numbering.Levels[0]._level.NumberingFormat.Val.Value);
    }

    [Fact]
    public void Test_Html_UnorderedList_Type() {
        string html = "<ul type=\"circle\"><li>A</li><li>B</li></ul>";

        var doc = html.LoadFromHtml(new HtmlToWordOptions());

        Assert.Equal("o", doc.Lists[0].Numbering.Levels[0]._level.LevelText.Val);
    }

    [Fact]
    public void Test_Html_Table_Structure() {
        string html = "<table><tr><td>A</td><td>B</td></tr><tr><td>C</td><td>D</td></tr></table>";

        var doc = html.LoadFromHtml(new HtmlToWordOptions());

        using MemoryStream ms = new MemoryStream();
        doc.Save(ms);
        ms.Position = 0;
        using WordprocessingDocument docx = WordprocessingDocument.Open(ms, false);
        var cells = docx.MainDocumentPart!.Document.Body!.Descendants<TableCell>().ToArray();
        Assert.Contains("A", cells[0].InnerText);
        Assert.Contains("D", cells[3].InnerText);
    }

    [Fact]
    public void Test_Html_Image_Base64_Conversion() {
        string assetPath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png");
        byte[] imageBytes = File.ReadAllBytes(assetPath);
        string base64 = Convert.ToBase64String(imageBytes);
        string html = $"<p><img src=\"data:image/png;base64,{base64}\" /></p>";

        var doc = html.LoadFromHtml(new HtmlToWordOptions());

        Assert.Single(doc.Images);
    }

    [Fact]
    public void Test_Html_Image_File_Conversion() {
        string assetPath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png");
        string uri = new Uri(assetPath).AbsoluteUri;
        string html = $"<p><img src=\"{uri}\" /></p>";

        var doc = html.LoadFromHtml(new HtmlToWordOptions());

        Assert.Single(doc.Images);
    }

    [Fact]
    public void Test_Html_ImageAlt_Preserved() {
        string assetPath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png");
        byte[] imageBytes = File.ReadAllBytes(assetPath);
        string base64 = Convert.ToBase64String(imageBytes);
        string html = $"<p><img src=\"data:image/png;base64,{base64}\" alt=\"Company logo\" /></p>";

        var doc = html.LoadFromHtml(new HtmlToWordOptions());

        Assert.Single(doc.Images);
        Assert.Equal("Company logo", doc.Images[0].Description);

        string roundTrip = doc.ToHtml();
        Assert.Contains("alt=\"Company logo\"", roundTrip, StringComparison.OrdinalIgnoreCase);
    }

    [Fact(Skip = "Top bookmark affects paragraph count")]
    public void Test_Html_HorizontalRule_RoundTrip() {
        string html = "<p>Before</p><hr><p>After</p>";

        var doc = html.LoadFromHtml(new HtmlToWordOptions());

        Assert.Equal(4, doc.Paragraphs.Count);
        
        string roundTrip = doc.ToHtml();
        Assert.Contains("<hr", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Before", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("After", roundTrip, StringComparison.OrdinalIgnoreCase);
    }
}
