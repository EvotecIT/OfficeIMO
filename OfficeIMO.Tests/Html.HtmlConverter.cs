using OfficeIMO.Word.Html;
using OfficeIMO.Word;
using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
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

    [Fact(Skip = "TODO: Implement heading conversion (h1-h6 -> WordParagraphStyles.Heading1-6)")]
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

    [Fact(Skip = "TODO: Implement list conversion (ul/ol -> WordList)")]
    public void Test_Html_Lists_RoundTrip() {
        string html = "<ul><li>Item 1<ul><li>Sub 1</li><li>Sub 2</li></ul></li><li>Item 2</li></ul><ol><li>First</li><li>Second</li></ol>";
        
        var doc = html.LoadFromHtml(new HtmlToWordOptions());
        string roundTrip = doc.ToHtml(new WordToHtmlOptions { IncludeListStyles = true });

        Assert.Contains("<ul", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<ol", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Sub 1", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Second", roundTrip, StringComparison.OrdinalIgnoreCase);
    }

    [Fact(Skip = "TODO: Implement table conversion (HTML table -> WordTable)")]
    public void Test_Html_Table_RoundTrip() {
        string html = "<table><tr><td>A</td><td>B</td></tr><tr><td>C</td><td>D</td></tr></table>";
        
        var doc = html.LoadFromHtml(new HtmlToWordOptions());
        string roundTrip = doc.ToHtml(new WordToHtmlOptions());

        Assert.Contains("<table>", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("A", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("D", roundTrip, StringComparison.OrdinalIgnoreCase);
    }

    [Fact(Skip = "TODO: Implement nested table support")]
    public void Test_Html_NestedTable_RoundTrip() {
        string html = "<table><tr><td>Outer</td><td><table><tr><td>Inner</td></tr></table></td></tr></table>";
        
        var doc = html.LoadFromHtml(new HtmlToWordOptions());
        string roundTrip = doc.ToHtml(new WordToHtmlOptions());

        int tableCount = roundTrip.Split(new string[] { "<table>" }, StringSplitOptions.None).Length - 1;
        Assert.True(tableCount >= 2);
        Assert.Contains("Inner", roundTrip, StringComparison.OrdinalIgnoreCase);
    }

    [Fact(Skip = "TODO: Implement image conversion (base64 -> WordImage)")]
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

    [Fact(Skip = "TODO: Implement image conversion (file URL -> WordImage)")]
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

    [Fact]
    public void Test_Html_Headings() {
        string html = "<h1>Heading 1</h1><h2>Heading 2</h2>";

        var doc = html.LoadFromHtml(new HtmlToWordOptions());

        Assert.Equal(WordParagraphStyles.Heading1, doc.Paragraphs[0].Style);
        Assert.Equal("Heading 1", doc.Paragraphs[0].Text);
        Assert.Equal(WordParagraphStyles.Heading2, doc.Paragraphs[1].Style);
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
    public void Test_Html_Lists_Structure() {
        string html = "<ul><li>Item 1<ul><li>Sub 1</li></ul></li><li>Item 2</li></ul>";

        var doc = html.LoadFromHtml(new HtmlToWordOptions());

        Assert.True(doc.Lists.Count > 0);
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
}