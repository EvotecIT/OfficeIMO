using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using AngleSharp.Html.Parser;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordHtmlDefaultDocumentTests {
    [Theory]
    [InlineData("del")]
    [InlineData("ins")]
    [InlineData("mark")]
    [InlineData("cite")]
    [InlineData("dfn")]
    [InlineData("time")]
    [InlineData("code")]
    public void InitialStyledTextKeepsItsSemanticTagAndMetadata(string tag) {
        string attributes = tag == "time" ? " datetime='2026-10-28'" : string.Empty;
        using var document = OfficeIMO.Html.HtmlConversionDocument.Parse(
            "<p style='color:#123456'><" + tag + attributes + ">Today</" + tag + "></p>").ToWord();
        if (tag == "code") Assert.Equal("HtmlCode", Assert.Single(document.Paragraphs, paragraph => paragraph.Text == "Today").CharacterStyleId);
        var parsed = new HtmlParser().ParseDocument(document.ToHtml());
        var semantic = Assert.Single(parsed.QuerySelectorAll("body " + tag));
        Assert.Equal("Today", semantic.TextContent);
        if (tag == "time") Assert.Equal("2026-10-28", semantic.GetAttribute("datetime"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task DefaultExportsWriteAStyledHtml5Document(bool asynchronous) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Delivery review");
        document.AddTable(1, 2).Rows[0].Cells[0].Paragraphs[0].Text = "Region";
        using var output = new MemoryStream();
        if (asynchronous) await document.SaveAsHtmlAsync(output);
        else document.SaveAsHtml(output);
        string html = Encoding.UTF8.GetString(output.ToArray());
        Assert.StartsWith("<!doctype html>", html, StringComparison.OrdinalIgnoreCase);
        var parsed = new HtmlParser().ParseDocument(html);
        Assert.NotNull(parsed.Head?.QuerySelector("meta[charset]"));
        Assert.NotNull(parsed.Head?.QuerySelector("meta[name=viewport]"));
        Assert.Contains("--office-font-body", parsed.Head?.QuerySelector("style")?.TextContent ?? string.Empty);
        Assert.Contains("Delivery review", parsed.Body?.TextContent ?? string.Empty, StringComparison.Ordinal);
        Assert.NotNull(parsed.Body?.QuerySelector("table"));
        Assert.StartsWith("<!doctype html>", document.ToHtml(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void FragmentAndCssOptOutsRemainExplicit() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Embedded article");
        string html = document.ToHtml(new WordHtmlExportOptions { EmitDocumentShell = false, IncludeDefaultCss = false });
        Assert.DoesNotContain("<!doctype", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("<style", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Embedded article", html, StringComparison.Ordinal);
    }
}
