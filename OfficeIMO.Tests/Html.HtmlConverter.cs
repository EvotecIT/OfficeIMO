using OfficeIMO.Html;
using System;
using System.IO;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    [Fact]
    public void Test_Html_RoundTrip() {
        string html = "<p>Hello <b>world</b> and <i>universe</i>.</p>";
        using MemoryStream ms = new MemoryStream();
        HtmlToWordConverter.Convert(html, ms, new HtmlToWordOptions { FontFamily = "Calibri" });

        ms.Position = 0;
        string roundTrip = WordToHtmlConverter.Convert(ms, new WordToHtmlOptions { IncludeStyles = true });

        Assert.Contains("<b>", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("</b>", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("world", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<i>", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("</i>", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("universe", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("font-family:Calibri", roundTrip, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Test_Html_Headings_RoundTrip() {
        string html = "<h1>Heading 1</h1><h2>Heading 2</h2><h3>Heading 3</h3><h4>Heading 4</h4><h5>Heading 5</h5><h6>Heading 6</h6>";
        using MemoryStream ms = new MemoryStream();
        HtmlToWordConverter.Convert(html, ms, new HtmlToWordOptions { FontFamily = "Calibri" });

        ms.Position = 0;
        string roundTrip = WordToHtmlConverter.Convert(ms, new WordToHtmlOptions { IncludeStyles = true });

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
        using MemoryStream ms = new MemoryStream();
        HtmlToWordConverter.Convert(html, ms, new HtmlToWordOptions());

        ms.Position = 0;
        string roundTrip = WordToHtmlConverter.Convert(ms, new WordToHtmlOptions { PreserveListStyles = true });

        Assert.Contains("<ul", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<ol", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Sub 1", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Second", roundTrip, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Test_Html_Table_RoundTrip() {
        string html = "<table><tr><td>A</td><td>B</td></tr><tr><td>C</td><td>D</td></tr></table>";
        using MemoryStream ms = new MemoryStream();
        HtmlToWordConverter.Convert(html, ms, new HtmlToWordOptions());

        ms.Position = 0;
        string roundTrip = WordToHtmlConverter.Convert(ms, new WordToHtmlOptions());

        Assert.Contains("<table>", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("A", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("D", roundTrip, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Test_Html_NestedTable_RoundTrip() {
        string html = "<table><tr><td>Outer</td><td><table><tr><td>Inner</td></tr></table></td></tr></table>";
        using MemoryStream ms = new MemoryStream();
        HtmlToWordConverter.Convert(html, ms, new HtmlToWordOptions());

        ms.Position = 0;
        string roundTrip = WordToHtmlConverter.Convert(ms, new WordToHtmlOptions());

        int tableCount = roundTrip.Split(new string[] { "<table>" }, StringSplitOptions.None).Length - 1;
        Assert.True(tableCount >= 2);
        Assert.Contains("Inner", roundTrip, StringComparison.OrdinalIgnoreCase);
    }
}