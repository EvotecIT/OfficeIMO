using System;
using System.IO;
using OfficeIMO.Html;
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
}
