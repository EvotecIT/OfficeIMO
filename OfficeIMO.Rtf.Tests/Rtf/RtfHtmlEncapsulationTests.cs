using OfficeIMO.Html;
using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlEncapsulationTests {
    [Theory]
    [InlineData("replace")]
    [InlineData("run")]
    [InlineData("append")]
    [InlineData("format")]
    public void Semantic_Edits_Invalidate_Original_Html_For_All_Exports(string mutation) {
        RtfDocument document = RtfDocument.Read(@"{\rtf1\ansi\fromhtml1{\*\htmltag <p>Original HTML</p>}\htmlrtf1 Original fallback\par}").Document;
        Assert.True(document.IsHtmlEncapsulationCurrent);
        RtfDocument clone = document.Clone();
        switch (mutation) {
            case "replace": clone.ReplaceText("Original", "Edited"); break;
            case "run": clone.Paragraphs[0].Runs[0].Text = "Edited text"; break;
            case "append": clone.AddParagraph("Edited paragraph"); break;
            case "format": clone.Paragraphs[0].Runs[0].Bold = true; break;
        }

        Assert.False(clone.IsHtmlEncapsulationCurrent);
        Assert.True(document.IsHtmlEncapsulationCurrent);
        RtfToHtmlResult converted = clone.ToHtmlResult();
        Assert.DoesNotContain("Original HTML", converted.Value, StringComparison.Ordinal);
        Assert.DoesNotContain(@"\fromhtml", clone.ToRtf(), StringComparison.Ordinal);
        Assert.Contains(converted.RtfReport.Diagnostics, diagnostic => diagnostic.Code == "RtfHtmlEncapsulationInvalidated");
        if (mutation == "format") Assert.Contains("<strong>", converted.Value, StringComparison.Ordinal);
        else Assert.Contains("Edited", converted.Value, StringComparison.Ordinal);
    }

    [Fact]
    public void Committed_Outlook_Fixture_Retains_Message_List_And_Table_Text() {
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "RtfCorpus", "producers",
            "microsoft-outlook", "outlook-16-mapi-html-encapsulated.rtf");
        RtfDocument document = RtfDocument.Load(path);
        string extracted = System.Net.WebUtility.HtmlDecode(document.HtmlEncapsulation!.Html);
        RtfToHtmlResult converted = document.ToHtmlResult();
        string html = System.Net.WebUtility.HtmlDecode(converted.Value);
        foreach (string text in new[] { "Outlook MAPI HTML fixture", "Zażółć gęślą jaźń", "First", "Second", "Key", "Value", "Bandage", "Ready" }) {
            Assert.Contains(text, extracted, StringComparison.Ordinal);
            Assert.Contains(text, html, StringComparison.Ordinal);
        }
        Assert.Contains("<table", html, StringComparison.Ordinal);
        Assert.False(converted.RtfReport.HasLoss);
    }

    [Theory]
    [InlineData(@"{\rtf1\ansi\fromhtml1{\*\htmltag <p>}Visible{\*\htmltag </p>}}", "<p>Visible</p>")]
    [InlineData(@"{\rtf1\ansi\fromhtml1{\*\htmltag <p>}A{\htmlrtf1 Hidden\htmlrtf0 B}C{\*\htmltag </p>}}", "<p>ABC</p>")]
    [InlineData(@"{\rtf1\ansi\fromhtml1{\*\mhtmltag <img src='rewritten'>}{\*\htmltag <p>}Original{\*\htmltag </p>}}", "<p>Original</p>")]
    [InlineData(@"{\rtf1\ansi\fromhtml1{\*\htmltag <p>}A & B < C{\*\htmltag </p>}}", "<p>A &amp; B &lt; C</p>")]
    [InlineData(@"{\rtf1\ansi\ansicpg1252\deff0\fromhtml1{\fonttbl{\f0\fcharset0 Arial;}{\f1\fcharset204 Arial;}}{\*\htmltag <p>}\htmlrtf1\f1\htmlrtf0 \'c0{\*\htmltag </p>}}", "<p>А</p>")]
    public void Extract_Uses_Shared_Text_Scoped_Suppression_And_Original_Tag_Content(string input, string expected) {
        string html = RtfDocument.Read(input).Document.HtmlEncapsulation!.Html;
        // HtmlEncode may choose equivalent numeric references for non-ASCII characters.
        Assert.Equal(System.Net.WebUtility.HtmlDecode(expected), System.Net.WebUtility.HtmlDecode(html));
    }

    [Fact]
    public void Read_Models_Outlook_Html_Encapsulation_And_Preserves_Lossless_Source() {
        const string rtf = @"{\rtf1\ansi\fromhtml1{\*\htmltag <p><b>Rich</b> message</p>}\htmlrtf1 Plain fallback}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.NotNull(result.Document.HtmlEncapsulation);
        Assert.Equal(1, result.Document.HtmlEncapsulation!.Version);
        Assert.Equal("<p><b>Rich</b> message</p>", result.Document.HtmlEncapsulation.Html);
        Assert.Equal(rtf, result.ToRtfLossless());
    }

    [Fact]
    public void Html_Conversion_Prefers_Encapsulated_Content_Through_Safe_Importer() {
        const string rtf = @"{\rtf1\ansi\fromhtml1{\*\htmltag <p><b>Rich</b> <a href='javascript:alert(1)'>message</a></p>}\htmlrtf1 Plain fallback}";
        RtfDocument document = RtfDocument.Read(rtf).Document;
        var options = new RtfToHtmlOptions();

        RtfToHtmlResult result = document.ToHtmlResult(options);
        string html = result.Value;

        Assert.Contains("<strong>Rich</strong>", html, StringComparison.Ordinal);
        Assert.Contains("message", html, StringComparison.Ordinal);
        Assert.DoesNotContain("Plain fallback", html, StringComparison.Ordinal);
        Assert.DoesNotContain("javascript", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(result.RtfDiagnostics, diagnostic => diagnostic.Code == "RtfHtmlEncapsulatedHtmlUsed");
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "HyperlinkRejectedByPolicy");
    }

    [Fact]
    public void Html_Conversion_Can_Use_Rtf_Fallback_Explicitly() {
        const string rtf = @"{\rtf1\ansi\fromhtml1{\*\htmltag <p><b>Rich</b> message</p>}\htmlrtf1 Plain fallback}";
        RtfDocument document = RtfDocument.Read(rtf).Document;

        string html = document.ToHtml(new RtfToHtmlOptions { PreferEncapsulatedHtml = false });

        Assert.Contains("Plain fallback", html, StringComparison.Ordinal);
        Assert.DoesNotContain("<strong>Rich</strong>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Normalized_Writer_Reemits_Encapsulation_Unless_Disabled() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Fallback");
        document.HtmlEncapsulation = new RtfHtmlEncapsulation(1, "<p>Rich Ω</p>");

        string normalized = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        string withoutHtml = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false, IncludeHtmlEncapsulation = false });
        RtfDocument roundTrip = RtfDocument.Read(normalized).Document;

        Assert.Contains(@"\fromhtml1", normalized, StringComparison.Ordinal);
        Assert.Contains(@"{\*\htmltag ", normalized, StringComparison.Ordinal);
        Assert.Equal("<p>Rich Ω</p>", roundTrip.HtmlEncapsulation!.Html);
        Assert.DoesNotContain(@"\fromhtml", withoutHtml, StringComparison.Ordinal);
    }
}
