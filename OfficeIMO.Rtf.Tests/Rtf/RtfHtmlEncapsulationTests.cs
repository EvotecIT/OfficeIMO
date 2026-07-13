using OfficeIMO.Html;
using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlEncapsulationTests {
    [Fact]
    public void Read_Models_Outlook_Html_Encapsulation_And_Preserves_Lossless_Source() {
        const string rtf = @"{\rtf1\ansi\fromhtml1{\*\htmltag <p><b>Rich</b> message</p>}Plain fallback}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.NotNull(result.Document.HtmlEncapsulation);
        Assert.Equal(1, result.Document.HtmlEncapsulation!.Version);
        Assert.Equal("<p><b>Rich</b> message</p>", result.Document.HtmlEncapsulation.Html);
        Assert.Equal(rtf, result.ToRtfLossless());
    }

    [Fact]
    public void Html_Conversion_Prefers_Encapsulated_Content_Through_Safe_Importer() {
        const string rtf = @"{\rtf1\ansi\fromhtml1{\*\htmltag <p><b>Rich</b> <a href='javascript:alert(1)'>message</a></p>}Plain fallback}";
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
        const string rtf = @"{\rtf1\ansi\fromhtml1{\*\htmltag <p><b>Rich</b> message</p>}Plain fallback}";
        RtfDocument document = RtfDocument.Read(rtf).Document;

        string html = document.ToHtml(new RtfToHtmlOptions { PreferEncapsulatedHtml = false });

        Assert.Contains("Plain fallback", html, StringComparison.Ordinal);
        Assert.DoesNotContain("<strong>Rich</strong>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Normalized_Writer_Reemits_Encapsulation_Unless_Disabled() {
        RtfDocument document = RtfDocument.Create();
        document.HtmlEncapsulation = new RtfHtmlEncapsulation(1, "<p>Rich Ω</p>");
        document.AddParagraph("Fallback");

        string normalized = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        string withoutHtml = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false, IncludeHtmlEncapsulation = false });
        RtfDocument roundTrip = RtfDocument.Read(normalized).Document;

        Assert.Contains(@"\fromhtml1", normalized, StringComparison.Ordinal);
        Assert.Contains(@"{\*\htmltag ", normalized, StringComparison.Ordinal);
        Assert.Equal("<p>Rich Ω</p>", roundTrip.HtmlEncapsulation!.Html);
        Assert.DoesNotContain(@"\fromhtml", withoutHtml, StringComparison.Ordinal);
    }
}
