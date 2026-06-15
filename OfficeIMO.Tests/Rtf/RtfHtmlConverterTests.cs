using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Html;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlConverterTests {
    [Fact]
    public void RtfDocument_ToHtml_Renders_Text_Formatting_Links_And_Escaping() {
        RtfDocument document = RtfDocument.Create();
        document.Info.Title = "Clinical note";
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.AddText("A < B ");
        paragraph.AddText("bold").SetBold();
        paragraph.AddText(" link").SetItalic().SetHyperlink(new Uri("https://example.test/patient?id=1&tab=note"));

        string html = document.ToHtml(new RtfHtmlSaveOptions {
            FragmentOnly = false
        });

        Assert.Contains("<title>Clinical note</title>", html, StringComparison.Ordinal);
        Assert.Contains("A &lt; B", html, StringComparison.Ordinal);
        Assert.Contains("<strong>bold</strong>", html, StringComparison.Ordinal);
        Assert.Contains("<a href=\"https://example.test/patient?id=1&amp;tab=note\"><em> link</em></a>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void RtfDocument_ToHtml_Wraps_List_Paragraphs() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Allergy").ListKind = RtfListKind.Bullet;
        document.AddParagraph("Medication").ListKind = RtfListKind.Bullet;

        string html = document.ToHtml(new RtfHtmlSaveOptions {
            NewLine = "\n"
        });

        Assert.Equal("<ul><li>Allergy</li>\n<li>Medication</li></ul>", html);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Paragraphs_Inlines_And_Hyperlinks() {
        const string html = "<p>Plain <strong>bold</strong> <em>italic</em> <a href=\"/chart/1\">chart</a><br>next</p>";

        RtfDocument document = html.ToRtfDocumentFromHtml(new RtfHtmlReadOptions {
            BaseUri = new Uri("https://example.test")
        });

        Assert.Single(document.Paragraphs);
        RtfParagraph paragraph = document.Paragraphs[0];
        Assert.Contains(paragraph.Runs, run => run.Text == "bold" && run.Bold);
        Assert.Contains(paragraph.Runs, run => run.Text == "italic" && run.Italic);
        Assert.Contains(paragraph.Runs, run => run.Text == "chart" && run.Hyperlink == new Uri("https://example.test/chart/1"));
        Assert.Contains(paragraph.Inlines, inline => inline is RtfBreak { Kind: RtfBreakKind.Line });
        Assert.Contains(paragraph.Runs, run => run.Text == "next");
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Lists_And_Tables() {
        const string html = "<ul><li>Allergy</li><li><strong>Medication</strong></li></ul><table><tr><th>Name</th><th>Value</th></tr><tr><td>Pulse</td><td>72</td></tr></table>";

        RtfDocument document = html.ToRtfDocumentFromHtml();

        Assert.Equal(RtfListKind.Bullet, document.Paragraphs[0].ListKind);
        Assert.Equal("Allergy", document.Paragraphs[0].ToPlainText());
        Assert.Equal(RtfListKind.Bullet, document.Paragraphs[1].ListKind);
        Assert.Contains(document.Paragraphs[1].Runs, run => run.Text == "Medication" && run.Bold);

        RtfTable table = Assert.IsType<RtfTable>(document.Blocks[2]);
        Assert.Equal(2, table.Rows.Count);
        Assert.Equal("Name", table.Rows[0].Cells[0].Paragraphs[0].ToPlainText());
        Assert.Equal("72", table.Rows[1].Cells[1].Paragraphs[0].ToPlainText());
    }

    [Fact]
    public void Html_Rtf_Html_RoundTrip_Preserves_Semantic_Text() {
        const string html = "<p>Assessment: <strong>stable</strong></p>";

        RtfDocument document = html.ToRtfDocumentFromHtml();
        string rtf = document.ToRtf();
        string roundTripHtml = RtfDocument.Read(rtf).Document.ToHtml();

        Assert.Contains("Assessment:", roundTripHtml, StringComparison.Ordinal);
        Assert.Contains("<strong>stable</strong>", roundTripHtml, StringComparison.Ordinal);
    }
}
