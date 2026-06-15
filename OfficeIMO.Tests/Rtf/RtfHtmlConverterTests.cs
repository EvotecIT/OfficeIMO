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
    public void Html_ToRtfDocument_Parses_Inline_Css_Formatting_And_Alignment() {
        const string html = "<p style=\"text-align:center !important\">Vitals <span style=\"font-weight:700 !important; font-style: italic; text-decoration: underline line-through; vertical-align: super\">critical</span><span style=\"vertical-align: sub\">low</span></p>";

        RtfDocument document = html.ToRtfDocumentFromHtml();

        RtfParagraph paragraph = Assert.Single(document.Paragraphs);
        Assert.Equal(RtfTextAlignment.Center, paragraph.Alignment);
        RtfRun critical = Assert.Single(paragraph.Runs, run => run.Text == "critical");
        Assert.True(critical.Bold);
        Assert.True(critical.Italic);
        Assert.True(critical.Underline);
        Assert.True(critical.Strike);
        Assert.Equal(RtfVerticalPosition.Superscript, critical.VerticalPosition);

        RtfRun low = Assert.Single(paragraph.Runs, run => run.Text == "low");
        Assert.Equal(RtfVerticalPosition.Subscript, low.VerticalPosition);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Css_Colors_Into_Rtf_Color_Table() {
        const string html = "<p><span style=\"color:#0c2238; background-color: rgb(255, 242, 204)\">Flag</span><span style=\"color: #0C2238\"> again</span><span style=\"background: yellow\"> note</span></p>";

        RtfDocument document = html.ToRtfDocumentFromHtml();

        Assert.Equal(3, document.Colors.Count);
        Assert.Equal("#0C2238", document.Colors[0].ToString());
        Assert.Equal("#FFF2CC", document.Colors[1].ToString());
        Assert.Equal("#FFFF00", document.Colors[2].ToString());

        RtfParagraph paragraph = Assert.Single(document.Paragraphs);
        RtfRun flag = Assert.Single(paragraph.Runs, run => run.Text == "Flag");
        Assert.Equal(1, flag.ForegroundColorIndex);
        Assert.Equal(2, flag.CharacterBackgroundColorIndex);

        RtfRun again = Assert.Single(paragraph.Runs, run => run.Text == " again");
        Assert.Equal(1, again.ForegroundColorIndex);
        Assert.Null(again.CharacterBackgroundColorIndex);

        RtfRun note = Assert.Single(paragraph.Runs, run => run.Text == " note");
        Assert.Null(note.ForegroundColorIndex);
        Assert.Equal(3, note.CharacterBackgroundColorIndex);

        string rtf = document.ToRtf();
        Assert.Contains(@"{\colortbl;\red12\green34\blue56;\red255\green242\blue204;\red255\green255\blue0;}", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\cf1 \chcbpat2 Flag\chcbpat0  again\cf0", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Run_Color_Styles() {
        RtfDocument document = RtfDocument.Create();
        int foreground = document.AddColor(12, 34, 56);
        int background = document.AddColor(255, 242, 204);
        RtfRun run = document.AddParagraph().AddText("Flag");
        run.ForegroundColorIndex = foreground;
        run.CharacterBackgroundColorIndex = background;

        string html = document.ToHtml();

        Assert.Equal("<p><span style=\"color:#0C2238;background-color:#FFF2CC;\">Flag</span></p>", html);
    }

    [Fact]
    public void Html_ToRtfDocument_Allows_Css_To_Override_Semantic_Formatting() {
        const string html = "<p><strong><em><u>marked <span style=\"font-weight:400; font-style: normal; text-decoration: none; vertical-align: baseline\">plain</span></u></em></strong></p>";

        RtfDocument document = html.ToRtfDocumentFromHtml();

        RtfParagraph paragraph = Assert.Single(document.Paragraphs);
        RtfRun marked = Assert.Single(paragraph.Runs, run => run.Text == "marked ");
        Assert.True(marked.Bold);
        Assert.True(marked.Italic);
        Assert.True(marked.Underline);

        RtfRun plain = Assert.Single(paragraph.Runs, run => run.Text == "plain");
        Assert.False(plain.Bold);
        Assert.False(plain.Italic);
        Assert.False(plain.Underline);
        Assert.False(plain.Strike);
        Assert.Equal(RtfVerticalPosition.Baseline, plain.VerticalPosition);
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
