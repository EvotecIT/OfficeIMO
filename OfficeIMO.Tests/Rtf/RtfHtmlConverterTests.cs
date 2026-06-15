using OfficeIMO.Rtf;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public partial class RtfHtmlConverterTests {
    [Fact]
    public void RtfDocument_ToHtml_Renders_Text_Formatting_Links_And_Escaping() {
        RtfDocument document = RtfDocument.Create();
        document.Info.Title = "Clinical note";
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.AddText("A < B ");
        paragraph.AddText("bold").SetBold();
        paragraph.AddText(" link").SetItalic().SetHyperlink(new Uri("https://example.test/patient?id=1&tab=note"));

        string html = document.ToHtml(new RtfToHtmlOptions {
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

        string html = document.ToHtml(new RtfToHtmlOptions {
            NewLine = "\n"
        });

        Assert.Equal("<ul><li>Allergy</li>\n<li>Medication</li></ul>", html);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Paragraphs_Inlines_And_Hyperlinks() {
        const string html = "<p>Plain <strong>bold</strong> <em>italic</em> <a href=\"/chart/1\">chart</a><br>next</p>";

        RtfDocument document = html.ToRtfDocument(new HtmlToRtfOptions {
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
    public void Html_ToRtfDocument_Resolves_Base_Element_Hyperlinks() {
        const string html = "<html><head><base href=\"https://example.test/root/\"></head><body><p><a href=\"chart/1\">chart</a></p></body></html>";

        RtfDocument document = html.ToRtfDocument();

        RtfRun run = Assert.Single(Assert.Single(document.Paragraphs).Runs, item => item.Text == "chart");
        Assert.Equal(new Uri("https://example.test/root/chart/1"), run.Hyperlink);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Inline_Css_Formatting_And_Alignment() {
        const string html = "<p style=\"text-align:center !important\">Vitals <span style=\"font-weight:700 !important; font-style: italic; text-decoration: underline line-through; vertical-align: super\">critical</span><span style=\"vertical-align: sub\">low</span></p>";

        RtfDocument document = html.ToRtfDocument();

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

        RtfDocument document = html.ToRtfDocument();

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
    public void Html_ToRtfDocument_Parses_Paragraph_Background_Color() {
        const string html = "<p style=\"background-color:#e6f2ff\">Assessment</p><p>Plain</p>";

        RtfDocument document = html.ToRtfDocument();

        Assert.Single(document.Colors);
        Assert.Equal("#E6F2FF", document.Colors[0].ToString());
        Assert.Equal(1, document.Paragraphs[0].BackgroundColorIndex);
        Assert.Null(document.Paragraphs[1].BackgroundColorIndex);

        string rtf = document.ToRtf();
        Assert.Contains(@"\cbpat1", rtf, StringComparison.Ordinal);

        RtfDocument roundTripDocument = RtfDocument.Read(rtf).Document;
        Assert.Equal(1, roundTripDocument.Paragraphs[0].BackgroundColorIndex);
        Assert.Null(roundTripDocument.Paragraphs[1].BackgroundColorIndex);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Paragraph_Background_Color() {
        RtfDocument document = RtfDocument.Create();
        int background = document.AddColor(230, 242, 255);
        document.AddParagraph("Assessment").SetBackgroundColor(background);
        document.AddParagraph("Plain");

        string html = document.ToHtml(new RtfToHtmlOptions {
            NewLine = "\n"
        });

        Assert.Equal("<p style=\"background-color:#E6F2FF;\">Assessment</p>\n<p>Plain</p>", html);

        RtfDocument roundTripDocument = html.ToRtfDocument();
        Assert.Equal(1, roundTripDocument.Paragraphs[0].BackgroundColorIndex);
        Assert.Null(roundTripDocument.Paragraphs[1].BackgroundColorIndex);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Paragraph_Spacing_And_Line_Height() {
        const string html = "<p style=\"margin-top:6pt;margin-bottom:12pt;line-height:18pt\">Exact</p><p style=\"line-height:150%\">Multiple</p>";

        RtfDocument document = html.ToRtfDocument();

        RtfParagraph exact = document.Paragraphs[0];
        Assert.Equal(120, exact.SpaceBeforeTwips);
        Assert.Equal(240, exact.SpaceAfterTwips);
        Assert.Equal(360, exact.LineSpacingTwips);
        Assert.False(exact.LineSpacingMultiple);

        RtfParagraph multiple = document.Paragraphs[1];
        Assert.Equal(360, multiple.LineSpacingTwips);
        Assert.True(multiple.LineSpacingMultiple);

        string rtf = document.ToRtf();
        Assert.Contains(@"\sb120", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\sa240", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\sl360", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\slmult0", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\slmult1", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Paragraph_Spacing_And_Line_Height() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Exact")
            .SetParagraphSpacing(beforeTwips: 120, afterTwips: 240)
            .SetLineSpacing(360, multiple: false);
        document.AddParagraph("Multiple")
            .SetLineSpacing(360, multiple: true);

        string html = document.ToHtml(new RtfToHtmlOptions {
            NewLine = "\n"
        });

        Assert.Equal("<p style=\"margin-top:6pt;margin-bottom:12pt;line-height:18pt;\">Exact</p>\n<p style=\"line-height:1.5;\">Multiple</p>", html);

        RtfDocument roundTripDocument = html.ToRtfDocument();
        Assert.Equal(120, roundTripDocument.Paragraphs[0].SpaceBeforeTwips);
        Assert.Equal(240, roundTripDocument.Paragraphs[0].SpaceAfterTwips);
        Assert.Equal(360, roundTripDocument.Paragraphs[0].LineSpacingTwips);
        Assert.False(roundTripDocument.Paragraphs[0].LineSpacingMultiple);
        Assert.Equal(360, roundTripDocument.Paragraphs[1].LineSpacingTwips);
        Assert.True(roundTripDocument.Paragraphs[1].LineSpacingMultiple);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Css_Font_Family_And_Size() {
        const string html = "<p><span style=\"font-family: 'Times New Roman', serif; font-size: 13.5pt\">Clinical</span><span style=\"font-family: Consolas, monospace; font-size: 18px\"> code</span></p>";

        RtfDocument document = html.ToRtfDocument();

        Assert.Contains(document.Fonts, font => font.Id == 1 && font.Name == "Times New Roman");
        Assert.Contains(document.Fonts, font => font.Id == 2 && font.Name == "Consolas");

        RtfParagraph paragraph = Assert.Single(document.Paragraphs);
        RtfRun clinical = Assert.Single(paragraph.Runs, run => run.Text == "Clinical");
        Assert.Equal(1, clinical.FontId);
        Assert.Equal(13.5d, clinical.FontSize);

        RtfRun code = Assert.Single(paragraph.Runs, run => run.Text == " code");
        Assert.Equal(2, code.FontId);
        Assert.Equal(13.5d, code.FontSize);

        string rtf = document.ToRtf();
        Assert.Contains(@"{\f1 Times New Roman;}", rtf, StringComparison.Ordinal);
        Assert.Contains(@"{\f2 Consolas;}", rtf, StringComparison.Ordinal);

        RtfDocument roundTrip = RtfDocument.Read(rtf).Document;
        RtfParagraph readParagraph = Assert.Single(roundTrip.Paragraphs);
        Assert.Contains(readParagraph.Runs, run => run.Text == "Clinical" && run.FontId == 1 && run.FontSize == 13.5d);
        Assert.Contains(readParagraph.Runs, run => run.Text == " code" && run.FontId == 2 && run.FontSize == 13.5d);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Run_Font_Styles() {
        RtfDocument document = RtfDocument.Create();
        int fontId = document.AddFont("Times New Roman");
        RtfRun run = document.AddParagraph().AddText("Clinical");
        run.FontId = fontId;
        run.FontSize = 13.5d;

        string html = document.ToHtml();

        Assert.Equal("<p><span style=\"font-family:&quot;Times New Roman&quot;;font-size:13.5pt;\">Clinical</span></p>", html);
    }

    [Fact]
    public void Html_ToRtfDocument_Maps_Headings_To_Outline_Levels() {
        const string html = "<h1>Assessment</h1><h3 style=\"text-align:right\">Plan</h3>";

        RtfDocument document = html.ToRtfDocument();

        Assert.Equal(2, document.Paragraphs.Count);
        Assert.Equal(0, document.Paragraphs[0].OutlineLevel);
        Assert.Contains(document.Paragraphs[0].Runs, run => run.Text == "Assessment" && run.Bold);
        Assert.Equal(2, document.Paragraphs[1].OutlineLevel);
        Assert.Equal(RtfTextAlignment.Right, document.Paragraphs[1].Alignment);

        string rtf = document.ToRtf();
        Assert.Contains(@"\outlinelevel0", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\outlinelevel2", rtf, StringComparison.Ordinal);

        RtfDocument roundTrip = RtfDocument.Read(rtf).Document;
        Assert.Equal(0, roundTrip.Paragraphs[0].OutlineLevel);
        Assert.Equal(2, roundTrip.Paragraphs[1].OutlineLevel);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Outline_Paragraphs_As_Headings() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Assessment").SetOutlineLevel(0);
        document.AddParagraph("Plan").SetOutlineLevel(2).SetAlignment(RtfTextAlignment.Right);

        string html = document.ToHtml(new RtfToHtmlOptions {
            NewLine = "\n"
        });

        Assert.Equal("<h1>Assessment</h1>\n<h3 style=\"text-align:right;\">Plan</h3>", html);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Paragraph_Indentation_Styles() {
        const string html = "<p style=\"margin-left:36pt; margin-right:18pt; text-indent:-12pt\">Indented</p><blockquote>Quoted</blockquote>";

        RtfDocument document = html.ToRtfDocument();

        Assert.Equal(2, document.Paragraphs.Count);
        RtfParagraph indented = document.Paragraphs[0];
        Assert.Equal(720, indented.LeftIndentTwips);
        Assert.Equal(360, indented.RightIndentTwips);
        Assert.Equal(-240, indented.FirstLineIndentTwips);

        RtfParagraph quoted = document.Paragraphs[1];
        Assert.Equal(720, quoted.LeftIndentTwips);

        string rtf = document.ToRtf();
        Assert.Contains(@"\li720", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\ri360", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\fi-240", rtf, StringComparison.Ordinal);

        RtfDocument roundTrip = RtfDocument.Read(rtf).Document;
        Assert.Equal(720, roundTrip.Paragraphs[0].LeftIndentTwips);
        Assert.Equal(360, roundTrip.Paragraphs[0].RightIndentTwips);
        Assert.Equal(-240, roundTrip.Paragraphs[0].FirstLineIndentTwips);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Paragraph_Indentation_Styles() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Indented").SetIndentation(leftTwips: 720, rightTwips: 360, firstLineTwips: -240);

        string html = document.ToHtml();

        Assert.Equal("<p style=\"margin-left:36pt;margin-right:18pt;text-indent:-12pt;\">Indented</p>", html);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Page_Break_Styles() {
        const string html = "<p style=\"page-break-before: always\">Before</p><p style=\"break-after: page\">After</p><p>Next</p>";

        RtfDocument document = html.ToRtfDocument();

        Assert.Equal(3, document.Paragraphs.Count);
        Assert.True(document.Paragraphs[0].PageBreakBefore);
        Assert.Contains(document.Paragraphs[1].Inlines, inline => inline is RtfBreak { Kind: RtfBreakKind.Page });

        string rtf = document.ToRtf();
        Assert.Contains(@"\pagebb", rtf, StringComparison.Ordinal);
        Assert.Contains(@"After\page \par", rtf, StringComparison.Ordinal);

        RtfDocument roundTrip = RtfDocument.Read(rtf).Document;
        Assert.True(roundTrip.Paragraphs[0].PageBreakBefore);
        Assert.Contains(roundTrip.Paragraphs[1].Inlines, inline => inline is RtfBreak { Kind: RtfBreakKind.Page });
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Page_Breaks() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Before").SetPagination(pageBreakBefore: true);
        RtfParagraph after = document.AddParagraph("After");
        after.AddPageBreak();

        string html = document.ToHtml(new RtfToHtmlOptions {
            NewLine = "\n"
        });

        Assert.Equal("<p style=\"page-break-before:always;break-before:page;\">Before</p>\n<p>After<br data-officeimo-rtf-break=\"page\" style=\"page-break-before:always;break-before:page;\"></p>", html);
    }

    [Fact]
    public void Html_ToRtfDocument_Allows_Css_To_Override_Semantic_Formatting() {
        const string html = "<p><strong><em><u>marked <span style=\"font-weight:400; font-style: normal; text-decoration: none; vertical-align: baseline\">plain</span></u></em></strong></p>";

        RtfDocument document = html.ToRtfDocument();

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
    public void Html_Rtf_Html_RoundTrip_Preserves_Semantic_Text() {
        const string html = "<p>Assessment: <strong>stable</strong></p>";

        RtfDocument document = html.ToRtfDocument();
        string rtf = document.ToRtf();
        string roundTripHtml = RtfDocument.Read(rtf).Document.ToHtml();

        Assert.Contains("Assessment:", roundTripHtml, StringComparison.Ordinal);
        Assert.Contains("<strong>stable</strong>", roundTripHtml, StringComparison.Ordinal);
    }
}
