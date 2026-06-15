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
    public void Html_ToRtfDocument_Parses_Css_Font_Family_And_Size() {
        const string html = "<p><span style=\"font-family: 'Times New Roman', serif; font-size: 13.5pt\">Clinical</span><span style=\"font-family: Consolas, monospace; font-size: 18px\"> code</span></p>";

        RtfDocument document = html.ToRtfDocumentFromHtml();

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

        RtfDocument document = html.ToRtfDocumentFromHtml();

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

        string html = document.ToHtml(new RtfHtmlSaveOptions {
            NewLine = "\n"
        });

        Assert.Equal("<h1>Assessment</h1>\n<h3 style=\"text-align:right;\">Plan</h3>", html);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Paragraph_Indentation_Styles() {
        const string html = "<p style=\"margin-left:36pt; margin-right:18pt; text-indent:-12pt\">Indented</p><blockquote>Quoted</blockquote>";

        RtfDocument document = html.ToRtfDocumentFromHtml();

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

        RtfDocument document = html.ToRtfDocumentFromHtml();

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

        string html = document.ToHtml(new RtfHtmlSaveOptions {
            NewLine = "\n"
        });

        Assert.Equal("<p style=\"page-break-before:always;break-before:page;\">Before</p>\n<p>After<br style=\"page-break-before:always;break-before:page;\"></p>", html);
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
    public void Html_ToRtfDocument_Parses_Table_Header_And_Cell_Styles() {
        const string html = "<table><thead><tr><th style=\"background-color:#f2f2f2;width:25%;vertical-align:middle\">Name</th><th style=\"text-align:right;width:72pt\">Value</th></tr></thead><tbody><tr><td style=\"background:#fff2cc;vertical-align:bottom\">Pulse</td><td>72</td></tr></tbody></table>";

        RtfDocument document = html.ToRtfDocumentFromHtml();

        RtfTable table = Assert.IsType<RtfTable>(Assert.Single(document.Blocks));
        Assert.True(table.Rows[0].RepeatHeader);
        RtfTableCell firstHeader = table.Rows[0].Cells[0];
        Assert.Equal(1, firstHeader.BackgroundColorIndex);
        Assert.Equal(1250, firstHeader.PreferredWidth);
        Assert.Equal(RtfTableWidthUnit.Percent, firstHeader.PreferredWidthUnit);
        Assert.Equal(RtfTableCellVerticalAlignment.Center, firstHeader.VerticalAlignment);
        Assert.Equal(RtfTextAlignment.Center, firstHeader.Paragraphs[0].Alignment);
        Assert.Contains(firstHeader.Paragraphs[0].Runs, run => run.Text == "Name" && run.Bold);

        RtfTableCell secondHeader = table.Rows[0].Cells[1];
        Assert.Equal(1440, secondHeader.PreferredWidth);
        Assert.Equal(RtfTableWidthUnit.Twips, secondHeader.PreferredWidthUnit);
        Assert.Equal(RtfTextAlignment.Right, secondHeader.Paragraphs[0].Alignment);

        RtfTableCell pulseCell = table.Rows[1].Cells[0];
        Assert.Equal(2, pulseCell.BackgroundColorIndex);
        Assert.Equal(RtfTableCellVerticalAlignment.Bottom, pulseCell.VerticalAlignment);

        string rtf = document.ToRtf();
        Assert.Contains(@"\trhdr", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clftsWidth2\clwWidth1250", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clftsWidth3\clwWidth1440", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clvertalc", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clvertalb", rtf, StringComparison.Ordinal);

        RtfTable roundTripTable = Assert.IsType<RtfTable>(Assert.Single(RtfDocument.Read(rtf).Document.Blocks));
        Assert.True(roundTripTable.Rows[0].RepeatHeader);
        Assert.Equal(RtfTableWidthUnit.Percent, roundTripTable.Rows[0].Cells[0].PreferredWidthUnit);
        Assert.Equal(1250, roundTripTable.Rows[0].Cells[0].PreferredWidth);
        Assert.Equal(RtfTableCellVerticalAlignment.Bottom, roundTripTable.Rows[1].Cells[0].VerticalAlignment);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Table_Header_And_Cell_Styles() {
        RtfDocument document = RtfDocument.Create();
        int headerBackground = document.AddColor(242, 242, 242);
        int bodyBackground = document.AddColor(255, 242, 204);
        RtfTable table = document.AddTable(2, 2);
        table.Rows[0].RepeatHeader = true;
        table.Rows[0].Cells[0]
            .SetBackgroundColor(headerBackground)
            .SetPreferredWidth(1250, RtfTableWidthUnit.Percent);
        table.Rows[0].Cells[0].VerticalAlignment = RtfTableCellVerticalAlignment.Center;
        table.Rows[0].Cells[0].AddParagraph("Name");
        table.Rows[0].Cells[1]
            .SetPreferredWidth(1440, RtfTableWidthUnit.Twips)
            .AddParagraph("Value");
        table.Rows[1].Cells[0].SetBackgroundColor(bodyBackground);
        table.Rows[1].Cells[0].VerticalAlignment = RtfTableCellVerticalAlignment.Bottom;
        table.Rows[1].Cells[0].AddParagraph("Pulse");
        table.Rows[1].Cells[1].AddParagraph("72");

        string html = document.ToHtml();

        Assert.Equal("<table><thead><tr><th style=\"background-color:#F2F2F2;width:25%;vertical-align:middle;\"><p>Name</p></th><th style=\"width:72pt;\"><p>Value</p></th></tr></thead><tbody><tr><td style=\"background-color:#FFF2CC;vertical-align:bottom;\"><p>Pulse</p></td><td><p>72</p></td></tr></tbody></table>", html);
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
