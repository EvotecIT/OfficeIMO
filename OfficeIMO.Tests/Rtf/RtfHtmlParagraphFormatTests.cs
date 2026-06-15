using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Html;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlParagraphFormatTests {
    [Fact]
    public void Html_ToRtfDocument_Parses_Paragraph_Borders() {
        const string html = "<p style=\"border:1pt solid #0c2238;border-left:2pt double red\">Boxed</p><p>Plain</p>";

        RtfDocument document = html.LoadFromHtml();

        RtfParagraph boxed = document.Paragraphs[0];
        Assert.Equal(RtfParagraphBorderStyle.Single, boxed.TopBorder.Style);
        Assert.Equal(20, boxed.TopBorder.Width);
        Assert.Equal(1, boxed.TopBorder.ColorIndex);
        Assert.Equal(RtfParagraphBorderStyle.Double, boxed.LeftBorder.Style);
        Assert.Equal(40, boxed.LeftBorder.Width);
        Assert.Equal(2, boxed.LeftBorder.ColorIndex);
        Assert.Equal(RtfParagraphBorderStyle.Single, boxed.RightBorder.Style);
        Assert.False(document.Paragraphs[1].TopBorder.HasAnyValue);

        string rtf = document.ToRtf();
        Assert.Contains(@"\brdrt\brdrs\brdrw20\brdrcf1", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\brdrl\brdrdb\brdrw40\brdrcf2", rtf, StringComparison.Ordinal);

        RtfParagraph roundTripBoxed = RtfDocument.Read(rtf).Document.Paragraphs[0];
        Assert.Equal(RtfParagraphBorderStyle.Single, roundTripBoxed.TopBorder.Style);
        Assert.Equal(RtfParagraphBorderStyle.Double, roundTripBoxed.LeftBorder.Style);
        Assert.Equal(40, roundTripBoxed.LeftBorder.Width);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Paragraph_Borders() {
        RtfDocument document = RtfDocument.Create();
        int dark = document.AddColor(12, 34, 56);
        int red = document.AddColor(255, 0, 0);
        document.AddParagraph("Boxed")
            .SetBorder(RtfParagraphBorderSide.Top, RtfParagraphBorderStyle.Single, width: 20, colorIndex: dark)
            .SetBorder(RtfParagraphBorderSide.Left, RtfParagraphBorderStyle.Double, width: 40, colorIndex: red)
            .SetBorder(RtfParagraphBorderSide.Bottom, RtfParagraphBorderStyle.Dotted);

        string html = document.ToHtml();

        Assert.Equal("<p style=\"border-top:1pt solid #0C2238;border-left:2pt double #FF0000;border-bottom:dotted;\">Boxed</p>", html);

        RtfParagraph roundTripBoxed = html.LoadFromHtml().Paragraphs[0];
        Assert.Equal(RtfParagraphBorderStyle.Single, roundTripBoxed.TopBorder.Style);
        Assert.Equal(20, roundTripBoxed.TopBorder.Width);
        Assert.Equal(1, roundTripBoxed.TopBorder.ColorIndex);
        Assert.Equal(RtfParagraphBorderStyle.Double, roundTripBoxed.LeftBorder.Style);
        Assert.Equal(40, roundTripBoxed.LeftBorder.Width);
        Assert.Equal(2, roundTripBoxed.LeftBorder.ColorIndex);
        Assert.Equal(RtfParagraphBorderStyle.Dotted, roundTripBoxed.BottomBorder.Style);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Document_And_Paragraph_Language_Direction() {
        const string html = "<html dir=\"rtl\" lang=\"ar-SA\"><body><p dir=\"ltr\">LTR <span lang=\"pl-PL\">Polish</span></p><p>RTL default</p></body></html>";

        RtfDocument document = html.LoadFromHtml();

        Assert.Equal(1025, document.Settings.DefaultLanguageId);
        Assert.Equal(RtfTextDirection.RightToLeft, document.Settings.Direction);
        Assert.Equal(RtfTextDirection.LeftToRight, document.Paragraphs[0].Direction);
        Assert.Null(document.Paragraphs[1].Direction);
        Assert.Null(document.Paragraphs[0].Runs.Single(run => run.Text == "LTR ").Direction);
        Assert.Null(document.Paragraphs[0].Runs.Single(run => run.Text == "LTR ").LanguageId);
        Assert.Equal(1045, document.Paragraphs[0].Runs.Single(run => run.Text == "Polish").LanguageId);
        Assert.Null(document.Paragraphs[1].Runs.Single(run => run.Text == "RTL default").Direction);

        string rtf = document.ToRtf();
        Assert.Contains(@"\deflang1025", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\rtldoc", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\pard\ltrpar", rtf, StringComparison.Ordinal);

        RtfDocument roundTrip = RtfDocument.Read(rtf).Document;
        Assert.Equal(1025, roundTrip.Settings.DefaultLanguageId);
        Assert.Equal(RtfTextDirection.RightToLeft, roundTrip.Settings.Direction);
        Assert.Equal(RtfTextDirection.LeftToRight, roundTrip.Paragraphs[0].Direction);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Document_And_Paragraph_Language_Direction() {
        RtfDocument document = RtfDocument.Create();
        document.Settings
            .SetDefaultLanguage(1025)
            .SetDirection(RtfTextDirection.RightToLeft);
        document.AddParagraph("LTR")
            .SetDirection(RtfTextDirection.LeftToRight);
        document.AddParagraph("Default");

        string html = document.ToHtml(new RtfHtmlSaveOptions { FragmentOnly = false });

        Assert.Contains("<html lang=\"ar-SA\" dir=\"rtl\" style=\"--officeimo-rtf-lang:1025;direction:rtl;unicode-bidi:isolate;--officeimo-rtf-direction:rtl;\">", html, StringComparison.Ordinal);
        Assert.Contains("<p dir=\"ltr\" style=\"direction:ltr;unicode-bidi:isolate;--officeimo-rtf-direction:ltr;\">LTR</p>", html, StringComparison.Ordinal);

        RtfDocument roundTrip = html.LoadFromHtml();
        Assert.Equal(1025, roundTrip.Settings.DefaultLanguageId);
        Assert.Equal(RtfTextDirection.RightToLeft, roundTrip.Settings.Direction);
        Assert.Equal(RtfTextDirection.LeftToRight, roundTrip.Paragraphs[0].Direction);
        Assert.Null(roundTrip.Paragraphs[1].Runs.Single(run => run.Text == "Default").Direction);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Inline_Page_And_Column_Breaks() {
        const string html = "<p>Before<br data-officeimo-rtf-break=\"page\">After<br data-officeimo-rtf-break=\"column\">Column<br style=\"page-break-before:always\">Styled</p>";

        RtfDocument document = html.LoadFromHtml();

        RtfParagraph paragraph = Assert.Single(document.Paragraphs);
        Assert.Collection(paragraph.Inlines,
            inline => Assert.Equal("Before", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(RtfBreakKind.Page, Assert.IsType<RtfBreak>(inline).Kind),
            inline => Assert.Equal("After", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(RtfBreakKind.Column, Assert.IsType<RtfBreak>(inline).Kind),
            inline => Assert.Equal("Column", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(RtfBreakKind.Page, Assert.IsType<RtfBreak>(inline).Kind),
            inline => Assert.Equal("Styled", Assert.IsType<RtfRun>(inline).Text));

        string rtf = document.ToRtf();
        Assert.Contains(@"\page", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\column", rtf, StringComparison.Ordinal);

        RtfParagraph roundTripParagraph = Assert.Single(RtfDocument.Read(rtf).Document.Paragraphs);
        Assert.Contains(roundTripParagraph.Inlines, inline => inline is RtfBreak { Kind: RtfBreakKind.Page });
        Assert.Contains(roundTripParagraph.Inlines, inline => inline is RtfBreak { Kind: RtfBreakKind.Column });
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Inline_Page_And_Column_Break_Metadata() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.AddText("Before");
        paragraph.AddPageBreak();
        paragraph.AddText("After");
        paragraph.AddColumnBreak();
        paragraph.AddText("Column");

        string html = document.ToHtml();

        Assert.Equal("<p>Before<br data-officeimo-rtf-break=\"page\" style=\"page-break-before:always;break-before:page;\">After<br data-officeimo-rtf-break=\"column\" style=\"break-before:column;\">Column</p>", html);

        RtfParagraph roundTripParagraph = Assert.Single(html.LoadFromHtml().Paragraphs);
        Assert.Collection(roundTripParagraph.Inlines,
            inline => Assert.Equal("Before", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(RtfBreakKind.Page, Assert.IsType<RtfBreak>(inline).Kind),
            inline => Assert.Equal("After", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(RtfBreakKind.Column, Assert.IsType<RtfBreak>(inline).Kind),
            inline => Assert.Equal("Column", Assert.IsType<RtfRun>(inline).Text));
    }
}
