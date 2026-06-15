using OfficeIMO.Rtf;
using OfficeIMO.Html.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlSectionLayoutTests {
    [Fact]
    public void RtfDocument_ToHtml_RoundTrips_Document_Page_Setup_And_Note_Settings() {
        RtfDocument document = RtfDocument.Create();
        int red = document.AddColor(255, 0, 0);
        int blue = document.AddColor(0, 0, 255);
        document.PageSetup.SetPaperSize(16838, 11906)
            .SetPrinterPaper(paperSize: 9, firstPageSource: 7, otherPagesSource: 8)
            .SetMargins(leftTwips: 1440, rightTwips: 720, topTwips: 1080, bottomTwips: 1080)
            .SetGutter(180, rtlGutter: true)
            .SetHeaderFooterDistance(headerDistanceTwips: 360, footerDistanceTwips: 540)
            .SetPageNumbering(start: 5, restart: true, format: RtfPageNumberFormat.LowerRoman, positionXTwips: 720, positionYTwips: 900)
            .SetLandscape()
            .SetDifferentFirstPageHeaderFooter();
        document.PageSetup.PageBorders.IncludeHeader = true;
        document.PageSetup.PageBorders.IncludeFooter = true;
        document.PageSetup.PageBorders.SnapToPageBorder = true;
        document.PageSetup.PageBorders.SetDisplayOptions(RtfPageBorderScope.AllExceptFirstPageInSection, displayBehindText: false, RtfPageBorderOffset.PageEdge);
        document.PageSetup.PageBorders.Top.Set(RtfPageBorderStyle.Single, width: 12, space: 24, colorIndex: red);
        document.PageSetup.PageBorders.Bottom.Set(RtfPageBorderStyle.Double, width: 18, space: 30, colorIndex: blue);
        document.NoteSettings
            .SetFootnoteNumbering(start: 5, restart: RtfNoteNumberRestart.EachSection, format: RtfNoteNumberFormat.LowerRoman)
            .SetFootnotePlacement(RtfFootnotePlacement.BeneathText)
            .SetEndnoteNumbering(start: 7, restart: RtfNoteNumberRestart.Continuous, format: RtfNoteNumberFormat.Arabic)
            .SetEndnotePlacement(RtfEndnotePlacement.DocumentEnd);
        document.AddParagraph("Landscape body");

        string html = document.ToHtml(new RtfHtmlSaveOptions { FragmentOnly = false, NewLine = "\n" });

        Assert.Contains("<meta name=\"officeimo-rtf-colors\" content=\"", html, StringComparison.Ordinal);
        Assert.Contains("<meta name=\"officeimo-rtf-document-layout\" content=\"", html, StringComparison.Ordinal);

        RtfDocument roundTrip = html.LoadFromHtml();
        Assert.Equal(2, roundTrip.Colors.Count);
        Assert.Equal(16838, roundTrip.PageSetup.PaperWidthTwips);
        Assert.Equal(11906, roundTrip.PageSetup.PaperHeightTwips);
        Assert.Equal(9, roundTrip.PageSetup.PrinterPaperSize);
        Assert.Equal(7, roundTrip.PageSetup.FirstPagePaperSource);
        Assert.Equal(8, roundTrip.PageSetup.OtherPagesPaperSource);
        Assert.Equal(1440, roundTrip.PageSetup.MarginLeftTwips);
        Assert.Equal(720, roundTrip.PageSetup.MarginRightTwips);
        Assert.Equal(360, roundTrip.PageSetup.HeaderDistanceTwips);
        Assert.Equal(540, roundTrip.PageSetup.FooterDistanceTwips);
        Assert.True(roundTrip.PageSetup.RtlGutter);
        Assert.Equal(5, roundTrip.PageSetup.PageNumberStart);
        Assert.True(roundTrip.PageSetup.PageNumberRestart);
        Assert.Equal(720, roundTrip.PageSetup.PageNumberPositionXTwips);
        Assert.Equal(900, roundTrip.PageSetup.PageNumberPositionYTwips);
        Assert.Equal(RtfPageNumberFormat.LowerRoman, roundTrip.PageSetup.PageNumberFormat);
        Assert.True(roundTrip.PageSetup.PageBorders.IncludeHeader);
        Assert.True(roundTrip.PageSetup.PageBorders.IncludeFooter);
        Assert.True(roundTrip.PageSetup.PageBorders.SnapToPageBorder);
        Assert.Equal(RtfPageBorderScope.AllExceptFirstPageInSection, roundTrip.PageSetup.PageBorders.Scope);
        Assert.False(roundTrip.PageSetup.PageBorders.DisplayBehindText);
        Assert.Equal(RtfPageBorderOffset.PageEdge, roundTrip.PageSetup.PageBorders.OffsetFrom);
        Assert.Equal(RtfPageBorderStyle.Single, roundTrip.PageSetup.PageBorders.Top.Style);
        Assert.Equal(1, roundTrip.PageSetup.PageBorders.Top.ColorIndex);
        Assert.Equal(RtfPageBorderStyle.Double, roundTrip.PageSetup.PageBorders.Bottom.Style);
        Assert.Equal(2, roundTrip.PageSetup.PageBorders.Bottom.ColorIndex);
        Assert.True(roundTrip.PageSetup.Landscape);
        Assert.True(roundTrip.PageSetup.DifferentFirstPageHeaderFooter);
        Assert.Equal(5, roundTrip.NoteSettings.FootnoteStartNumber);
        Assert.Equal(RtfNoteNumberRestart.EachSection, roundTrip.NoteSettings.FootnoteRestart);
        Assert.Equal(RtfNoteNumberFormat.LowerRoman, roundTrip.NoteSettings.FootnoteNumberFormat);
        Assert.Equal(RtfFootnotePlacement.BeneathText, roundTrip.NoteSettings.FootnotePlacement);
        Assert.Equal(7, roundTrip.NoteSettings.EndnoteStartNumber);
        Assert.Equal(RtfNoteNumberRestart.Continuous, roundTrip.NoteSettings.EndnoteRestart);
        Assert.Equal(RtfNoteNumberFormat.Arabic, roundTrip.NoteSettings.EndnoteNumberFormat);
        Assert.Equal(RtfEndnotePlacement.DocumentEnd, roundTrip.NoteSettings.EndnotePlacement);
        Assert.Equal("Landscape body", Assert.Single(roundTrip.Paragraphs).ToPlainText());

        string rtf = roundTrip.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        Assert.Contains(@"\paperw16838\paperh11906", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\pgbrdrt\brdrs\brdrw12\brsp24\brdrcf1", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\ftnstart5\ftnrestart\ftnnrlc\ftntj", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void RtfDocument_ToHtml_RoundTrips_Sections_Layout_And_Block_Ownership() {
        RtfDocument document = RtfDocument.Create();
        int green = document.AddColor(0, 128, 0);
        RtfSection first = document.AddSection(RtfSectionBreakKind.OddPage);
        first.PageSetup.SetPaperSize(10000, 12000)
            .SetMargins(leftTwips: 720)
            .SetGutter(240, rtlGutter: true)
            .SetHeaderFooterDistance(headerDistanceTwips: 300, footerDistanceTwips: 420)
            .SetPageNumbering(start: 3, restart: true, format: RtfPageNumberFormat.UpperRoman)
            .SetLandscape()
            .SetDifferentFirstPageHeaderFooter();
        first.PageSetup.PageBorders.Left.Set(RtfPageBorderStyle.Dotted, width: 8, space: 12, colorIndex: green);
        first.ColumnCount = 2;
        first.ColumnSpaceTwips = 720;
        first.ColumnSeparator = true;
        first.Direction = RtfTextDirection.RightToLeft;
        first.AddColumn(widthTwips: 3000, spaceAfterTwips: 360);
        first.AddColumn(widthTwips: 4000, spaceAfterTwips: 0);
        first.SetVerticalAlignment(RtfSectionVerticalAlignment.Center);
        first.LineNumbering.Set(countBy: 2, distanceFromTextTwips: 360, startNumber: 10, restart: RtfLineNumberRestart.EachPage);
        first.NoteSettings.SetFootnoteNumbering(start: 3, restart: RtfNoteNumberRestart.EachPage, format: RtfNoteNumberFormat.UpperLetter);
        first.AddParagraph("First");
        RtfTable table = first.AddTable(1, 1);
        table.Rows[0].Cells[0].AddParagraph("Section table");

        RtfSection second = document.AddSection(RtfSectionBreakKind.Continuous);
        second.ColumnCount = 1;
        second.VerticalAlignment = RtfSectionVerticalAlignment.Bottom;
        second.LineNumbering.Set(countBy: 0, restart: RtfLineNumberRestart.Continuous);
        second.AddParagraph("Second");

        string html = document.ToHtml(new RtfHtmlSaveOptions { FragmentOnly = true, NewLine = "\n" });

        Assert.Contains("data-officeimo-rtf-section=\"true\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-rtf-section-layout=\"", html, StringComparison.Ordinal);
        Assert.Contains("<table>", html, StringComparison.Ordinal);

        RtfDocument roundTrip = html.LoadFromHtml();
        Assert.Equal(2, roundTrip.Sections.Count);
        Assert.Equal(2, roundTrip.Paragraphs.Count);
        RtfSection roundTripFirst = roundTrip.Sections[0];
        Assert.Equal(RtfSectionBreakKind.OddPage, roundTripFirst.BreakKind);
        Assert.Equal(2, roundTripFirst.Blocks.Count);
        Assert.Equal("First", Assert.IsType<RtfParagraph>(roundTripFirst.Blocks[0]).ToPlainText());
        Assert.IsType<RtfTable>(roundTripFirst.Blocks[1]);
        Assert.Equal(2, roundTripFirst.ColumnCount);
        Assert.Equal(720, roundTripFirst.ColumnSpaceTwips);
        Assert.True(roundTripFirst.ColumnSeparator);
        Assert.Equal(2, roundTripFirst.Columns.Count);
        Assert.Equal(3000, roundTripFirst.Columns[0].WidthTwips);
        Assert.Equal(360, roundTripFirst.Columns[0].SpaceAfterTwips);
        Assert.Equal(4000, roundTripFirst.Columns[1].WidthTwips);
        Assert.Equal(0, roundTripFirst.Columns[1].SpaceAfterTwips);
        Assert.Equal(RtfTextDirection.RightToLeft, roundTripFirst.Direction);
        Assert.Equal(RtfSectionVerticalAlignment.Center, roundTripFirst.VerticalAlignment);
        Assert.Equal(10000, roundTripFirst.PageSetup.PaperWidthTwips);
        Assert.Equal(12000, roundTripFirst.PageSetup.PaperHeightTwips);
        Assert.Equal(240, roundTripFirst.PageSetup.GutterWidthTwips);
        Assert.True(roundTripFirst.PageSetup.RtlGutter);
        Assert.Equal(RtfPageNumberFormat.UpperRoman, roundTripFirst.PageSetup.PageNumberFormat);
        Assert.Equal(RtfPageBorderStyle.Dotted, roundTripFirst.PageSetup.PageBorders.Left.Style);
        Assert.Equal(1, roundTripFirst.PageSetup.PageBorders.Left.ColorIndex);
        Assert.Equal(2, roundTripFirst.LineNumbering.CountBy);
        Assert.Equal(360, roundTripFirst.LineNumbering.DistanceFromTextTwips);
        Assert.Equal(10, roundTripFirst.LineNumbering.StartNumber);
        Assert.Equal(RtfLineNumberRestart.EachPage, roundTripFirst.LineNumbering.Restart);
        Assert.Equal(3, roundTripFirst.NoteSettings.FootnoteStartNumber);
        Assert.Equal(RtfNoteNumberFormat.UpperLetter, roundTripFirst.NoteSettings.FootnoteNumberFormat);

        RtfSection roundTripSecond = roundTrip.Sections[1];
        Assert.Equal(RtfSectionBreakKind.Continuous, roundTripSecond.BreakKind);
        Assert.Equal("Second", Assert.IsType<RtfParagraph>(Assert.Single(roundTripSecond.Blocks)).ToPlainText());
        Assert.Equal(RtfSectionVerticalAlignment.Bottom, roundTripSecond.VerticalAlignment);
        Assert.Equal(0, roundTripSecond.LineNumbering.CountBy);
        Assert.Equal(RtfLineNumberRestart.Continuous, roundTripSecond.LineNumbering.Restart);

        string rtf = roundTrip.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        Assert.Contains(@"\sectd\sbkodd", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\cols2\colsx720", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\colno1\colw3000\colsr360", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\sectd\sbknone", rtf, StringComparison.Ordinal);
    }
}
