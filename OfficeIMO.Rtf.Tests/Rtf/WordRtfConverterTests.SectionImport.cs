using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using System.Linq;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public partial class WordRtfConverterTests {
    [Fact]
    public void Rtf_To_Word_Bridge_Applies_PageSetup_Headers_And_Footers() {
        RtfDocument rtfDocument = RtfDocument.Create();
        rtfDocument.PageSetup.SetPaperSize(16838, 11906)
            .SetMargins(leftTwips: 1440, rightTwips: 720, topTwips: 1080, bottomTwips: 1080)
            .SetGutter(180, rtlGutter: true)
            .SetHeaderFooterDistance(headerDistanceTwips: 360, footerDistanceTwips: 540)
            .SetPageNumbering(start: 5, restart: true, format: RtfPageNumberFormat.LowerRoman)
            .SetLandscape()
            .SetDifferentFirstPageHeaderFooter();
        int red = rtfDocument.AddColor(255, 0, 0);
        int blue = rtfDocument.AddColor(0, 0, 255);
        rtfDocument.PageSetup.PageBorders.SetDisplayOptions(RtfPageBorderScope.AllExceptFirstPageInSection, displayBehindText: false, RtfPageBorderOffset.PageEdge);
        rtfDocument.PageSetup.PageBorders.Top.Set(RtfPageBorderStyle.Single, width: 12, space: 24, colorIndex: red);
        rtfDocument.PageSetup.PageBorders.Bottom.Set(RtfPageBorderStyle.Double, width: 18, space: 30, colorIndex: blue);
        rtfDocument.NoteSettings
            .SetFootnoteNumbering(start: 5, restart: RtfNoteNumberRestart.EachSection, format: RtfNoteNumberFormat.LowerRoman)
            .SetFootnotePlacement(RtfFootnotePlacement.BeneathText)
            .SetEndnoteNumbering(start: 7, restart: RtfNoteNumberRestart.Continuous, format: RtfNoteNumberFormat.Arabic)
            .SetEndnotePlacement(RtfEndnotePlacement.DocumentEnd);
        rtfDocument.Settings.FacingPages = true;
        rtfDocument.Settings.MirrorMargins = true;
        rtfDocument.AddHeader().AddParagraph("Header ").AddText("bold").SetBold();
        rtfDocument.AddHeader(RtfHeaderFooterKind.FirstHeader).AddParagraph("First header");
        rtfDocument.AddHeader(RtfHeaderFooterKind.LeftHeader).AddParagraph("Even header");
        rtfDocument.AddFooter().AddParagraph("Footer text");
        rtfDocument.AddFooter(RtfHeaderFooterKind.LeftFooter).AddParagraph("Even footer");
        rtfDocument.AddParagraph("Body");

        using WordDocument word = rtfDocument.ToWordDocument();

        Assert.Equal(PageOrientationValues.Landscape, word.PageOrientation);
        Assert.Equal(16838U, word.PageSettings.Width?.Value);
        Assert.Equal(11906U, word.PageSettings.Height?.Value);
        Assert.Equal(1440U, word.Margins.Left.Value);
        Assert.Equal(720U, word.Margins.Right.Value);
        Assert.Equal(1080, word.Margins.Top);
        Assert.Equal(1080, word.Margins.Bottom);
        Assert.Equal(180U, word.Margins.Gutter.Value);
        Assert.Equal(360U, word.Margins.HeaderDistance.Value);
        Assert.Equal(540U, word.Margins.FooterDistance.Value);
        Assert.True(word.RtlGutter);
        Assert.Equal(5, word.PageNumberType.Start?.Value);
        Assert.Equal(NumberFormatValues.LowerRoman, word.PageNumberType.Format?.Value);
        Assert.Equal(BorderValues.Single, word.Borders.TopStyle);
        Assert.Equal(12U, word.Borders.TopSize?.Value);
        Assert.Equal(24U, word.Borders.TopSpace?.Value);
        Assert.Equal("FF0000", word.Borders.TopColorHex);
        Assert.Equal(BorderValues.Double, word.Borders.BottomStyle);
        Assert.Equal(18U, word.Borders.BottomSize?.Value);
        Assert.Equal(30U, word.Borders.BottomSpace?.Value);
        Assert.Equal("0000FF", word.Borders.BottomColorHex);
        PageBorders appliedPageBorders = word.Sections[0]._sectionProperties.GetFirstChild<PageBorders>()!;
        Assert.Equal(PageBorderDisplayValues.NotFirstPage, appliedPageBorders.Display?.Value);
        Assert.Equal(PageBorderZOrderValues.Front, appliedPageBorders.ZOrder?.Value);
        Assert.Equal(PageBorderOffsetValues.Page, appliedPageBorders.OffsetFrom?.Value);
        Assert.Equal(5, (int?)word.FootnoteProperties.NumberingStart?.Val?.Value);
        Assert.Equal(RestartNumberValues.EachSection, word.FootnoteProperties.NumberingRestart?.Val?.Value);
        Assert.Equal(NumberFormatValues.LowerRoman, word.FootnoteProperties.NumberingFormat?.Val?.Value);
        Assert.Equal(FootnotePositionValues.BeneathText, word.FootnoteProperties.FootnotePosition?.Val?.Value);
        Assert.Equal(7, (int?)word.EndnoteProperties.NumberingStart?.Val?.Value);
        Assert.Equal(RestartNumberValues.Continuous, word.EndnoteProperties.NumberingRestart?.Val?.Value);
        Assert.Equal(NumberFormatValues.Decimal, word.EndnoteProperties.NumberingFormat?.Val?.Value);
        Assert.Equal(EndnotePositionValues.DocumentEnd, word.EndnoteProperties.EndnotePosition?.Val?.Value);
        Assert.True(word.DifferentFirstPage);
        Assert.True(word.DifferentOddAndEvenPages);
        Assert.True(word.Settings.MirrorMargins);
        Assert.NotNull(word.Header?.Default);
        Assert.NotNull(word.Header?.First);
        Assert.NotNull(word.Header?.Even);
        Assert.NotNull(word.Footer?.Default);
        Assert.NotNull(word.Footer?.Even);
        Assert.Equal("Header bold", string.Concat(word.Header!.Default!.Paragraphs.Select(paragraph => paragraph.Text)));
        Assert.Equal("First header", string.Concat(word.Header!.First!.Paragraphs.Select(paragraph => paragraph.Text)));
        Assert.Equal("Even header", string.Concat(word.Header!.Even!.Paragraphs.Select(paragraph => paragraph.Text)));
        Assert.Contains(word.Header.Default.Paragraphs, paragraph => paragraph.Text == "bold" && paragraph.Bold);
        Assert.Equal("Footer text", string.Concat(word.Footer!.Default!.Paragraphs.Select(paragraph => paragraph.Text)));
        Assert.Equal("Even footer", string.Concat(word.Footer!.Even!.Paragraphs.Select(paragraph => paragraph.Text)));
    }

    [Fact]
    public void Word_To_Rtf_Bridge_Preserves_Sections_Columns_And_Section_PageSetup() {
        using WordDocument word = WordDocument.Create();
        word.Sections[0].AddParagraph("First section");
        WordSection second = word.AddSection(SectionMarkValues.Continuous);
        second.AddParagraph("Second section");
        WordTable secondTable = second.AddTable(1, 2);
        secondTable.Rows[0].Cells[0].AddParagraph("S2A", removeExistingParagraphs: true);
        secondTable.Rows[0].Cells[1].AddParagraph("S2B", removeExistingParagraphs: true);

        WordSection first = word.Sections[0];
        first.PageOrientation = PageOrientationValues.Landscape;
        first.PageSettings.Width = 16838U;
        first.PageSettings.Height = 11906U;
        first.Margins.Left = 1440U;
        first.Margins.Right = 720U;
        first.Margins.Gutter = 240U;
        first.Margins.HeaderDistance = 300U;
        first.Margins.FooterDistance = 420U;
        first.RtlGutter = true;
        first.AddPageNumbering(3, NumberFormatValues.UpperRoman);
        first.Borders.LeftStyle = BorderValues.Dotted;
        first.Borders.LeftSize = 8U;
        first.Borders.LeftSpace = 12U;
        first.Borders.RightStyle = BorderValues.Dashed;
        first.Borders.RightSize = 10U;
        first.Borders.RightSpace = 14U;
        PageBorders sectionPageBorders = first._sectionProperties.GetFirstChild<PageBorders>()!;
        sectionPageBorders.Display = PageBorderDisplayValues.FirstPage;
        sectionPageBorders.ZOrder = PageBorderZOrderValues.Back;
        sectionPageBorders.OffsetFrom = PageBorderOffsetValues.Page;
        first.AddFootnoteProperties(NumberFormatValues.UpperLetter, FootnotePositionValues.PageBottom, RestartNumberValues.EachPage, startNumber: 3);
        first.AddEndnoteProperties(NumberFormatValues.LowerLetter, EndnotePositionValues.SectionEnd, RestartNumberValues.EachSection, startNumber: 9);
        first.DifferentFirstPage = true;
        first._sectionProperties.Append(new VerticalTextAlignmentOnPage { Val = VerticalJustificationValues.Center });
        first.ColumnCount = 2;
        first.ColumnsSpace = 720;
        first.HasColumnSeparator = true;
        Columns firstColumns = first._sectionProperties.GetFirstChild<Columns>()!;
        firstColumns.EqualWidth = false;
        firstColumns.RemoveAllChildren<Column>();
        firstColumns.Append(new Column { Width = "3000", Space = "360" });
        firstColumns.Append(new Column { Width = "4000", Space = "0" });
        first._sectionProperties.Append(new LineNumberType {
            CountBy = 2,
            Distance = "360",
            Start = 10,
            Restart = LineNumberRestartValues.NewPage
        });
        second.ColumnCount = 1;

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfWriteOptions { IncludeGenerator = false });

        Assert.Equal(2, rtfDocument.Sections.Count);
        Assert.Equal(RtfSectionBreakKind.Continuous, rtfDocument.Sections[1].BreakKind);
        Assert.Equal(2, rtfDocument.Sections[0].ColumnCount);
        Assert.Equal(720, rtfDocument.Sections[0].ColumnSpaceTwips);
        Assert.True(rtfDocument.Sections[0].ColumnSeparator);
        Assert.Collection(rtfDocument.Sections[0].Columns,
            column => {
                Assert.Equal(3000, column.WidthTwips);
                Assert.Equal(360, column.SpaceAfterTwips);
            },
            column => {
                Assert.Equal(4000, column.WidthTwips);
                Assert.Equal(0, column.SpaceAfterTwips);
            });
        Assert.Equal(16838, rtfDocument.Sections[0].PageSetup.PaperWidthTwips);
        Assert.True(rtfDocument.Sections[0].PageSetup.Landscape);
        Assert.Equal(240, rtfDocument.Sections[0].PageSetup.GutterWidthTwips);
        Assert.Equal(300, rtfDocument.Sections[0].PageSetup.HeaderDistanceTwips);
        Assert.Equal(420, rtfDocument.Sections[0].PageSetup.FooterDistanceTwips);
        Assert.True(rtfDocument.Sections[0].PageSetup.RtlGutter);
        Assert.Equal(3, rtfDocument.Sections[0].PageSetup.PageNumberStart);
        Assert.True(rtfDocument.Sections[0].PageSetup.PageNumberRestart);
        Assert.Equal(RtfPageNumberFormat.UpperRoman, rtfDocument.Sections[0].PageSetup.PageNumberFormat);
        Assert.Equal(RtfPageBorderScope.FirstPageInSection, rtfDocument.Sections[0].PageSetup.PageBorders.Scope);
        Assert.True(rtfDocument.Sections[0].PageSetup.PageBorders.DisplayBehindText);
        Assert.Equal(RtfPageBorderOffset.PageEdge, rtfDocument.Sections[0].PageSetup.PageBorders.OffsetFrom);
        Assert.Equal(RtfPageBorderStyle.Dotted, rtfDocument.Sections[0].PageSetup.PageBorders.Left.Style);
        Assert.Equal(8, rtfDocument.Sections[0].PageSetup.PageBorders.Left.Width);
        Assert.Equal(12, rtfDocument.Sections[0].PageSetup.PageBorders.Left.Space);
        Assert.Equal(RtfPageBorderStyle.Dashed, rtfDocument.Sections[0].PageSetup.PageBorders.Right.Style);
        Assert.Equal(10, rtfDocument.Sections[0].PageSetup.PageBorders.Right.Width);
        Assert.Equal(14, rtfDocument.Sections[0].PageSetup.PageBorders.Right.Space);
        Assert.Equal(3, rtfDocument.Sections[0].NoteSettings.FootnoteStartNumber);
        Assert.Equal(RtfNoteNumberRestart.EachPage, rtfDocument.Sections[0].NoteSettings.FootnoteRestart);
        Assert.Equal(RtfNoteNumberFormat.UpperLetter, rtfDocument.Sections[0].NoteSettings.FootnoteNumberFormat);
        Assert.Equal(RtfFootnotePlacement.PageBottom, rtfDocument.Sections[0].NoteSettings.FootnotePlacement);
        Assert.Equal(9, rtfDocument.Sections[0].NoteSettings.EndnoteStartNumber);
        Assert.Equal(RtfNoteNumberRestart.EachSection, rtfDocument.Sections[0].NoteSettings.EndnoteRestart);
        Assert.Equal(RtfNoteNumberFormat.LowerLetter, rtfDocument.Sections[0].NoteSettings.EndnoteNumberFormat);
        Assert.Equal(RtfEndnotePlacement.SectionEnd, rtfDocument.Sections[0].NoteSettings.EndnotePlacement);
        Assert.Equal(2, rtfDocument.Sections[0].LineNumbering.CountBy);
        Assert.Equal(360, rtfDocument.Sections[0].LineNumbering.DistanceFromTextTwips);
        Assert.Equal(10, rtfDocument.Sections[0].LineNumbering.StartNumber);
        Assert.Equal(RtfLineNumberRestart.EachPage, rtfDocument.Sections[0].LineNumbering.Restart);
        Assert.Equal(RtfSectionVerticalAlignment.Center, rtfDocument.Sections[0].VerticalAlignment);
        Assert.True(rtfDocument.Sections[0].PageSetup.DifferentFirstPageHeaderFooter);
        RtfTable sectionTable = Assert.IsType<RtfTable>(rtfDocument.Sections[1].Blocks[1]);
        Assert.Equal("S2A", sectionTable.Rows[0].Cells[0].Paragraphs[0].ToPlainText());
        Assert.Equal("S2B", sectionTable.Rows[0].Cells[1].Paragraphs[0].ToPlainText());
        Assert.Contains(@"\sectd", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\cols2", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\colno1\colw3000\colsr360\colno2\colw4000\colsr0", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\titlepg", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\guttersxn240", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\headery300", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\footery420", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\rtlgutter", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\pgnstarts3", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\pgnucrm", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\pgbrdropt41", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\pgbrdrl\brdrdot\brdrw8\brsp12", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\pgbrdrr\brdrdash\brdrw10\brsp14", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\ftnstart3", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\ftnrstpg", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\ftnnauc", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\ftnbj", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\aftnstart9", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\aftnrestart", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\aftnnalc", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\aendnotes", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\linemod2", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\linex360", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\linestarts10", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\lineppage", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\vertalc", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\linebetcol", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\sbknone", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void Word_To_Rtf_Bridge_Exports_Single_Section_When_Only_Vertical_Alignment_Is_Set() {
        using WordDocument word = WordDocument.Create();
        word.Sections[0]._sectionProperties.Append(new VerticalTextAlignmentOnPage { Val = VerticalJustificationValues.Bottom });
        word.Sections[0].AddParagraph("Single section");

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfWriteOptions { IncludeGenerator = false });

        RtfSection section = Assert.Single(rtfDocument.Sections);
        Assert.Equal(RtfSectionVerticalAlignment.Bottom, section.VerticalAlignment);
        Assert.Contains(@"\sectd", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\vertalb", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void Rtf_To_Word_Bridge_Preserves_Sections_Columns_And_Section_PageSetup() {
        RtfDocument rtfDocument = RtfDocument.Create();
        RtfSection first = rtfDocument.AddSection();
        first.PageSetup.SetPaperSize(16838, 11906)
            .SetMargins(leftTwips: 1440, rightTwips: 720, topTwips: 1080, bottomTwips: 1080)
            .SetGutter(240, rtlGutter: true)
            .SetHeaderFooterDistance(headerDistanceTwips: 300, footerDistanceTwips: 420)
            .SetPageNumbering(start: 3, restart: true, format: RtfPageNumberFormat.UpperRoman)
            .SetLandscape()
            .SetDifferentFirstPageHeaderFooter();
        first.PageSetup.PageBorders.SetDisplayOptions(RtfPageBorderScope.FirstPageInSection, displayBehindText: true, RtfPageBorderOffset.PageEdge);
        first.PageSetup.PageBorders.Left.Set(RtfPageBorderStyle.Dotted, width: 8, space: 12);
        first.PageSetup.PageBorders.Right.Set(RtfPageBorderStyle.Dashed, width: 10, space: 14);
        first.NoteSettings
            .SetFootnoteNumbering(start: 3, restart: RtfNoteNumberRestart.EachPage, format: RtfNoteNumberFormat.UpperLetter)
            .SetFootnotePlacement(RtfFootnotePlacement.PageBottom)
            .SetEndnoteNumbering(start: 9, restart: RtfNoteNumberRestart.EachSection, format: RtfNoteNumberFormat.LowerLetter)
            .SetEndnotePlacement(RtfEndnotePlacement.SectionEnd);
        first.LineNumbering.Set(countBy: 2, distanceFromTextTwips: 360, startNumber: 10, restart: RtfLineNumberRestart.EachPage);
        first.VerticalAlignment = RtfSectionVerticalAlignment.Justified;
        first.ColumnCount = 2;
        first.ColumnSpaceTwips = 720;
        first.ColumnSeparator = true;
        first.AddColumn(widthTwips: 3000, spaceAfterTwips: 360);
        first.AddColumn(widthTwips: 4000, spaceAfterTwips: 0);
        first.AddParagraph("First section");
        RtfSection second = rtfDocument.AddSection(RtfSectionBreakKind.Continuous);
        second.ColumnCount = 1;
        second.AddParagraph("Second section");
        RtfTable sectionTable = second.AddTable(0, 1);
        RtfTableRow sectionRow = sectionTable.AddRow();
        sectionRow.AddCell(2400).AddParagraph("S2A");
        sectionRow.AddCell(4800).AddParagraph("S2B");

        using WordDocument word = rtfDocument.ToWordDocument();

        Assert.Equal(2, word.Sections.Count);
        Assert.True(word.Sections[0].PageOrientation == PageOrientationValues.Landscape);
        Assert.Equal(16838U, word.Sections[0].PageSettings.Width?.Value);
        Assert.Equal(1440U, word.Sections[0].Margins.Left.Value);
        Assert.Equal(240U, word.Sections[0].Margins.Gutter.Value);
        Assert.Equal(300U, word.Sections[0].Margins.HeaderDistance.Value);
        Assert.Equal(420U, word.Sections[0].Margins.FooterDistance.Value);
        Assert.True(word.Sections[0].RtlGutter);
        Assert.Equal(3, word.Sections[0].PageNumberType.Start?.Value);
        Assert.Equal(NumberFormatValues.UpperRoman, word.Sections[0].PageNumberType.Format?.Value);
        Assert.Equal(BorderValues.Dotted, word.Sections[0].Borders.LeftStyle);
        Assert.Equal(8U, word.Sections[0].Borders.LeftSize?.Value);
        Assert.Equal(12U, word.Sections[0].Borders.LeftSpace?.Value);
        Assert.Equal(BorderValues.Dashed, word.Sections[0].Borders.RightStyle);
        Assert.Equal(10U, word.Sections[0].Borders.RightSize?.Value);
        Assert.Equal(14U, word.Sections[0].Borders.RightSpace?.Value);
        PageBorders wordSectionPageBorders = word.Sections[0]._sectionProperties.GetFirstChild<PageBorders>()!;
        Assert.Equal(PageBorderDisplayValues.FirstPage, wordSectionPageBorders.Display?.Value);
        Assert.Equal(PageBorderZOrderValues.Back, wordSectionPageBorders.ZOrder?.Value);
        Assert.Equal(PageBorderOffsetValues.Page, wordSectionPageBorders.OffsetFrom?.Value);
        Assert.Equal(3, (int?)word.Sections[0].FootnoteProperties.NumberingStart?.Val?.Value);
        Assert.Equal(RestartNumberValues.EachPage, word.Sections[0].FootnoteProperties.NumberingRestart?.Val?.Value);
        Assert.Equal(NumberFormatValues.UpperLetter, word.Sections[0].FootnoteProperties.NumberingFormat?.Val?.Value);
        Assert.Equal(FootnotePositionValues.PageBottom, word.Sections[0].FootnoteProperties.FootnotePosition?.Val?.Value);
        Assert.Equal(9, (int?)word.Sections[0].EndnoteProperties.NumberingStart?.Val?.Value);
        Assert.Equal(RestartNumberValues.EachSection, word.Sections[0].EndnoteProperties.NumberingRestart?.Val?.Value);
        Assert.Equal(NumberFormatValues.LowerLetter, word.Sections[0].EndnoteProperties.NumberingFormat?.Val?.Value);
        Assert.Equal(EndnotePositionValues.SectionEnd, word.Sections[0].EndnoteProperties.EndnotePosition?.Val?.Value);
        LineNumberType wordLineNumbering = word.Sections[0]._sectionProperties.GetFirstChild<LineNumberType>()!;
        Assert.NotNull(wordLineNumbering);
        Assert.Equal(2, (int?)wordLineNumbering.CountBy?.Value);
        Assert.Equal("360", wordLineNumbering.Distance?.Value);
        Assert.Equal(10, (int?)wordLineNumbering.Start?.Value);
        Assert.Equal(LineNumberRestartValues.NewPage, wordLineNumbering.Restart?.Value);
        VerticalTextAlignmentOnPage wordVerticalAlignment = word.Sections[0]._sectionProperties.GetFirstChild<VerticalTextAlignmentOnPage>()!;
        Assert.NotNull(wordVerticalAlignment);
        Assert.Equal(VerticalJustificationValues.Both, wordVerticalAlignment.Val?.Value);
        Assert.True(word.Sections[0].DifferentFirstPage);
        Assert.Equal(2, word.Sections[0].ColumnCount);
        Assert.Equal(720, word.Sections[0].ColumnsSpace);
        Assert.True(word.Sections[0].HasColumnSeparator);
        Columns wordColumns = word.Sections[0]._sectionProperties.GetFirstChild<Columns>()!;
        Assert.False(wordColumns.EqualWidth?.Value ?? true);
        Assert.Collection(wordColumns.Elements<Column>(),
            column => {
                Assert.Equal("3000", column.Width?.Value);
                Assert.Equal("360", column.Space?.Value);
            },
            column => {
                Assert.Equal("4000", column.Width?.Value);
                Assert.Equal("0", column.Space?.Value);
            });
        Assert.Equal(1, word.Sections[1].ColumnCount);
        Assert.Equal("First section", string.Concat(word.Sections[0].Paragraphs.Select(paragraph => paragraph.Text)));
        Assert.Equal("Second section", string.Concat(word.Sections[1].Paragraphs.Select(paragraph => paragraph.Text)));
        WordTable wordSectionTable = Assert.Single(word.Sections[1].Tables);
        Assert.Equal("S2A", GetCellText(wordSectionTable, 0, 0));
        Assert.Equal("S2B", GetCellText(wordSectionTable, 0, 1));
    }
}
