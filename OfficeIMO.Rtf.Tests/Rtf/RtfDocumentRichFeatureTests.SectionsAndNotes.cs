using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Diagnostics;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public partial class RtfDocumentRichFeatureTests {
    [Fact]
    public void Write_And_Read_Page_Setup() {
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
        document.AddParagraph("Landscape body");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.StartsWith(@"{\rtf1\ansi\deff0\paperw16838\paperh11906\psz9\binfsxn7\binsxn8\margl1440\margr720\margt1080\margb1080\gutter180\headery360\footery540\rtlgutter\pgnstarts5\pgnrestart\pgnx720\pgny900\pgnlcrm\pgbrdrhead\pgbrdrfoot\pgbrdropt34\pgbrdrsnap\pgbrdrt\brdrs\brdrw12\brsp24\brdrcf1\pgbrdrb\brdrdb\brdrw18\brsp30\brdrcf2\landscape\titlepg", rtf, StringComparison.Ordinal);
        Assert.Equal(16838, read.Document.PageSetup.PaperWidthTwips);
        Assert.Equal(11906, read.Document.PageSetup.PaperHeightTwips);
        Assert.Equal(9, read.Document.PageSetup.PrinterPaperSize);
        Assert.Equal(7, read.Document.PageSetup.FirstPagePaperSource);
        Assert.Equal(8, read.Document.PageSetup.OtherPagesPaperSource);
        Assert.Equal(1440, read.Document.PageSetup.MarginLeftTwips);
        Assert.Equal(720, read.Document.PageSetup.MarginRightTwips);
        Assert.Equal(1080, read.Document.PageSetup.MarginTopTwips);
        Assert.Equal(1080, read.Document.PageSetup.MarginBottomTwips);
        Assert.Equal(180, read.Document.PageSetup.GutterWidthTwips);
        Assert.Equal(360, read.Document.PageSetup.HeaderDistanceTwips);
        Assert.Equal(540, read.Document.PageSetup.FooterDistanceTwips);
        Assert.True(read.Document.PageSetup.RtlGutter);
        Assert.Equal(5, read.Document.PageSetup.PageNumberStart);
        Assert.True(read.Document.PageSetup.PageNumberRestart);
        Assert.Equal(720, read.Document.PageSetup.PageNumberPositionXTwips);
        Assert.Equal(900, read.Document.PageSetup.PageNumberPositionYTwips);
        Assert.Equal(RtfPageNumberFormat.LowerRoman, read.Document.PageSetup.PageNumberFormat);
        Assert.True(read.Document.PageSetup.PageBorders.IncludeHeader);
        Assert.True(read.Document.PageSetup.PageBorders.IncludeFooter);
        Assert.True(read.Document.PageSetup.PageBorders.SnapToPageBorder);
        Assert.Equal(RtfPageBorderScope.AllExceptFirstPageInSection, read.Document.PageSetup.PageBorders.Scope);
        Assert.False(read.Document.PageSetup.PageBorders.DisplayBehindText);
        Assert.Equal(RtfPageBorderOffset.PageEdge, read.Document.PageSetup.PageBorders.OffsetFrom);
        Assert.Equal(RtfPageBorderStyle.Single, read.Document.PageSetup.PageBorders.Top.Style);
        Assert.Equal(red, read.Document.PageSetup.PageBorders.Top.ColorIndex);
        Assert.Equal(RtfPageBorderStyle.Double, read.Document.PageSetup.PageBorders.Bottom.Style);
        Assert.Equal(blue, read.Document.PageSetup.PageBorders.Bottom.ColorIndex);
        Assert.True(read.Document.PageSetup.Landscape);
        Assert.True(read.Document.PageSetup.DifferentFirstPageHeaderFooter);
        Assert.Equal("Landscape body", Assert.Single(read.Document.Paragraphs).ToPlainText());
    }

    [Fact]
    public void Write_And_Read_Note_Settings() {
        RtfDocument document = RtfDocument.Create();
        document.NoteSettings
            .SetFootnoteNumbering(start: 5, restart: RtfNoteNumberRestart.EachSection, format: RtfNoteNumberFormat.LowerRoman)
            .SetFootnotePlacement(RtfFootnotePlacement.BeneathText)
            .SetEndnoteNumbering(start: 7, restart: RtfNoteNumberRestart.Continuous, format: RtfNoteNumberFormat.Arabic)
            .SetEndnotePlacement(RtfEndnotePlacement.DocumentEnd);
        document.AddParagraph("Notes body");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.StartsWith(@"{\rtf1\ansi\deff0\ftnstart5\ftnrestart\ftnnrlc\ftntj\aftnstart7\aftnrstcont\aftnnar\aenddoc", rtf, StringComparison.Ordinal);
        Assert.Equal(5, read.Document.NoteSettings.FootnoteStartNumber);
        Assert.Equal(RtfNoteNumberRestart.EachSection, read.Document.NoteSettings.FootnoteRestart);
        Assert.Equal(RtfNoteNumberFormat.LowerRoman, read.Document.NoteSettings.FootnoteNumberFormat);
        Assert.Equal(RtfFootnotePlacement.BeneathText, read.Document.NoteSettings.FootnotePlacement);
        Assert.Equal(7, read.Document.NoteSettings.EndnoteStartNumber);
        Assert.Equal(RtfNoteNumberRestart.Continuous, read.Document.NoteSettings.EndnoteRestart);
        Assert.Equal(RtfNoteNumberFormat.Arabic, read.Document.NoteSettings.EndnoteNumberFormat);
        Assert.Equal(RtfEndnotePlacement.DocumentEnd, read.Document.NoteSettings.EndnotePlacement);
        Assert.Equal("Notes body", Assert.Single(read.Document.Paragraphs).ToPlainText());
    }

    [Theory]
    [InlineData(RtfFootnotePlacement.PageBottom, @"\ftnbj")]
    [InlineData(RtfFootnotePlacement.BeneathText, @"\ftntj")]
    [InlineData(RtfFootnotePlacement.SectionEnd, @"\endnotes")]
    [InlineData(RtfFootnotePlacement.DocumentEnd, @"\enddoc")]
    public void Write_And_Read_Footnote_Placement_Controls(RtfFootnotePlacement placement, string controlWord) {
        RtfDocument document = RtfDocument.Create();
        document.NoteSettings.SetFootnotePlacement(placement);
        document.AddParagraph("Footnote placement");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(controlWord, rtf, StringComparison.Ordinal);
        Assert.Equal(placement, read.Document.NoteSettings.FootnotePlacement);
    }

    [Theory]
    [InlineData(RtfEndnotePlacement.SectionEnd, @"\aendnotes")]
    [InlineData(RtfEndnotePlacement.DocumentEnd, @"\aenddoc")]
    [InlineData(RtfEndnotePlacement.PageBottom, @"\aftnbj")]
    [InlineData(RtfEndnotePlacement.BeneathText, @"\aftntj")]
    public void Write_And_Read_Endnote_Placement_Controls(RtfEndnotePlacement placement, string controlWord) {
        RtfDocument document = RtfDocument.Create();
        document.NoteSettings.SetEndnotePlacement(placement);
        document.AddParagraph("Endnote placement");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(controlWord, rtf, StringComparison.Ordinal);
        Assert.Equal(placement, read.Document.NoteSettings.EndnotePlacement);
    }

    [Fact]
    public void Write_And_Read_Document_Settings() {
        RtfDocument document = RtfDocument.Create();
        document.Settings
            .SetCharacterSet(RtfDocumentCharacterSet.Ansi, 1250)
            .SetUnicodeSkipCount(2)
            .SetDefaultTabWidth(720)
            .SetDefaultLanguage(1045)
            .SetDefaultFarEastLanguage(1041)
            .SetDefaultAlternateLanguage(1033)
            .SetView(kind: 4, scale: 125, zoomKind: 2, backspaceBehavior: 1)
            .SetHyphenation(automatic: false, caps: true, consecutiveLimit: 3, zoneTwips: 360)
            .SetProtection(forms: true, revisions: false, annotations: true, readOnly: false)
            .SetRevisionTracking(enabled: true, displayStyle: 3, barPlacement: 2)
            .SetDrawingGrid(horizontalSpacingTwips: 120, verticalSpacingTwips: 180, horizontalOriginTwips: 720, verticalOriginTwips: 900, horizontalShow: 2, verticalShow: 3, snapToGrid: true, useMargins: false);
        document.Settings.WidowOrphanControl = true;
        document.Settings.FacingPages = true;
        document.Settings.MirrorMargins = true;
        document.Settings.Direction = RtfTextDirection.RightToLeft;
        document.AddParagraph("Settings body");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.StartsWith(@"{\rtf1\ansi\ansicpg1250\deff0\uc2\deftab720\deflang1045\deflangfe1041\adeflang1033\viewkind4\viewscale125\viewzk2\viewbksp1\widowctrl\hyphauto0\hyphcaps\hyphconsec3\hyphhotz360\facingp\margmirror\formprot\revprot0\annotprot\readprot0\revisions\revprop3\revbar2\dghspace120\dgvspace180\dghorigin720\dgvorigin900\dghshow2\dgvshow3\dgsnap\dgmargin0\rtldoc", rtf, StringComparison.Ordinal);
        Assert.Equal(RtfDocumentCharacterSet.Ansi, read.Document.Settings.CharacterSet);
        Assert.Equal(1250, read.Document.Settings.AnsiCodePage);
        Assert.Equal(2, read.Document.Settings.UnicodeSkipCount);
        Assert.Equal(720, read.Document.Settings.DefaultTabWidthTwips);
        Assert.Equal(1045, read.Document.Settings.DefaultLanguageId);
        Assert.Equal(1041, read.Document.Settings.DefaultFarEastLanguageId);
        Assert.Equal(1033, read.Document.Settings.DefaultAlternateLanguageId);
        Assert.Equal(4, read.Document.Settings.ViewKind);
        Assert.Equal(125, read.Document.Settings.ViewScale);
        Assert.Equal(2, read.Document.Settings.ZoomKind);
        Assert.Equal(1, read.Document.Settings.ViewBackspaceBehavior);
        Assert.True(read.Document.Settings.WidowOrphanControl);
        Assert.False(read.Document.Settings.AutoHyphenation);
        Assert.True(read.Document.Settings.HyphenateCaps);
        Assert.Equal(3, read.Document.Settings.ConsecutiveHyphenLimit);
        Assert.Equal(360, read.Document.Settings.HyphenationZoneTwips);
        Assert.True(read.Document.Settings.FacingPages);
        Assert.True(read.Document.Settings.MirrorMargins);
        Assert.True(read.Document.Settings.FormProtection);
        Assert.False(read.Document.Settings.RevisionProtection);
        Assert.True(read.Document.Settings.AnnotationProtection);
        Assert.False(read.Document.Settings.ReadOnlyProtection);
        Assert.True(read.Document.Settings.TrackRevisions);
        Assert.Equal(3, read.Document.Settings.RevisionDisplayStyle);
        Assert.Equal(2, read.Document.Settings.RevisionBarPlacement);
        Assert.Equal(120, read.Document.Settings.DrawingGridHorizontalSpacingTwips);
        Assert.Equal(180, read.Document.Settings.DrawingGridVerticalSpacingTwips);
        Assert.Equal(720, read.Document.Settings.DrawingGridHorizontalOriginTwips);
        Assert.Equal(900, read.Document.Settings.DrawingGridVerticalOriginTwips);
        Assert.Equal(2, read.Document.Settings.DrawingGridHorizontalShow);
        Assert.Equal(3, read.Document.Settings.DrawingGridVerticalShow);
        Assert.True(read.Document.Settings.SnapToDrawingGrid);
        Assert.False(read.Document.Settings.DrawingGridUsesMargins);
        Assert.Equal(RtfTextDirection.RightToLeft, read.Document.Settings.Direction);
        Assert.Equal("Settings body", Assert.Single(read.Document.Paragraphs).ToPlainText());
    }

    [Theory]
    [InlineData(RtfDocumentCharacterSet.Mac, @"{\rtf1\mac\deff0")]
    [InlineData(RtfDocumentCharacterSet.Pc, @"{\rtf1\pc\deff0")]
    [InlineData(RtfDocumentCharacterSet.Pca, @"{\rtf1\pca\deff0")]
    public void Write_And_Read_Document_Character_Set(RtfDocumentCharacterSet characterSet, string expectedPrefix) {
        RtfDocument document = RtfDocument.Create();
        document.Settings.SetCharacterSet(characterSet);
        document.AddParagraph("Body");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.StartsWith(expectedPrefix, rtf, StringComparison.Ordinal);
        Assert.Equal(characterSet, read.Document.Settings.CharacterSet);
        Assert.Null(read.Document.Settings.AnsiCodePage);
        Assert.Equal("Body", Assert.Single(read.Document.Paragraphs).ToPlainText());
    }

    [Fact]
    public void Write_And_Read_Sectioned_Document_With_Columns_And_Page_Setup() {
        RtfDocument document = RtfDocument.Create();
        RtfSection first = document.AddSection(RtfSectionBreakKind.OddPage);
        first.PageSetup.SetPaperSize(10000, 12000)
            .SetPrinterPaper(paperSize: 5, firstPageSource: 6, otherPagesSource: 7)
            .SetMargins(leftTwips: 720)
            .SetGutter(240, rtlGutter: true)
            .SetHeaderFooterDistance(headerDistanceTwips: 300, footerDistanceTwips: 420)
            .SetPageNumbering(start: 3, restart: true, format: RtfPageNumberFormat.UpperRoman)
            .SetLandscape()
            .SetDifferentFirstPageHeaderFooter();
        first.ColumnCount = 2;
        first.ColumnSpaceTwips = 720;
        first.ColumnSeparator = true;
        first.Direction = RtfTextDirection.RightToLeft;
        first.AddColumn(widthTwips: 3000, spaceAfterTwips: 360);
        first.AddColumn(widthTwips: 4000, spaceAfterTwips: 0);
        first.SetVerticalAlignment(RtfSectionVerticalAlignment.Center);
        first.LineNumbering.Set(countBy: 2, distanceFromTextTwips: 360, startNumber: 10, restart: RtfLineNumberRestart.EachPage);
        first.NoteSettings
            .SetFootnoteNumbering(start: 3, restart: RtfNoteNumberRestart.EachPage, format: RtfNoteNumberFormat.UpperLetter)
            .SetFootnotePlacement(RtfFootnotePlacement.PageBottom)
            .SetEndnoteNumbering(start: 9, restart: RtfNoteNumberRestart.EachSection, format: RtfNoteNumberFormat.LowerLetter)
            .SetEndnotePlacement(RtfEndnotePlacement.SectionEnd);
        first.PageSetup.PageBorders.SetDisplayOptions(RtfPageBorderScope.FirstPageInSection, displayBehindText: true, RtfPageBorderOffset.PageEdge);
        first.PageSetup.PageBorders.Left.Set(RtfPageBorderStyle.Dotted, width: 8, space: 12);
        first.PageSetup.PageBorders.Right.Set(RtfPageBorderStyle.Dashed, width: 10, space: 14);
        first.AddParagraph("First");
        RtfSection second = document.AddSection(RtfSectionBreakKind.Continuous);
        second.ColumnCount = 1;
        second.VerticalAlignment = RtfSectionVerticalAlignment.Bottom;
        second.LineNumbering.Set(countBy: 0, restart: RtfLineNumberRestart.Continuous);
        second.AddParagraph("Second");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"\sectd\sbkodd\pgwsxn10000\pghsxn12000\psz5\binfsxn6\binsxn7\marglsxn720\guttersxn240\headery300\footery420\rtlgutter\pgnstarts3\pgnrestart\pgnucrm\pgbrdropt41\pgbrdrl\brdrdot\brdrw8\brsp12\pgbrdrr\brdrdash\brdrw10\brsp14\lndscpsxn\titlepg\vertalc\rtlsect\ftnstart3\ftnrstpg\ftnnauc\ftnbj\aftnstart9\aftnrestart\aftnnalc\aendnotes\linemod2\linex360\linestarts10\lineppage\cols2\colsx720\colno1\colw3000\colsr360\colno2\colw4000\colsr0\linebetcol", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\sectd\sbknone\vertalb\linemod0\linecont\cols1", rtf, StringComparison.Ordinal);
        Assert.Equal(2, document.Paragraphs.Count);
        Assert.Equal(2, read.Document.Sections.Count);
        Assert.Equal("First", Assert.IsType<RtfParagraph>(Assert.Single(read.Document.Sections[0].Blocks)).ToPlainText());
        Assert.Equal(RtfSectionBreakKind.OddPage, read.Document.Sections[0].BreakKind);
        Assert.Equal(2, read.Document.Sections[0].ColumnCount);
        Assert.Equal(2, read.Document.Sections[0].Columns.Count);
        Assert.Equal(3000, read.Document.Sections[0].Columns[0].WidthTwips);
        Assert.Equal(360, read.Document.Sections[0].Columns[0].SpaceAfterTwips);
        Assert.Equal(4000, read.Document.Sections[0].Columns[1].WidthTwips);
        Assert.Equal(0, read.Document.Sections[0].Columns[1].SpaceAfterTwips);
        Assert.True(read.Document.Sections[0].ColumnSeparator);
        Assert.Equal(2, read.Document.Sections[0].LineNumbering.CountBy);
        Assert.Equal(360, read.Document.Sections[0].LineNumbering.DistanceFromTextTwips);
        Assert.Equal(10, read.Document.Sections[0].LineNumbering.StartNumber);
        Assert.Equal(RtfLineNumberRestart.EachPage, read.Document.Sections[0].LineNumbering.Restart);
        Assert.Equal(RtfSectionVerticalAlignment.Center, read.Document.Sections[0].VerticalAlignment);
        Assert.Equal(RtfTextDirection.RightToLeft, read.Document.Sections[0].Direction);
        Assert.Equal(5, read.Document.Sections[0].PageSetup.PrinterPaperSize);
        Assert.Equal(6, read.Document.Sections[0].PageSetup.FirstPagePaperSource);
        Assert.Equal(7, read.Document.Sections[0].PageSetup.OtherPagesPaperSource);
        Assert.Equal(240, read.Document.Sections[0].PageSetup.GutterWidthTwips);
        Assert.Equal(300, read.Document.Sections[0].PageSetup.HeaderDistanceTwips);
        Assert.Equal(420, read.Document.Sections[0].PageSetup.FooterDistanceTwips);
        Assert.True(read.Document.Sections[0].PageSetup.RtlGutter);
        Assert.Equal(3, read.Document.Sections[0].PageSetup.PageNumberStart);
        Assert.True(read.Document.Sections[0].PageSetup.PageNumberRestart);
        Assert.Equal(RtfPageNumberFormat.UpperRoman, read.Document.Sections[0].PageSetup.PageNumberFormat);
        Assert.True(read.Document.Sections[0].PageSetup.DifferentFirstPageHeaderFooter);
        Assert.Equal(3, read.Document.Sections[0].NoteSettings.FootnoteStartNumber);
        Assert.Equal(RtfNoteNumberRestart.EachPage, read.Document.Sections[0].NoteSettings.FootnoteRestart);
        Assert.Equal(RtfNoteNumberFormat.UpperLetter, read.Document.Sections[0].NoteSettings.FootnoteNumberFormat);
        Assert.Equal(RtfFootnotePlacement.PageBottom, read.Document.Sections[0].NoteSettings.FootnotePlacement);
        Assert.Equal(9, read.Document.Sections[0].NoteSettings.EndnoteStartNumber);
        Assert.Equal(RtfNoteNumberRestart.EachSection, read.Document.Sections[0].NoteSettings.EndnoteRestart);
        Assert.Equal(RtfNoteNumberFormat.LowerLetter, read.Document.Sections[0].NoteSettings.EndnoteNumberFormat);
        Assert.Equal(RtfEndnotePlacement.SectionEnd, read.Document.Sections[0].NoteSettings.EndnotePlacement);
        Assert.Equal(RtfPageBorderScope.FirstPageInSection, read.Document.Sections[0].PageSetup.PageBorders.Scope);
        Assert.True(read.Document.Sections[0].PageSetup.PageBorders.DisplayBehindText);
        Assert.Equal(RtfPageBorderOffset.PageEdge, read.Document.Sections[0].PageSetup.PageBorders.OffsetFrom);
        Assert.Equal(RtfPageBorderStyle.Dotted, read.Document.Sections[0].PageSetup.PageBorders.Left.Style);
        Assert.Equal(RtfPageBorderStyle.Dashed, read.Document.Sections[0].PageSetup.PageBorders.Right.Style);
        Assert.Equal("Second", Assert.IsType<RtfParagraph>(Assert.Single(read.Document.Sections[1].Blocks)).ToPlainText());
        Assert.Equal(RtfSectionBreakKind.Continuous, read.Document.Sections[1].BreakKind);
        Assert.Equal(RtfSectionVerticalAlignment.Bottom, read.Document.Sections[1].VerticalAlignment);
        Assert.Equal(0, read.Document.Sections[1].LineNumbering.CountBy);
        Assert.Equal(RtfLineNumberRestart.Continuous, read.Document.Sections[1].LineNumbering.Restart);
    }

    [Theory]
    [InlineData(RtfSectionVerticalAlignment.Top, @"\vertalt")]
    [InlineData(RtfSectionVerticalAlignment.Center, @"\vertalc")]
    [InlineData(RtfSectionVerticalAlignment.Bottom, @"\vertalb")]
    [InlineData(RtfSectionVerticalAlignment.Justified, @"\vertalj")]
    public void Write_And_Read_Section_Vertical_Alignment_Controls(RtfSectionVerticalAlignment alignment, string controlWord) {
        RtfDocument document = RtfDocument.Create();
        RtfSection section = document.AddSection();
        section.SetVerticalAlignment(alignment);
        section.AddParagraph("Vertical alignment");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(controlWord, rtf, StringComparison.Ordinal);
        Assert.Equal(alignment, Assert.Single(read.Document.Sections).VerticalAlignment);
    }

    [Theory]
    [InlineData(RtfLineNumberRestart.EachSection, @"\linerestart")]
    [InlineData(RtfLineNumberRestart.EachPage, @"\lineppage")]
    [InlineData(RtfLineNumberRestart.Continuous, @"\linecont")]
    public void Write_And_Read_Line_Numbering_Restart_Controls(RtfLineNumberRestart restart, string controlWord) {
        RtfDocument document = RtfDocument.Create();
        RtfSection section = document.AddSection();
        section.LineNumbering.Set(restart: restart);
        section.AddParagraph("Line numbering");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(controlWord, rtf, StringComparison.Ordinal);
        Assert.Equal(restart, Assert.Single(read.Document.Sections).LineNumbering.Restart);
    }

    [Fact]
    public void Write_And_Read_Headers_And_Footers() {
        RtfDocument document = RtfDocument.Create();
        document.AddHeader().AddParagraph("Default header").AddText(" bold").SetBold();
        document.AddHeader(RtfHeaderFooterKind.FirstHeader).AddParagraph("First header");
        document.AddFooter().AddParagraph("Page footer");
        document.AddParagraph("Body");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"{\header\pard\ql Default header\b  bold\b0 \par", rtf, StringComparison.Ordinal);
        Assert.Contains(@"{\headerf\pard\ql First header\par", rtf, StringComparison.Ordinal);
        Assert.Contains(@"{\footer\pard\ql Page footer\par", rtf, StringComparison.Ordinal);
        Assert.Equal("Body", Assert.Single(read.Document.Paragraphs).ToPlainText());
        Assert.Equal(3, read.Document.HeaderFooters.Count);
        Assert.Equal(RtfHeaderFooterKind.Header, read.Document.HeaderFooters[0].Kind);
        Assert.Equal("Default header bold", read.Document.HeaderFooters[0].ToPlainText());
        Assert.Contains(read.Document.HeaderFooters[0].Paragraphs[0].Runs, run => run.Text == " bold" && run.Bold);
        Assert.Equal(RtfHeaderFooterKind.FirstHeader, read.Document.HeaderFooters[1].Kind);
        Assert.Equal("First header", read.Document.HeaderFooters[1].ToPlainText());
        Assert.Equal(RtfHeaderFooterKind.Footer, read.Document.HeaderFooters[2].Kind);
        Assert.Equal("Page footer", read.Document.HeaderFooters[2].ToPlainText());
    }

    [Fact]
    public void Write_And_Read_Run_Attached_Footnotes() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.AddText("Body");
        paragraph.AddFootnote("1", "Footnote text");
        paragraph.AddText(" after");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"\super 1{\footnote\pard\ql Footnote text\par", rtf, StringComparison.Ordinal);
        Assert.Equal("Body1 after", Assert.Single(read.Document.Paragraphs).ToPlainText());
        RtfRun referenceRun = read.Document.Paragraphs[0].Runs.Single(run => run.Text == "1");
        Assert.Equal(RtfVerticalPosition.Superscript, referenceRun.VerticalPosition);
        Assert.NotNull(referenceRun.Note);
        Assert.Equal(RtfNoteKind.Footnote, referenceRun.Note!.Kind);
        Assert.Equal("Footnote text", referenceRun.Note.ToPlainText());
        RtfNote note = Assert.Single(read.Document.Notes);
        Assert.Same(referenceRun.Note, note);
    }

    [Fact]
    public void Write_And_Read_Run_Attached_Endnotes() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.AddText("Body");
        paragraph.AddEndnote("i", "Endnote text");
        paragraph.AddText(" after");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"\super i{\endnote\pard\ql Endnote text\par", rtf, StringComparison.Ordinal);
        Assert.Equal("Bodyi after", Assert.Single(read.Document.Paragraphs).ToPlainText());
        RtfRun referenceRun = read.Document.Paragraphs[0].Runs.Single(run => run.Text == "i");
        Assert.Equal(RtfVerticalPosition.Superscript, referenceRun.VerticalPosition);
        Assert.NotNull(referenceRun.Note);
        Assert.Equal(RtfNoteKind.Endnote, referenceRun.Note!.Kind);
        Assert.Equal("Endnote text", referenceRun.Note.ToPlainText());
        RtfNote note = Assert.Single(read.Document.Notes);
        Assert.Same(referenceRun.Note, note);
    }
}
