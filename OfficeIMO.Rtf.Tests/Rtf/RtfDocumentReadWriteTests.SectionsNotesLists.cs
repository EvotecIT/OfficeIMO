using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Diagnostics;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public partial class RtfDocumentReadWriteTests {
    [Fact]
    public void Read_Binds_Page_Setup_And_Preserves_Source_Losslessly() {
        const string rtf = @"{\rtf1\ansi\paperw11906\paperh16838\margl1440\margr720\margt1080\margb1080\gutter180\headery360\footery540\rtlgutter\pgnstarts5\pgnrestart\pgnx720\pgny900\pgnlcrm\pgbrdrhead\pgbrdrfoot\pgbrdropt34\pgbrdrsnap\pgbrdrt\brdrs\brdrw12\brsp24\brdrcf1\pgbrdrb\brdrdb\brdrw18\brsp30\brdrcf2\landscape\titlepg\pard Body\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        Assert.Equal(11906, result.Document.PageSetup.PaperWidthTwips);
        Assert.Equal(16838, result.Document.PageSetup.PaperHeightTwips);
        Assert.Equal(1440, result.Document.PageSetup.MarginLeftTwips);
        Assert.Equal(720, result.Document.PageSetup.MarginRightTwips);
        Assert.Equal(1080, result.Document.PageSetup.MarginTopTwips);
        Assert.Equal(1080, result.Document.PageSetup.MarginBottomTwips);
        Assert.Equal(180, result.Document.PageSetup.GutterWidthTwips);
        Assert.Equal(360, result.Document.PageSetup.HeaderDistanceTwips);
        Assert.Equal(540, result.Document.PageSetup.FooterDistanceTwips);
        Assert.True(result.Document.PageSetup.RtlGutter);
        Assert.Equal(5, result.Document.PageSetup.PageNumberStart);
        Assert.True(result.Document.PageSetup.PageNumberRestart);
        Assert.Equal(720, result.Document.PageSetup.PageNumberPositionXTwips);
        Assert.Equal(900, result.Document.PageSetup.PageNumberPositionYTwips);
        Assert.Equal(RtfPageNumberFormat.LowerRoman, result.Document.PageSetup.PageNumberFormat);
        Assert.True(result.Document.PageSetup.PageBorders.IncludeHeader);
        Assert.True(result.Document.PageSetup.PageBorders.IncludeFooter);
        Assert.True(result.Document.PageSetup.PageBorders.SnapToPageBorder);
        Assert.Equal(RtfPageBorderScope.AllExceptFirstPageInSection, result.Document.PageSetup.PageBorders.Scope);
        Assert.False(result.Document.PageSetup.PageBorders.DisplayBehindText);
        Assert.Equal(RtfPageBorderOffset.PageEdge, result.Document.PageSetup.PageBorders.OffsetFrom);
        Assert.Equal(RtfPageBorderStyle.Single, result.Document.PageSetup.PageBorders.Top.Style);
        Assert.Equal(12, result.Document.PageSetup.PageBorders.Top.Width);
        Assert.Equal(24, result.Document.PageSetup.PageBorders.Top.Space);
        Assert.Equal(1, result.Document.PageSetup.PageBorders.Top.ColorIndex);
        Assert.Equal(RtfPageBorderStyle.Double, result.Document.PageSetup.PageBorders.Bottom.Style);
        Assert.Equal(18, result.Document.PageSetup.PageBorders.Bottom.Width);
        Assert.Equal(30, result.Document.PageSetup.PageBorders.Bottom.Space);
        Assert.Equal(2, result.Document.PageSetup.PageBorders.Bottom.ColorIndex);
        Assert.True(result.Document.PageSetup.Landscape);
        Assert.True(result.Document.PageSetup.DifferentFirstPageHeaderFooter);
        Assert.Equal("Body", Assert.Single(result.Document.Paragraphs).ToPlainText());
    }

    [Fact]
    public void Read_Binds_Note_Settings_And_Preserves_Source_Losslessly() {
        const string rtf = @"{\rtf1\ansi\ftnstart5\ftnrestart\ftnnrlc\ftntj\aftnstart7\aftnrstcont\aftnnar\aenddoc\pard Body\par}";

        RtfReadResult result = RtfDocument.Read(rtf);
        RtfNoteSettings settings = result.Document.NoteSettings;

        Assert.Equal(rtf, result.ToRtfLossless());
        Assert.Equal(5, settings.FootnoteStartNumber);
        Assert.Equal(RtfNoteNumberRestart.EachSection, settings.FootnoteRestart);
        Assert.Equal(RtfNoteNumberFormat.LowerRoman, settings.FootnoteNumberFormat);
        Assert.Equal(RtfFootnotePlacement.BeneathText, settings.FootnotePlacement);
        Assert.Equal(7, settings.EndnoteStartNumber);
        Assert.Equal(RtfNoteNumberRestart.Continuous, settings.EndnoteRestart);
        Assert.Equal(RtfNoteNumberFormat.Arabic, settings.EndnoteNumberFormat);
        Assert.Equal(RtfEndnotePlacement.DocumentEnd, settings.EndnotePlacement);
        Assert.Equal("Body", Assert.Single(result.Document.Paragraphs).ToPlainText());
    }

    [Fact]
    public void Read_Binds_Document_Settings_And_Preserves_Source_Losslessly() {
        const string rtf = @"{\rtf1\ansi\uc2\deftab720\deflang1045\deflangfe1041\adeflang1033\viewkind4\viewscale125\viewzk2\viewbksp1\widowctrl\hyphauto0\hyphcaps\hyphconsec3\hyphhotz360\facingp\margmirror\formprot\revprot0\annotprot\readprot0\revisions\revprop3\revbar2\dghspace120\dgvspace180\dghorigin720\dgvorigin900\dghshow2\dgvshow3\dgsnap\dgmargin0\pard Body\par}";

        RtfReadResult result = RtfDocument.Read(rtf);
        RtfDocumentSettings settings = result.Document.Settings;

        Assert.Equal(rtf, result.ToRtfLossless());
        Assert.Equal(2, settings.UnicodeSkipCount);
        Assert.Equal(720, settings.DefaultTabWidthTwips);
        Assert.Equal(1045, settings.DefaultLanguageId);
        Assert.Equal(1041, settings.DefaultFarEastLanguageId);
        Assert.Equal(1033, settings.DefaultAlternateLanguageId);
        Assert.Equal(4, settings.ViewKind);
        Assert.Equal(125, settings.ViewScale);
        Assert.Equal(2, settings.ZoomKind);
        Assert.Equal(1, settings.ViewBackspaceBehavior);
        Assert.True(settings.WidowOrphanControl);
        Assert.False(settings.AutoHyphenation);
        Assert.True(settings.HyphenateCaps);
        Assert.Equal(3, settings.ConsecutiveHyphenLimit);
        Assert.Equal(360, settings.HyphenationZoneTwips);
        Assert.True(settings.FacingPages);
        Assert.True(settings.MirrorMargins);
        Assert.True(settings.FormProtection);
        Assert.False(settings.RevisionProtection);
        Assert.True(settings.AnnotationProtection);
        Assert.False(settings.ReadOnlyProtection);
        Assert.True(settings.TrackRevisions);
        Assert.Equal(3, settings.RevisionDisplayStyle);
        Assert.Equal(2, settings.RevisionBarPlacement);
        Assert.Equal(120, settings.DrawingGridHorizontalSpacingTwips);
        Assert.Equal(180, settings.DrawingGridVerticalSpacingTwips);
        Assert.Equal(720, settings.DrawingGridHorizontalOriginTwips);
        Assert.Equal(900, settings.DrawingGridVerticalOriginTwips);
        Assert.Equal(2, settings.DrawingGridHorizontalShow);
        Assert.Equal(3, settings.DrawingGridVerticalShow);
        Assert.True(settings.SnapToDrawingGrid);
        Assert.False(settings.DrawingGridUsesMargins);
        Assert.Equal("Body", Assert.Single(result.Document.Paragraphs).ToPlainText());
    }

    [Fact]
    public void Read_Binds_Sections_Columns_And_Section_Page_Setup_Losslessly() {
        const string rtf = @"{\rtf1\ansi\sectd\sbkodd\cols2\colsx720\colno1\colw3000\colsr360\colno2\colw4000\colsr0\linebetcol\linemod2\linex360\linestarts10\lineppage\vertalj\pgwsxn10000\pghsxn12000\marglsxn720\guttersxn240\headery300\footery420\rtlgutter\pgnstarts3\pgnrestart\pgnucrm\lndscpsxn\titlepg\ftnstart3\ftnrstpg\ftnnauc\ftnbj\aftnstart9\aftnrestart\aftnnalc\aendnotes\pgbrdropt41\pgbrdrl\brdrdot\brdrw8\brsp12\brdrcf3\pgbrdrr\brdrdash\brdrw10\brsp14\brdrcf4\pard First\par\sect\sectd\sbknone\vertalt\cols1\linemod0\linecont\pgncont\pgndec\pard Second\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        Assert.False(result.Document.PageSetup.HasAnyValue);
        Assert.Equal(2, result.Document.Sections.Count);
        RtfSection first = result.Document.Sections[0];
        Assert.Equal(RtfSectionBreakKind.OddPage, first.BreakKind);
        Assert.Equal(2, first.ColumnCount);
        Assert.Equal(720, first.ColumnSpaceTwips);
        Assert.True(first.ColumnSeparator);
        Assert.Collection(first.Columns,
            column => {
                Assert.Equal(3000, column.WidthTwips);
                Assert.Equal(360, column.SpaceAfterTwips);
            },
            column => {
                Assert.Equal(4000, column.WidthTwips);
                Assert.Equal(0, column.SpaceAfterTwips);
            });
        Assert.Equal(2, first.LineNumbering.CountBy);
        Assert.Equal(360, first.LineNumbering.DistanceFromTextTwips);
        Assert.Equal(10, first.LineNumbering.StartNumber);
        Assert.Equal(RtfLineNumberRestart.EachPage, first.LineNumbering.Restart);
        Assert.Equal(RtfSectionVerticalAlignment.Justified, first.VerticalAlignment);
        Assert.Equal(10000, first.PageSetup.PaperWidthTwips);
        Assert.Equal(12000, first.PageSetup.PaperHeightTwips);
        Assert.Equal(720, first.PageSetup.MarginLeftTwips);
        Assert.Equal(240, first.PageSetup.GutterWidthTwips);
        Assert.Equal(300, first.PageSetup.HeaderDistanceTwips);
        Assert.Equal(420, first.PageSetup.FooterDistanceTwips);
        Assert.True(first.PageSetup.RtlGutter);
        Assert.Equal(3, first.PageSetup.PageNumberStart);
        Assert.True(first.PageSetup.PageNumberRestart);
        Assert.Equal(RtfPageNumberFormat.UpperRoman, first.PageSetup.PageNumberFormat);
        Assert.True(first.PageSetup.Landscape);
        Assert.True(first.PageSetup.DifferentFirstPageHeaderFooter);
        Assert.Equal(3, first.NoteSettings.FootnoteStartNumber);
        Assert.Equal(RtfNoteNumberRestart.EachPage, first.NoteSettings.FootnoteRestart);
        Assert.Equal(RtfNoteNumberFormat.UpperLetter, first.NoteSettings.FootnoteNumberFormat);
        Assert.Equal(RtfFootnotePlacement.PageBottom, first.NoteSettings.FootnotePlacement);
        Assert.Equal(9, first.NoteSettings.EndnoteStartNumber);
        Assert.Equal(RtfNoteNumberRestart.EachSection, first.NoteSettings.EndnoteRestart);
        Assert.Equal(RtfNoteNumberFormat.LowerLetter, first.NoteSettings.EndnoteNumberFormat);
        Assert.Equal(RtfEndnotePlacement.SectionEnd, first.NoteSettings.EndnotePlacement);
        Assert.Equal(RtfPageBorderScope.FirstPageInSection, first.PageSetup.PageBorders.Scope);
        Assert.True(first.PageSetup.PageBorders.DisplayBehindText);
        Assert.Equal(RtfPageBorderOffset.PageEdge, first.PageSetup.PageBorders.OffsetFrom);
        Assert.Equal(RtfPageBorderStyle.Dotted, first.PageSetup.PageBorders.Left.Style);
        Assert.Equal(8, first.PageSetup.PageBorders.Left.Width);
        Assert.Equal(12, first.PageSetup.PageBorders.Left.Space);
        Assert.Equal(3, first.PageSetup.PageBorders.Left.ColorIndex);
        Assert.Equal(RtfPageBorderStyle.Dashed, first.PageSetup.PageBorders.Right.Style);
        Assert.Equal(10, first.PageSetup.PageBorders.Right.Width);
        Assert.Equal(14, first.PageSetup.PageBorders.Right.Space);
        Assert.Equal(4, first.PageSetup.PageBorders.Right.ColorIndex);
        Assert.Equal("First", Assert.IsType<RtfParagraph>(Assert.Single(first.Blocks)).ToPlainText());
        RtfSection second = result.Document.Sections[1];
        Assert.Equal(RtfSectionBreakKind.Continuous, second.BreakKind);
        Assert.Equal(1, second.ColumnCount);
        Assert.Equal(0, second.LineNumbering.CountBy);
        Assert.Equal(RtfLineNumberRestart.Continuous, second.LineNumbering.Restart);
        Assert.Equal(RtfSectionVerticalAlignment.Top, second.VerticalAlignment);
        Assert.False(second.PageSetup.PageNumberRestart);
        Assert.Equal(RtfPageNumberFormat.Decimal, second.PageSetup.PageNumberFormat);
        Assert.Equal("Second", Assert.IsType<RtfParagraph>(Assert.Single(second.Blocks)).ToPlainText());
    }

    [Fact]
    public void Read_Binds_Header_Footer_Destinations_And_Preserves_Source_Losslessly() {
        const string rtf = @"{\rtf1\ansi{\header\pard Header {\b bold}\par}{\footer\pard Footer\par}\pard Body\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        Assert.Equal("Body", Assert.Single(result.Document.Paragraphs).ToPlainText());
        Assert.Equal(2, result.Document.HeaderFooters.Count);
        Assert.Equal(RtfHeaderFooterKind.Header, result.Document.HeaderFooters[0].Kind);
        Assert.Equal("Header bold", result.Document.HeaderFooters[0].ToPlainText());
        Assert.Contains(result.Document.HeaderFooters[0].Paragraphs[0].Runs, run => run.Text == "bold" && run.Bold);
        Assert.Equal(RtfHeaderFooterKind.Footer, result.Document.HeaderFooters[1].Kind);
        Assert.Equal("Footer", result.Document.HeaderFooters[1].ToPlainText());
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "RTF101");
    }

    [Fact]
    public void Read_Binds_Footnotes_Without_Leaking_Note_Text_Into_Body() {
        const string rtf = @"{\rtf1\ansi\pard Body{\super 1}{\footnote\pard Footnote {\i text}\par} after\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        RtfParagraph body = Assert.Single(result.Document.Paragraphs);
        Assert.Equal("Body1 after", body.ToPlainText());
        Assert.DoesNotContain("Footnote", body.ToPlainText(), StringComparison.Ordinal);
        RtfRun referenceRun = body.Runs.Single(run => run.Text == "1");
        Assert.NotNull(referenceRun.Note);
        Assert.Equal(RtfNoteKind.Footnote, referenceRun.Note!.Kind);
        Assert.Equal("Footnote text", referenceRun.Note.ToPlainText());
        Assert.Contains(referenceRun.Note.Paragraphs[0].Runs, run => run.Text == "text" && run.Italic);
        Assert.Same(referenceRun.Note, Assert.Single(result.Document.Notes));
    }

    [Fact]
    public void Read_Binds_Endnotes_Without_Leaking_Note_Text_Into_Body() {
        const string rtf = @"{\rtf1\ansi\pard Body{\super i}{\endnote\pard Endnote {\b text}\par} after\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        RtfParagraph body = Assert.Single(result.Document.Paragraphs);
        Assert.Equal("Bodyi after", body.ToPlainText());
        Assert.DoesNotContain("Endnote", body.ToPlainText(), StringComparison.Ordinal);
        RtfRun referenceRun = body.Runs.Single(run => run.Text == "i");
        Assert.NotNull(referenceRun.Note);
        Assert.Equal(RtfNoteKind.Endnote, referenceRun.Note!.Kind);
        Assert.Equal("Endnote text", referenceRun.Note.ToPlainText());
        Assert.Contains(referenceRun.Note.Paragraphs[0].Runs, run => run.Text == "text" && run.Bold);
        Assert.Same(referenceRun.Note, Assert.Single(result.Document.Notes));
    }

    [Fact]
    public void Read_Binds_List_Table_Overrides_And_Paragraph_List_Metadata_Losslessly() {
        const string rtf = @"{\rtf1\ansi{\*\listtable{\list{\listlevel\levelnfc23\levelnfcn23\levelstartat1{\leveltext\'01\u8226 ?;}{\levelnumbers;}\fi-360\li720}{\listname Bullet;}\listid100}}{\*\listoverridetable{\listoverride\listid100\listoverridecount0\ls3}}\pard\ls3\ilvl0 Item\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        RtfListDefinition definition = Assert.Single(result.Document.ListDefinitions);
        Assert.Equal(100, definition.Id);
        Assert.Equal("Bullet", definition.Name);
        RtfListLevel level = Assert.Single(definition.Levels);
        Assert.Equal(RtfListKind.Bullet, level.Kind);
        Assert.Equal(23, level.NumberFormat);
        Assert.Equal(720, level.LeftIndentTwips);
        Assert.Equal(-360, level.FirstLineIndentTwips);
        RtfListOverride listOverride = Assert.Single(result.Document.ListOverrides);
        Assert.Equal(3, listOverride.Id);
        Assert.Equal(100, listOverride.ListId);
        RtfParagraph paragraph = Assert.Single(result.Document.Paragraphs);
        Assert.Equal("Item", paragraph.ToPlainText());
        Assert.Equal(3, paragraph.ListId);
        Assert.Equal(100, paragraph.ListDefinitionId);
        Assert.Equal(0, paragraph.ListLevel);
        Assert.Equal(RtfListKind.Bullet, paragraph.ListKind);
    }

    [Fact]
    public void Read_Binds_Rich_List_Level_Metadata_Losslessly() {
        const string rtf = @"{\rtf1\ansi{\*\listtable{\list{\listlevel\levelnfc0\levelnfcn2\leveljc2\leveljcn1\levelfollow1\levelstartat7\levelspace120\levelindent240\levellegal1\levelnorestart1\levelpicture3\levelpicturenosize{\leveltext\'03\'00.;}{\levelnumbers\'01\'00;}\fi-360\li1080}{\listname Rich;}\listid100}}{\*\listoverridetable{\listoverride\listid100\listoverridecount0\ls3}}\pard\ls3\ilvl0 Item\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        RtfListDefinition definition = Assert.Single(result.Document.ListDefinitions);
        RtfListLevel level = Assert.Single(definition.Levels);
        Assert.Equal(0, level.NumberFormat);
        Assert.Equal(2, level.NumberFormatN);
        Assert.Equal(RtfListLevelAlignment.Right, level.Alignment);
        Assert.Equal(RtfListLevelAlignment.Center, level.AlignmentN);
        Assert.Equal(RtfListLevelFollowCharacter.Space, level.FollowCharacter);
        Assert.Equal(7, level.StartAt);
        Assert.Equal(120, level.SpaceTwips);
        Assert.Equal(240, level.IndentTwips);
        Assert.True(level.LegalNumbering);
        Assert.True(level.NoRestart);
        Assert.Equal(3, level.PictureIndex);
        Assert.True(level.PictureNoSize);
        Assert.Equal(1080, level.LeftIndentTwips);
        Assert.Equal(-360, level.FirstLineIndentTwips);
    }

    [Fact]
    public void Read_Binds_List_Level_Overrides_Losslessly() {
        const string rtf = @"{\rtf1\ansi{\*\listtable{\list{\listlevel\levelnfc0\levelnfcn0\levelstartat1{\leveltext\'03\'00.;}{\levelnumbers\'01\'00;}\fi-360\li720}{\listname Decimal;}\listid100}}{\*\listoverridetable{\listoverride\listid100\listoverridecount2{\lfolevel\listoverrideformat1\listoverridestartat1\levelstartat9}{\lfolevel\listoverridestartat0}\ls3}}\pard\ls3\ilvl0 Item\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        RtfListOverride listOverride = Assert.Single(result.Document.ListOverrides);
        Assert.Equal(3, listOverride.Id);
        Assert.Equal(100, listOverride.ListId);
        Assert.Equal(2, listOverride.OverrideCount);
        Assert.Collection(listOverride.LevelOverrides,
            levelOverride => {
                Assert.True(levelOverride.OverrideFormat);
                Assert.True(levelOverride.OverrideStartAt);
                Assert.Equal(9, levelOverride.StartAt);
            },
            levelOverride => {
                Assert.Null(levelOverride.OverrideFormat);
                Assert.False(levelOverride.OverrideStartAt);
                Assert.Null(levelOverride.StartAt);
            });
    }

    [Fact]
    public void Read_Binds_Tab_Stops_And_Preserves_Source_Losslessly() {
        const string rtf = @"{\rtf1\ansi\pard\tx1440\tqr\tldot\tx2880\tqdec\tlmdot\tx4320\tb5000 Name\tab Amount\tab 12.34\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        RtfParagraph paragraph = Assert.Single(result.Document.Paragraphs);
        Assert.Equal("Name\tAmount\t12.34", paragraph.ToPlainText());
        Assert.Collection(paragraph.TabStops,
            tabStop => {
                Assert.Equal(1440, tabStop.PositionTwips);
                Assert.Equal(RtfTabAlignment.Left, tabStop.Alignment);
                Assert.Equal(RtfTabLeader.None, tabStop.Leader);
            },
            tabStop => {
                Assert.Equal(2880, tabStop.PositionTwips);
                Assert.Equal(RtfTabAlignment.Right, tabStop.Alignment);
                Assert.Equal(RtfTabLeader.Dots, tabStop.Leader);
            },
            tabStop => {
                Assert.Equal(4320, tabStop.PositionTwips);
                Assert.Equal(RtfTabAlignment.Decimal, tabStop.Alignment);
                Assert.Equal(RtfTabLeader.MiddleDots, tabStop.Leader);
            },
            tabStop => {
                Assert.Equal(5000, tabStop.PositionTwips);
                Assert.Equal(RtfTabAlignment.Bar, tabStop.Alignment);
                Assert.Equal(RtfTabLeader.None, tabStop.Leader);
            });
    }

    [Fact]
    public void Read_Binds_Inline_Breaks_And_Preserves_Source_Losslessly() {
        const string rtf = @"{\rtf1\ansi\pard Before\line Line\softline SoftLine\page Page\softpage SoftPage\column Column\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        RtfParagraph paragraph = Assert.Single(result.Document.Paragraphs);
        Assert.Equal("Before" + Environment.NewLine + "Line" + Environment.NewLine + "SoftLine\fPage\fSoftPage\vColumn", paragraph.ToPlainText());
        Assert.Collection(paragraph.Inlines,
            inline => Assert.Equal("Before", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(RtfBreakKind.Line, Assert.IsType<RtfBreak>(inline).Kind),
            inline => Assert.Equal("Line", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(RtfBreakKind.SoftLine, Assert.IsType<RtfBreak>(inline).Kind),
            inline => Assert.Equal("SoftLine", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(RtfBreakKind.Page, Assert.IsType<RtfBreak>(inline).Kind),
            inline => Assert.Equal("Page", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(RtfBreakKind.SoftPage, Assert.IsType<RtfBreak>(inline).Kind),
            inline => Assert.Equal("SoftPage", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(RtfBreakKind.Column, Assert.IsType<RtfBreak>(inline).Kind),
            inline => Assert.Equal("Column", Assert.IsType<RtfRun>(inline).Text));
    }

    [Fact]
    public void Read_And_Write_Bound_Hostile_List_Levels() {
        const string rtf = @"{\rtf1\ansi\pard\ls1\ilvl2147483647 Item\par}";

        RtfReadResult result = RtfDocument.Read(rtf);
        RtfParagraph paragraph = Assert.Single(result.Document.Paragraphs);

        Assert.Equal(8, paragraph.ListLevel);
        string rewritten = result.Document.ToRtf();
        Assert.Contains(@"\ilvl8", rewritten, StringComparison.Ordinal);
        Assert.DoesNotContain("2147483647", rewritten, StringComparison.Ordinal);
    }
}
