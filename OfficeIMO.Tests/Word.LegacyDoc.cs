using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc;
using OfficeIMO.Word.LegacyDoc.Diagnostics;
using OfficeIMO.Word.LegacyDoc.Model;
using OfficeIMO.Shared;
using System.Text;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenMcdf;
using Xunit;
using Version = OpenMcdf.Version;
using StorageModeFlags = OpenMcdf.StorageModeFlags;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsPlainTextParagraphs() {
            byte[] docBytes = LegacyDocTestBuilder.CreateSimpleDoc("First paragraph", "Second paragraph");

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            Assert.True(result.HasDocument);
            Assert.Equal(2, result.ImportReport.ParagraphCount);
            Assert.Equal(2, result.Document.Paragraphs.Count);
            Assert.Equal("First paragraph", result.Document.Paragraphs[0].Text);
            Assert.Equal("Second paragraph", result.Document.Paragraphs[1].Text);
            Assert.True(result.Document.WasLoadedFromLegacyDoc);
            Assert.Equal(string.Empty, result.Document.FilePath);

            using WordDocument reloaded = WordDocument.Load(new MemoryStream(result.Document.SaveAsByteArray()));
            Assert.Equal("First paragraph", reloaded.Paragraphs[0].Text);
            Assert.Equal("Second paragraph", reloaded.Paragraphs[1].Text);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsSimpleTable() {
            byte[] docBytes = LegacyDocTestBuilder.CreateSimpleDocWithTable();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            Assert.True(result.HasDocument);
            WordTable table = Assert.Single(result.Document.Tables);
            Assert.Equal(2, table.Rows.Count);
            Assert.Equal(2, table.Rows[0].Cells.Count);
            Assert.Equal("A1", table.Rows[0].Cells[0].Paragraphs[0].Text);
            Assert.Equal("B1", table.Rows[0].Cells[1].Paragraphs[0].Text);
            Assert.Equal("A2", table.Rows[1].Cells[0].Paragraphs[0].Text);
            Assert.Equal("B2", table.Rows[1].Cells[1].Paragraphs[0].Text);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsExplicitTableMarkerTrailingEmptyCell() {
            byte[] docBytes = LegacyDocTestBuilder.CreateUnicodeDocWithExplicitTableMarkersAndTrailingEmptyCell();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            Assert.True(result.HasDocument);
            WordTable table = Assert.Single(result.Document.Tables);
            WordTableRow row = Assert.Single(table.Rows);
            Assert.Equal(2, row.Cells.Count);
            Assert.Equal("A1", row.Cells[0].Paragraphs[0].Text);
            Assert.Equal(string.Empty, row.Cells[1].Paragraphs[0].Text);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsTableCellWidthsFromRowDefinition() {
            byte[] docBytes = LegacyDocTestBuilder.CreateUnicodeDocWithTableCellWidths();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            Assert.True(result.HasDocument);
            WordTable table = Assert.Single(result.Document.Tables);
            WordTableRow row = Assert.Single(table.Rows);
            Assert.Equal(1440, row.Cells[0].Width);
            Assert.Equal(TableWidthUnitValues.Dxa, row.Cells[0].WidthType);
            Assert.Equal(2880, row.Cells[1].Width);
            Assert.Equal(TableWidthUnitValues.Dxa, row.Cells[1].WidthType);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsFormattedTableCellRuns() {
            byte[] docBytes = LegacyDocTestBuilder.CreateUnicodeDocWithFormattedTableCell();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            Assert.True(result.HasDocument);
            WordTable table = Assert.Single(result.Document.Tables);
            WordParagraph firstCellRun = table.Rows[0].Cells[0].Paragraphs[0];
            WordParagraph secondCellRun = table.Rows[0].Cells[1].Paragraphs[0];
            Assert.Equal("A1", firstCellRun.Text);
            Assert.True(firstCellRun.Bold);
            Assert.Equal("B1", secondCellRun.Text);
            Assert.False(secondCellRun.Bold);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsTableCellParagraphFormatting() {
            byte[] docBytes = LegacyDocTestBuilder.CreateUnicodeDocWithFormattedTableCellParagraph();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            Assert.True(result.HasDocument);
            WordTable table = Assert.Single(result.Document.Tables);
            WordParagraph firstCellParagraph = table.Rows[0].Cells[0].Paragraphs[0];
            WordParagraph secondCellParagraph = table.Rows[0].Cells[1].Paragraphs[0];
            Assert.Equal("A1", firstCellParagraph.Text);
            Assert.Equal(JustificationValues.Center, firstCellParagraph.ParagraphAlignment);
            Assert.Equal(120, firstCellParagraph.LineSpacingAfter);
            Assert.Equal(360, firstCellParagraph.IndentationBefore);
            Assert.Equal("B1", secondCellParagraph.Text);
            Assert.Null(secondCellParagraph.ParagraphAlignment);
            Assert.Null(secondCellParagraph.LineSpacingAfter);
            Assert.Null(secondCellParagraph.IndentationBefore);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsTabsAsWordTabRuns() {
            byte[] docBytes = LegacyDocTestBuilder.CreateSimpleDoc("Left\tRight");

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            Assert.True(result.HasDocument);
            Paragraph paragraph = Assert.Single(result.Document._wordprocessingDocument!.MainDocumentPart!.Document.Body!.Elements<Paragraph>());
            Assert.Equal(1, paragraph.Descendants<TabChar>().Count());
            Assert.DoesNotContain(paragraph.Descendants<Text>(), text => text.Text.Contains('\t'));
            Assert.Equal(new[] { "Left", "Right" }, paragraph.Descendants<Text>().Select(text => text.Text).ToArray());
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsLineAndPageBreaksAsWordBreakRuns() {
            byte[] docBytes = LegacyDocTestBuilder.CreateSimpleDoc("Line\vBreak\fPage");

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            Assert.True(result.HasDocument);
            Paragraph paragraph = Assert.Single(result.Document._wordprocessingDocument!.MainDocumentPart!.Document.Body!.Elements<Paragraph>());
            Break[] breaks = paragraph.Descendants<Break>().ToArray();
            Assert.Equal(2, breaks.Length);
            Assert.Null(breaks[0].Type);
            Assert.Equal(BreakValues.Page, breaks[1].Type!.Value);
            Assert.DoesNotContain(paragraph.Descendants<Text>(), text => text.Text.Contains('\v') || text.Text.Contains('\f'));
            Assert.Equal(new[] { "Line", "Break", "Page" }, paragraph.Descendants<Text>().Select(text => text.Text).ToArray());
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsDocumentPropertiesAndCustomProperties() {
            DateTime created = new DateTime(2026, 6, 29, 8, 0, 0, DateTimeKind.Utc);
            DateTime modified = new DateTime(2026, 6, 29, 9, 15, 0, DateTimeKind.Utc);
            byte[] docBytes = LegacyDocTestBuilder.CreateSimpleDocWithDocumentProperties(created, modified, "Metadata paragraph");

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            Assert.True(result.HasDocument);
            Assert.Equal(13, result.ImportReport.DocumentPropertyCount);
            Assert.Equal("Legacy DOC Metadata Title", result.Document.BuiltinDocumentProperties.Title);
            Assert.Equal("Legacy DOC metadata subject", result.Document.BuiltinDocumentProperties.Subject);
            Assert.Equal("OfficeIMO Legacy Import", result.Document.BuiltinDocumentProperties.Creator);
            Assert.Equal("doc, metadata, officeimo", result.Document.BuiltinDocumentProperties.Keywords);
            Assert.Equal("OLE SummaryInformation comments", result.Document.BuiltinDocumentProperties.Description);
            Assert.Equal("Legacy Category", result.Document.BuiltinDocumentProperties.Category);
            AssertSameInstant(created, result.Document.BuiltinDocumentProperties.Created);
            AssertSameInstant(modified, result.Document.BuiltinDocumentProperties.Modified);
            Assert.Equal("EvotecIT", result.Document.ApplicationProperties.Company);
            Assert.Equal("Document Manager", result.Document.ApplicationProperties.Manager?.Text);
            Assert.Equal("Ready", result.Document.CustomDocumentProperties["ReleaseStatus"].Text);
            Assert.True(result.Document.CustomDocumentProperties["Reviewed"].Bool);
            Assert.Equal(2003, result.Document.CustomDocumentProperties["Ticket"].NumberInteger);

            using WordDocument converted = WordDocument.Load(new MemoryStream(result.Document.SaveAsByteArray()));
            Assert.False(converted.WasLoadedFromLegacyDoc);
            Assert.Equal("Legacy DOC Metadata Title", converted.BuiltinDocumentProperties.Title);
            Assert.Equal("EvotecIT", converted.ApplicationProperties.Company);
            Assert.Equal("Ready", converted.CustomDocumentProperties["ReleaseStatus"].Text);
            Assert.True(converted.CustomDocumentProperties["Reviewed"].Bool);
            Assert.Equal(2003, converted.CustomDocumentProperties["Ticket"].NumberInteger);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsDirectBoldItalicRuns() {
            byte[] docBytes = LegacyDocTestBuilder.CreateUnicodeDocWithDirectCharacterFormatting();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph[] runs = result.Document.Paragraphs.ToArray();
            Assert.Equal(3, runs.Length);
            Assert.Equal("plain ", runs[0].Text);
            Assert.False(runs[0].Bold);
            Assert.False(runs[0].Italic);
            Assert.Equal("bold ", runs[1].Text);
            Assert.True(runs[1].Bold);
            Assert.False(runs[1].Italic);
            Assert.Equal("italic", runs[2].Text);
            Assert.False(runs[2].Bold);
            Assert.True(runs[2].Italic);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsDirectUnderlineSizeColorStrikeVerticalAndHighlightRuns() {
            byte[] docBytes = LegacyDocTestBuilder.CreateUnicodeDocWithExtendedDirectCharacterFormatting();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph[] runs = result.Document.Paragraphs.ToArray();
            Assert.Equal(17, runs.Length);
            Assert.Equal("plain ", runs[0].Text);
            Assert.Null(runs[0].Underline);
            Assert.False(runs[0].Strike);
            Assert.False(runs[0].DoubleStrike);
            Assert.False(runs[0].Outline);
            Assert.False(runs[0].Shadow);
            Assert.False(runs[0].Emboss);
            Assert.Null(runs[0]._runProperties?.Imprint);
            Assert.Null(runs[0]._runProperties?.Vanish);
            Assert.Equal(CapsStyle.None, runs[0].CapsStyle);
            Assert.Null(runs[0].VerticalTextAlignment);
            Assert.Null(runs[0].Highlight);
            Assert.Equal("under ", runs[1].Text);
            Assert.Equal(UnderlineValues.Single, runs[1].Underline);
            Assert.Equal("sized ", runs[2].Text);
            Assert.Equal(14, runs[2].FontSize);
            Assert.Equal("red ", runs[3].Text);
            Assert.Equal("ff0000", runs[3].ColorHex);
            Assert.Equal("strike ", runs[4].Text);
            Assert.True(runs[4].Strike);
            Assert.Equal("double ", runs[5].Text);
            Assert.True(runs[5].DoubleStrike);
            Assert.Equal("outline ", runs[6].Text);
            Assert.True(runs[6].Outline);
            Assert.Equal("shadow ", runs[7].Text);
            Assert.True(runs[7].Shadow);
            Assert.Equal("emboss ", runs[8].Text);
            Assert.True(runs[8].Emboss);
            Assert.Equal("imprint ", runs[9].Text);
            Assert.NotNull(runs[9]._runProperties?.Imprint);
            Assert.Equal("hidden ", runs[10].Text);
            Assert.NotNull(runs[10]._runProperties?.Vanish);
            Assert.Equal("caps ", runs[11].Text);
            Assert.Equal(CapsStyle.Caps, runs[11].CapsStyle);
            Assert.Equal("small ", runs[12].Text);
            Assert.Equal(CapsStyle.SmallCaps, runs[12].CapsStyle);
            Assert.Equal("super ", runs[13].Text);
            Assert.Equal(VerticalPositionValues.Superscript, runs[13].VerticalTextAlignment);
            Assert.Equal("sub ", runs[14].Text);
            Assert.Equal(VerticalPositionValues.Subscript, runs[14].VerticalTextAlignment);
            Assert.Equal("mark ", runs[15].Text);
            Assert.Equal(HighlightColorValues.Yellow, runs[15].Highlight);
            Assert.Equal("direct", runs[16].Text);
            Assert.Equal("336699", runs[16].ColorHex);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsFontFamilyRunsThroughFontTable() {
            byte[] docBytes = LegacyDocTestBuilder.CreateUnicodeDocWithFontFamilyFormatting();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph[] runs = result.Document.Paragraphs.ToArray();
            Assert.Equal(2, runs.Length);
            Assert.Equal("plain ", runs[0].Text);
            Assert.Null(runs[0].FontFamily);
            Assert.Equal("font", runs[1].Text);
            Assert.Equal("Courier New", runs[1].FontFamily);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsParagraphAlignment() {
            byte[] docBytes = LegacyDocTestBuilder.CreateUnicodeDocWithParagraphAlignment();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph[] paragraphs = result.Document.Paragraphs.ToArray();
            Assert.Equal(3, paragraphs.Length);
            Assert.Equal("left", paragraphs[0].Text);
            Assert.Null(paragraphs[0].ParagraphAlignment);
            Assert.Equal("center", paragraphs[1].Text);
            Assert.Equal(JustificationValues.Center, paragraphs[1].ParagraphAlignment);
            Assert.Equal("right", paragraphs[2].Text);
            Assert.Equal(JustificationValues.Right, paragraphs[2].ParagraphAlignment);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsParagraphSpacingAndIndentation() {
            byte[] docBytes = LegacyDocTestBuilder.CreateUnicodeDocWithParagraphSpacingAndIndentation();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph[] paragraphs = result.Document.Paragraphs.ToArray();
            Assert.Equal(3, paragraphs.Length);
            Assert.Equal("plain", paragraphs[0].Text);
            Assert.Null(paragraphs[0].LineSpacingBefore);
            Assert.Null(paragraphs[0].IndentationBefore);
            Assert.Equal("formatted", paragraphs[1].Text);
            Assert.Equal(240, paragraphs[1].LineSpacingBefore);
            Assert.Equal(120, paragraphs[1].LineSpacingAfter);
            Assert.Equal(360, paragraphs[1].LineSpacing);
            Assert.Equal(720, paragraphs[1].IndentationBefore);
            Assert.Equal(360, paragraphs[1].IndentationAfter);
            Assert.Equal(240, paragraphs[1].IndentationFirstLine);
            Assert.Equal("hanging", paragraphs[2].Text);
            Assert.Equal(720, paragraphs[2].IndentationBefore);
            Assert.Equal(360, paragraphs[2].IndentationHanging);
            Assert.Null(paragraphs[2].IndentationFirstLine);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsParagraphPaginationFlags() {
            byte[] docBytes = LegacyDocTestBuilder.CreateUnicodeDocWithParagraphPaginationFlags();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph[] paragraphs = result.Document.Paragraphs.ToArray();
            Assert.Equal(2, paragraphs.Length);
            Assert.Equal("plain", paragraphs[0].Text);
            Assert.False(paragraphs[0].KeepLinesTogether);
            Assert.False(paragraphs[0].KeepWithNext);
            Assert.False(paragraphs[0].PageBreakBefore);
            Assert.False(paragraphs[0].AvoidWidowAndOrphan);
            Assert.Equal("pagination", paragraphs[1].Text);
            Assert.True(paragraphs[1].KeepLinesTogether);
            Assert.True(paragraphs[1].KeepWithNext);
            Assert.True(paragraphs[1].PageBreakBefore);
            Assert.True(paragraphs[1].AvoidWidowAndOrphan);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsParagraphTabStops() {
            byte[] docBytes = LegacyDocTestBuilder.CreateUnicodeDocWithParagraphTabStops();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph[] paragraphs = result.Document.Paragraphs.ToArray();
            Assert.Equal(3, paragraphs.Length);
            Assert.Equal("plain", paragraphs[0].Text);
            Assert.Empty(paragraphs[0].TabStops);
            Assert.Equal("tabs", paragraphs[1].Text);
            Assert.Equal(3, paragraphs[1].TabStops.Count);
            Assert.Equal(1440, paragraphs[1].TabStops[0].Position);
            Assert.Equal(TabStopValues.Left, paragraphs[1].TabStops[0].Alignment);
            Assert.Equal(TabStopLeaderCharValues.None, paragraphs[1].TabStops[0].Leader);
            Assert.Equal(2880, paragraphs[1].TabStops[1].Position);
            Assert.Equal(TabStopValues.Decimal, paragraphs[1].TabStops[1].Alignment);
            Assert.Equal(TabStopLeaderCharValues.Dot, paragraphs[1].TabStops[1].Leader);
            Assert.Equal(4320, paragraphs[1].TabStops[2].Position);
            Assert.Equal(TabStopValues.Right, paragraphs[1].TabStops[2].Alignment);
            Assert.Equal(TabStopLeaderCharValues.Underscore, paragraphs[1].TabStops[2].Leader);
            Assert.Equal("clear", paragraphs[2].Text);
            Assert.Equal(2, paragraphs[2].TabStops.Count);
            Assert.Equal(1440, paragraphs[2].TabStops[0].Position);
            Assert.Equal(TabStopValues.Clear, paragraphs[2].TabStops[0].Alignment);
            Assert.Equal(2160, paragraphs[2].TabStops[1].Position);
            Assert.Equal(TabStopValues.Bar, paragraphs[2].TabStops[1].Alignment);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsStyleLevelParagraphTabStops() {
            byte[] docBytes = LegacyDocTestBuilder.CreateUnicodeDocWithStyleLevelParagraphTabStops();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph paragraph = Assert.Single(result.Document.Paragraphs);
            Assert.Equal("style tabs", paragraph.Text);
            Assert.Equal("LegacyDocTabStyle", paragraph.StyleId);

            Styles styles = result.Document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            Style tabStyle = Assert.Single(styles.Elements<Style>(), style => style.StyleId == "LegacyDocTabStyle");
            StyleParagraphProperties paragraphProperties = Assert.IsType<StyleParagraphProperties>(tabStyle.StyleParagraphProperties);
            Tabs tabs = Assert.IsType<Tabs>(paragraphProperties.GetFirstChild<Tabs>());
            TabStop[] tabStops = tabs.Elements<TabStop>().ToArray();
            Assert.Equal(2, tabStops.Length);
            TabStop centerStop = Assert.Single(tabStops, tabStop => tabStop.Position?.Value == 1800);
            Assert.Equal(TabStopValues.Center, centerStop.Val!.Value);
            Assert.Equal(TabStopLeaderCharValues.Dot, centerStop.Leader!.Value);
            TabStop clearStop = Assert.Single(tabStops, tabStop => tabStop.Position?.Value == 3600);
            Assert.Equal(TabStopValues.Clear, clearStop.Val!.Value);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsStyleLevelCapsDoubleStrikeAndVerticalPosition() {
            byte[] docBytes = LegacyDocTestBuilder.CreateUnicodeDocWithStyleLevelCapsDoubleStrikeAndVerticalPosition();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph[] paragraphs = result.Document.Paragraphs.ToArray();
            Assert.Equal(4, paragraphs.Length);
            Assert.Equal("caps style", paragraphs[0].Text);
            Assert.Equal("small style", paragraphs[1].Text);
            Assert.Equal("super style", paragraphs[2].Text);
            Assert.Equal("sub style", paragraphs[3].Text);
            Assert.Equal("LegacyDocCapsDouble", paragraphs[0].StyleId);
            Assert.Equal("LegacyDocSmallCaps", paragraphs[1].StyleId);
            Assert.Equal("LegacyDocSuper", paragraphs[2].StyleId);
            Assert.Equal("LegacyDocSub", paragraphs[3].StyleId);

            Styles styles = result.Document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            Style capsStyle = Assert.Single(styles.Elements<Style>(), style => style.StyleId == "LegacyDocCapsDouble");
            StyleRunProperties capsProperties = Assert.IsType<StyleRunProperties>(capsStyle.StyleRunProperties);
            Assert.NotNull(capsProperties.GetFirstChild<Caps>());
            Assert.NotNull(capsProperties.GetFirstChild<DoubleStrike>());
            Assert.NotNull(capsProperties.GetFirstChild<Outline>());
            Assert.NotNull(capsProperties.GetFirstChild<Shadow>());
            Assert.NotNull(capsProperties.GetFirstChild<Emboss>());
            Assert.NotNull(capsProperties.GetFirstChild<Imprint>());
            Assert.NotNull(capsProperties.GetFirstChild<Vanish>());

            Style smallCapsStyle = Assert.Single(styles.Elements<Style>(), style => style.StyleId == "LegacyDocSmallCaps");
            StyleRunProperties smallCapsProperties = Assert.IsType<StyleRunProperties>(smallCapsStyle.StyleRunProperties);
            Assert.NotNull(smallCapsProperties.GetFirstChild<SmallCaps>());

            Style superStyle = Assert.Single(styles.Elements<Style>(), style => style.StyleId == "LegacyDocSuper");
            StyleRunProperties superProperties = Assert.IsType<StyleRunProperties>(superStyle.StyleRunProperties);
            VerticalTextAlignment superPosition = Assert.IsType<VerticalTextAlignment>(superProperties.GetFirstChild<VerticalTextAlignment>());
            Assert.Equal(VerticalPositionValues.Superscript, superPosition.Val!.Value);

            Style subStyle = Assert.Single(styles.Elements<Style>(), style => style.StyleId == "LegacyDocSub");
            StyleRunProperties subProperties = Assert.IsType<StyleRunProperties>(subStyle.StyleRunProperties);
            VerticalTextAlignment subPosition = Assert.IsType<VerticalTextAlignment>(subProperties.GetFirstChild<VerticalTextAlignment>());
            Assert.Equal(VerticalPositionValues.Subscript, subPosition.Val!.Value);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsBuiltInStyleLevelFormatting() {
            byte[] docBytes = LegacyDocTestBuilder.CreateUnicodeDocWithBuiltInStyleLevelFormatting();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph paragraph = Assert.Single(result.Document.Paragraphs);
            Assert.Equal("built in heading", paragraph.Text);
            Assert.Equal(WordParagraphStyles.Heading1.ToStringStyle(), paragraph.StyleId);

            Styles styles = result.Document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            Style headingStyle = Assert.Single(styles.Elements<Style>(), style => style.StyleId == WordParagraphStyles.Heading1.ToStringStyle());
            StyleParagraphProperties paragraphProperties = Assert.IsType<StyleParagraphProperties>(headingStyle.StyleParagraphProperties);
            Justification justification = Assert.IsType<Justification>(paragraphProperties.GetFirstChild<Justification>());
            Assert.Equal(JustificationValues.Center, justification.Val!.Value);
            SpacingBetweenLines spacing = Assert.IsType<SpacingBetweenLines>(paragraphProperties.GetFirstChild<SpacingBetweenLines>());
            Assert.Equal("240", spacing.Before!.Value);
            Assert.Equal("120", spacing.After!.Value);
            Tabs tabs = Assert.IsType<Tabs>(paragraphProperties.GetFirstChild<Tabs>());
            TabStop[] tabStops = tabs.Elements<TabStop>().ToArray();
            Assert.Equal(2, tabStops.Length);
            Assert.Equal(TabStopValues.Left, tabStops[0].Val!.Value);
            Assert.Equal(1440, tabStops[0].Position!.Value);
            Assert.Equal(TabStopValues.Right, tabStops[1].Val!.Value);
            Assert.Equal(TabStopLeaderCharValues.Underscore, tabStops[1].Leader!.Value);
            Assert.Equal(4320, tabStops[1].Position!.Value);
            StyleRunProperties runProperties = Assert.IsType<StyleRunProperties>(headingStyle.StyleRunProperties);
            Assert.NotNull(runProperties.GetFirstChild<Bold>());
            Assert.NotNull(runProperties.GetFirstChild<Outline>());
            Assert.NotNull(runProperties.GetFirstChild<Shadow>());
            Assert.NotNull(runProperties.GetFirstChild<Emboss>());
            Assert.NotNull(runProperties.GetFirstChild<Imprint>());
            Assert.NotNull(runProperties.GetFirstChild<Vanish>());
            Underline underline = Assert.IsType<Underline>(runProperties.GetFirstChild<Underline>());
            Assert.Equal(UnderlineValues.Single, underline.Val!.Value);
            Highlight highlight = Assert.IsType<Highlight>(runProperties.GetFirstChild<Highlight>());
            Assert.Equal(HighlightColorValues.Yellow, highlight.Val!.Value);
            Color color = Assert.IsType<Color>(runProperties.GetFirstChild<Color>());
            Assert.Equal("336699", color.Val!.Value);
            FontSize fontSize = Assert.IsType<FontSize>(runProperties.GetFirstChild<FontSize>());
            Assert.Equal("32", fontSize.Val!.Value);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsStyleInheritanceFromBuiltInStyle() {
            byte[] docBytes = LegacyDocTestBuilder.CreateUnicodeDocWithInheritedBuiltInStyleFormatting();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph paragraph = Assert.Single(result.Document.Paragraphs);
            Assert.Equal("inherited heading", paragraph.Text);
            Assert.Equal("LegacyDocInheritedHeading", paragraph.StyleId);

            Styles styles = result.Document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            Style headingStyle = Assert.Single(styles.Elements<Style>(), style => style.StyleId == WordParagraphStyles.Heading1.ToStringStyle());
            BasedOn headingBasedOn = Assert.IsType<BasedOn>(headingStyle.GetFirstChild<BasedOn>());
            Assert.Equal(WordParagraphStyles.Normal.ToStringStyle(), headingBasedOn.Val!.Value);
            StyleParagraphProperties headingParagraphProperties = Assert.IsType<StyleParagraphProperties>(headingStyle.StyleParagraphProperties);
            Justification headingJustification = Assert.IsType<Justification>(headingParagraphProperties.GetFirstChild<Justification>());
            Assert.Equal(JustificationValues.Center, headingJustification.Val!.Value);
            StyleRunProperties headingRunProperties = Assert.IsType<StyleRunProperties>(headingStyle.StyleRunProperties);
            Assert.NotNull(headingRunProperties.GetFirstChild<Bold>());
            Color headingColor = Assert.IsType<Color>(headingRunProperties.GetFirstChild<Color>());
            Assert.Equal("336699", headingColor.Val!.Value);

            Style childStyle = Assert.Single(styles.Elements<Style>(), style => style.StyleId == "LegacyDocInheritedHeading");
            BasedOn childBasedOn = Assert.IsType<BasedOn>(childStyle.GetFirstChild<BasedOn>());
            Assert.Equal(WordParagraphStyles.Heading1.ToStringStyle(), childBasedOn.Val!.Value);
            StyleRunProperties childRunProperties = Assert.IsType<StyleRunProperties>(childStyle.StyleRunProperties);
            Assert.NotNull(childRunProperties.GetFirstChild<Italic>());
        }

        [Fact]
        public void LegacyDoc_NormalLoad_RoutesOleDocIntoProjectedWordDocument() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                File.WriteAllBytes(docPath, LegacyDocTestBuilder.CreateSimpleDoc("Normal load"));

                using WordDocument document = WordDocument.Load(docPath);

                Assert.True(document.WasLoadedFromLegacyDoc);
                Assert.Equal(string.Empty, document.FilePath);
                WordParagraph paragraph = Assert.Single(document.Paragraphs);
                Assert.Equal("Normal load", paragraph.Text);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ImportsWordComDocFixture() {
            string docPath = GetFixtureDoc(Path.Combine("LegacyDocCorpus", "ComSimpleParagraphs.doc"));

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(docPath);

            result.EnsureNoImportErrors();
            Assert.True(result.HasDocument);
            Assert.True(result.Document.WasLoadedFromLegacyDoc);
            Assert.Equal(string.Empty, result.Document.FilePath);

            string[] paragraphs = result.Document.Paragraphs
                .Select(paragraph => paragraph.Text)
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .ToArray();

            Assert.Contains("First COM paragraph", paragraphs);
            Assert.Contains("Second COM paragraph", paragraphs);
        }

        [Fact]
        public void LegacyDoc_CorpusImportReports_MatchCheckedInBaselines() {
            string corpusDirectory = Path.Combine(GetWordTestsProjectRoot(), "Documents", "LegacyDocCorpus");
            string[] docPaths = Directory.GetFiles(corpusDirectory, "*.doc", SearchOption.AllDirectories)
                .Where(path => !Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal))
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToArray();

            Assert.NotEmpty(docPaths);

            bool updateBaselines = string.Equals(
                Environment.GetEnvironmentVariable("OFFICEIMO_UPDATE_LEGACY_DOC_CORPUS_BASELINES"),
                "1",
                StringComparison.Ordinal);
            var missingBaselines = new List<string>();
            foreach (string docPath in docPaths) {
                using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(docPath);
                string actual = NormalizeLegacyDocBaselineText(result.ImportReport.ToMarkdown());
                string baselinePath = Path.ChangeExtension(docPath, ".import-report.md");

                if (updateBaselines) {
                    File.WriteAllText(baselinePath, actual, Encoding.UTF8);
                    continue;
                }

                if (!File.Exists(baselinePath)) {
                    missingBaselines.Add(GetRelativePath(corpusDirectory, baselinePath));
                    continue;
                }

                string expected = NormalizeLegacyDocBaselineText(File.ReadAllText(baselinePath, Encoding.UTF8));
                Assert.Equal(expected, actual);
            }

            Assert.True(
                missingBaselines.Count == 0,
                "Missing legacy DOC corpus baselines. Run with OFFICEIMO_UPDATE_LEGACY_DOC_CORPUS_BASELINES=1 to create: "
                    + string.Join(", ", missingBaselines));
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ReportsMissingWordDocumentStream() {
            byte[] docBytes = LegacyDocTestBuilder.CreateCompoundWithoutWordDocumentStream();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            Assert.False(result.HasDocument);
            Assert.True(result.HasImportErrors);
            LegacyDocImportDiagnostic diagnostic = Assert.Single(result.Diagnostics);
            Assert.Equal("DOC-WORDDOCUMENT-MISSING", diagnostic.Code);
            Assert.Equal(LegacyDocDiagnosticSeverity.Error, diagnostic.Severity);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ReportsUnsupportedPreWord97FibVersion() {
            byte[] docBytes = LegacyDocTestBuilder.CreateSimpleDocWithFibVersion(0x0065, "Older body");

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            Assert.False(result.HasDocument);
            Assert.True(result.HasImportErrors);
            LegacyDocImportDiagnostic diagnostic = Assert.Single(result.Diagnostics);
            Assert.Equal("DOC-FIB-INVALID", diagnostic.Code);
            Assert.Equal(LegacyDocDiagnosticSeverity.Error, diagnostic.Severity);
            Assert.Contains("Unsupported Word FIB version 0x0065", diagnostic.Message);
            Assert.Equal(1, result.ImportReport.ErrorCount);
            Assert.Equal(1, result.ImportReport.DiagnosticsByCode["DOC-FIB-INVALID"]);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ReportsUnsupportedCompoundFeatures() {
            byte[] docBytes = LegacyDocTestBuilder.CreateSimpleDocWithUnsupportedFeatureStorage("Preserve-only body");

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            Assert.True(result.HasDocument);
            Assert.Equal(2, result.UnsupportedFeatures.Count);
            Assert.Equal(2, result.ImportReport.UnsupportedFeatureCount);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByKind[LegacyDocUnsupportedFeatureKind.VbaProject]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByKind[LegacyDocUnsupportedFeatureKind.OleObject]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByCode["DOC-MACROS-PRESENT"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByCode["DOC-OLE-OBJECTS-PRESENT"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByDetail["VbaProject|DOC-MACROS-PRESENT|Compound:VbaProjectStorage"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByDetail["OleObject|DOC-OLE-OBJECTS-PRESENT|Compound:OleObjectStorage"]);
            Assert.Contains(result.UnsupportedFeatures, feature => feature.EntryPath == "_VBA_PROJECT_CUR");
            Assert.Contains(result.UnsupportedFeatures, feature => feature.EntryPath == "ObjectPool");

            string markdown = result.ImportReport.ToMarkdown();
            Assert.Contains("| Unsupported features | 2 |", markdown);
            Assert.Contains("| VbaProject | DOC-MACROS-PRESENT | Compound:VbaProjectStorage | _VBA_PROJECT_CUR |", markdown);
            Assert.Contains("| OleObject | DOC-OLE-OBJECTS-PRESENT | Compound:OleObjectStorage | ObjectPool |", markdown);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ReportsUnsupportedActiveXAndEmbeddedPackageFeatures() {
            byte[] docBytes = LegacyDocTestBuilder.CreateSimpleDocWithActiveXAndEmbeddedPackageStorage("ActiveX body");

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            Assert.True(result.HasDocument);
            Assert.Equal(3, result.UnsupportedFeatures.Count);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByKind[LegacyDocUnsupportedFeatureKind.OleObject]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByKind[LegacyDocUnsupportedFeatureKind.ActiveXControl]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByKind[LegacyDocUnsupportedFeatureKind.EmbeddedPackage]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByCode["DOC-OLE-OBJECTS-PRESENT"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByCode["DOC-ACTIVEX-CONTROLS-PRESENT"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByCode["DOC-EMBEDDED-PACKAGES-PRESENT"]);
            Assert.Contains(result.UnsupportedFeatures, feature => feature.EntryPath == "ActiveX");
            Assert.Contains(result.UnsupportedFeatures, feature => feature.EntryPath == "ObjectPool");
            Assert.Contains(result.UnsupportedFeatures, feature => feature.EntryPath == "\u0001Ole10Native");

            string markdown = result.ImportReport.ToMarkdown();
            Assert.Contains("| ActiveXControl | DOC-ACTIVEX-CONTROLS-PRESENT | Compound:ActiveXControlStorage | ActiveX |", markdown);
            Assert.Contains("| EmbeddedPackage | DOC-EMBEDDED-PACKAGES-PRESENT | Compound:EmbeddedPackageStorage | \u0001Ole10Native |", markdown);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ReportsUnsupportedBinaryDataStream() {
            byte[] docBytes = LegacyDocTestBuilder.CreateSimpleDocWithDataStream("Data body");

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            Assert.True(result.HasDocument);
            LegacyDocUnsupportedFeature feature = Assert.Single(result.UnsupportedFeatures);
            Assert.Equal(LegacyDocUnsupportedFeatureKind.BinaryData, feature.Kind);
            Assert.Equal("DOC-BINARY-DATA-STREAM-PRESENT", feature.Code);
            Assert.Equal("Compound:BinaryDataStream", feature.DetailCode);
            Assert.Equal("Data", feature.EntryPath);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByKind[LegacyDocUnsupportedFeatureKind.BinaryData]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByCode["DOC-BINARY-DATA-STREAM-PRESENT"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByDetail["BinaryData|DOC-BINARY-DATA-STREAM-PRESENT|Compound:BinaryDataStream"]);

            string markdown = result.ImportReport.ToMarkdown();
            Assert.Contains("| BinaryData | DOC-BINARY-DATA-STREAM-PRESENT | Compound:BinaryDataStream | Data |", markdown);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ReportsUnsupportedFastSaveAndPictureFibFlags() {
            const ushort flags = 0x0200 | 0x0004 | 0x0008;
            byte[] docBytes = LegacyDocTestBuilder.CreateSimpleDocWithFibFlags(flags, "Fast-saved body");

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            Assert.True(result.HasDocument);
            Assert.Equal(2, result.UnsupportedFeatures.Count);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByKind[LegacyDocUnsupportedFeatureKind.FastSave]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByKind[LegacyDocUnsupportedFeatureKind.Picture]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByCode["DOC-FAST-SAVE-PRESENT"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByCode["DOC-PICTURES-PRESENT"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByDetail["FastSave|DOC-FAST-SAVE-PRESENT|Fib:FComplex"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByDetail["Picture|DOC-PICTURES-PRESENT|Fib:FHasPic"]);

            string markdown = result.ImportReport.ToMarkdown();
            Assert.Contains("| FastSave | DOC-FAST-SAVE-PRESENT | Fib:FComplex |  |", markdown);
            Assert.Contains("| Picture | DOC-PICTURES-PRESENT | Fib:FHasPic |  |", markdown);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ReportsUnsupportedQuickSaveCountFibFlag() {
            const ushort flags = 0x0200 | 0x0030;
            byte[] docBytes = LegacyDocTestBuilder.CreateSimpleDocWithFibFlags(flags, "Quick-saved body");

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            LegacyDocUnsupportedFeature feature = Assert.Single(result.UnsupportedFeatures);
            Assert.Equal(LegacyDocUnsupportedFeatureKind.FastSave, feature.Kind);
            Assert.Equal("DOC-FAST-SAVE-PRESENT", feature.Code);
            Assert.Equal("Fib:CQuickSaves", feature.DetailCode);
            Assert.Contains("3 quick-save revision", feature.Description);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByDetail["FastSave|DOC-FAST-SAVE-PRESENT|Fib:CQuickSaves"]);

            string markdown = result.ImportReport.ToMarkdown();
            Assert.Contains("| FastSave | DOC-FAST-SAVE-PRESENT | Fib:CQuickSaves |  |", markdown);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ReportsUnsupportedRevisionTrackingDopFlags() {
            const uint revisionMarkingAndLockFlags = 0x40008000;
            byte[] docBytes = LegacyDocTestBuilder.CreateSimpleDocWithRevisionTrackingDop(revisionMarkingAndLockFlags, "Tracked body");

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            Assert.True(result.HasDocument);
            Assert.Equal("Tracked body", Assert.Single(result.Document.Paragraphs).Text);
            LegacyDocUnsupportedFeature feature = Assert.Single(result.UnsupportedFeatures);
            Assert.Equal(LegacyDocUnsupportedFeatureKind.RevisionTracking, feature.Kind);
            Assert.Equal("DOC-REVISION-TRACKING-PRESENT", feature.Code);
            Assert.Equal("DopBase:FRevMarking+FLockRev", feature.DetailCode);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByKind[LegacyDocUnsupportedFeatureKind.RevisionTracking]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByCode["DOC-REVISION-TRACKING-PRESENT"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByDetail["RevisionTracking|DOC-REVISION-TRACKING-PRESENT|DopBase:FRevMarking+FLockRev"]);

            string markdown = result.ImportReport.ToMarkdown();
            Assert.Contains("| RevisionTracking | DOC-REVISION-TRACKING-PRESENT | DopBase:FRevMarking+FLockRev |  |", markdown);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ReportsUnsupportedStoryCounts() {
            byte[] docBytes = LegacyDocTestBuilder.CreateSimpleDocWithUnsupportedStoryCounts("Body story");

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            Assert.True(result.HasDocument);
            Assert.Equal("Body story", Assert.Single(result.Document.Paragraphs).Text);
            Assert.Equal(6, result.UnsupportedFeatures.Count);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByKind[LegacyDocUnsupportedFeatureKind.HeaderFooter]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByKind[LegacyDocUnsupportedFeatureKind.Footnote]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByKind[LegacyDocUnsupportedFeatureKind.Endnote]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByKind[LegacyDocUnsupportedFeatureKind.Comment]);
            Assert.Equal(2, result.ImportReport.UnsupportedFeaturesByKind[LegacyDocUnsupportedFeatureKind.TextBox]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByCode["DOC-HEADER-FOOTER-STORIES-PRESENT"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByCode["DOC-FOOTNOTE-STORIES-PRESENT"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByCode["DOC-ENDNOTE-STORIES-PRESENT"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByCode["DOC-COMMENT-STORIES-PRESENT"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByCode["DOC-TEXTBOX-STORIES-PRESENT"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByCode["DOC-HEADER-TEXTBOX-STORIES-PRESENT"]);

            string markdown = result.ImportReport.ToMarkdown();
            Assert.Contains("| Unsupported features | 6 |", markdown);
            Assert.Contains("| HeaderFooter | DOC-HEADER-FOOTER-STORIES-PRESENT | Fib:CcpHdd |  |", markdown);
            Assert.Contains("| Footnote | DOC-FOOTNOTE-STORIES-PRESENT | Fib:CcpFtn |  |", markdown);
            Assert.Contains("| Endnote | DOC-ENDNOTE-STORIES-PRESENT | Fib:CcpEdn |  |", markdown);
            Assert.Contains("| Comment | DOC-COMMENT-STORIES-PRESENT | Fib:CcpAtn |  |", markdown);
            Assert.Contains("| TextBox | DOC-TEXTBOX-STORIES-PRESENT | Fib:CcpTxbx |  |", markdown);
            Assert.Contains("| TextBox | DOC-HEADER-TEXTBOX-STORIES-PRESENT | Fib:CcpHdrTxbx |  |", markdown);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_DoesNotProjectUnsupportedStoryTextIntoBody() {
            byte[] docBytes = LegacyDocTestBuilder.CreateSimpleDocWithUnsupportedHeaderFooterStoryText("Body story", "Header leak");

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            Assert.True(result.HasDocument);
            Assert.Equal("Body story", Assert.Single(result.Document.Paragraphs).Text);
            Assert.DoesNotContain(result.Document.Paragraphs, paragraph => paragraph.Text.Contains("Header leak", StringComparison.Ordinal));
            LegacyDocUnsupportedFeature feature = Assert.Single(result.UnsupportedFeatures);
            Assert.Equal(LegacyDocUnsupportedFeatureKind.HeaderFooter, feature.Kind);
            Assert.Equal("DOC-HEADER-FOOTER-STORIES-PRESENT", feature.Code);
            Assert.Equal("Fib:CcpHdd", feature.DetailCode);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ReportsMultipleSectionsAndBlocksNativeDocResave() {
            byte[] docBytes = LegacyDocTestBuilder.CreateSimpleDocWithMultipleSectionDescriptors("Section one");
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

                result.EnsureNoImportErrors();
                LegacyDocUnsupportedFeature feature = Assert.Single(result.UnsupportedFeatures);
                Assert.Equal(LegacyDocUnsupportedFeatureKind.Section, feature.Kind);
                Assert.Equal("DOC-MULTIPLE-SECTIONS-PRESENT", feature.Code);
                Assert.Equal("Fib:PlcfSed", feature.DetailCode);
                Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByKind[LegacyDocUnsupportedFeatureKind.Section]);
                Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByCode["DOC-MULTIPLE-SECTIONS-PRESENT"]);
                Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByDetail["Section|DOC-MULTIPLE-SECTIONS-PRESENT|Fib:PlcfSed"]);
                Assert.Contains("| Section | DOC-MULTIPLE-SECTIONS-PRESENT | Fib:PlcfSed |  |", result.ImportReport.ToMarkdown());

                using WordDocument document = WordDocument.Load(new MemoryStream(docBytes));

                Assert.Contains(document.LegacyDocUnsupportedFeatures, item => item.Kind == LegacyDocUnsupportedFeatureKind.Section);
                NotSupportedException exception = Assert.Throws<NotSupportedException>(() => document.Save(docPath));
                Assert.Contains("DOC-MULTIPLE-SECTIONS-PRESENT", exception.Message);
                Assert.False(File.Exists(docPath));
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsParagraphBoundarySectionBreaks() {
            byte[] docBytes = LegacyDocTestBuilder.CreateSimpleDocWithTwoSectionPageSetup();
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

                result.EnsureNoImportErrors();
                Assert.Empty(result.UnsupportedFeatures);

                WordDocument document = result.Document;
                Assert.True(document.WasLoadedFromLegacyDoc);
                Assert.Equal(2, document.Sections.Count);
                Assert.Equal("Portrait section", Assert.Single(document.Sections[0].Paragraphs).Text);
                Assert.Equal("Landscape section", Assert.Single(document.Sections[1].Paragraphs).Text);
                Assert.Equal(PageOrientationValues.Landscape, document.Sections[1].PageOrientation);
                Assert.Equal((uint)15840, document.Sections[1].PageSettings.Width!.Value);
                Assert.Equal((uint)12240, document.Sections[1].PageSettings.Height!.Value);
                Assert.Equal(720, document.Sections[1].Margins.Top);
                Assert.Equal((uint)720, document.Sections[1].Margins.Right!.Value);
                Assert.Equal(720, document.Sections[1].Margins.Bottom);
                Assert.Equal((uint)720, document.Sections[1].Margins.Left!.Value);

                document.Save(docPath);

                using WordDocument reloaded = WordDocument.Load(docPath);
                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                Assert.Equal(2, reloaded.Sections.Count);
                Assert.Equal("Portrait section", Assert.Single(reloaded.Sections[0].Paragraphs).Text);
                Assert.Equal("Landscape section", Assert.Single(reloaded.Sections[1].Paragraphs).Text);
                Assert.Equal(PageOrientationValues.Landscape, reloaded.Sections[1].PageOrientation);
                Assert.Equal((uint)15840, reloaded.Sections[1].PageSettings.Width!.Value);
                Assert.Equal((uint)12240, reloaded.Sections[1].PageSettings.Height!.Value);
                Assert.Equal(720, reloaded.Sections[1].Margins.Top);
                Assert.Equal((uint)720, reloaded.Sections[1].Margins.Right!.Value);
                Assert.Equal(720, reloaded.Sections[1].Margins.Bottom);
                Assert.Equal((uint)720, reloaded.Sections[1].Margins.Left!.Value);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Theory]
        [InlineData(0, "continuous", "Continuous section")]
        [InlineData(1, "nextColumn", "Next-column section")]
        [InlineData(2, "nextPage", "Next-page section")]
        [InlineData(3, "evenPage", "Even-page section")]
        [InlineData(4, "oddPage", "Odd-page section")]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsParagraphBoundarySectionBreakType(int sectionBreakOperand, string expectedSectionTypeKey, string sectionText) {
            byte[] docBytes = LegacyDocTestBuilder.CreateSimpleDocWithSectionBreakKind(sectionBreakOperand, sectionText);

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            Assert.Empty(result.UnsupportedFeatures);

            WordDocument document = result.Document;
            Assert.True(document.WasLoadedFromLegacyDoc);
            Assert.Equal(2, document.Sections.Count);
            Assert.Equal("Before continuous section", Assert.Single(document.Sections[0].Paragraphs).Text);
            Assert.Equal(sectionText, Assert.Single(document.Sections[1].Paragraphs).Text);
            Assert.Equal(GetSectionMarkValue(expectedSectionTypeKey), GetParagraphSectionType(document));
        }

        [Fact]
        public void LegacyDoc_NormalLoad_ExposesUnsupportedCompoundFeaturesOnProjectedDocument() {
            byte[] docBytes = LegacyDocTestBuilder.CreateSimpleDocWithUnsupportedFeatureStorage("Normal load with unsupported features");

            using WordDocument document = WordDocument.Load(new MemoryStream(docBytes));

            Assert.True(document.WasLoadedFromLegacyDoc);
            Assert.Equal(2, document.LegacyDocUnsupportedFeatures.Count);
            Assert.Contains(document.LegacyDocUnsupportedFeatures, feature => feature.Kind == LegacyDocUnsupportedFeatureKind.VbaProject);
            Assert.Contains(document.LegacyDocUnsupportedFeatures, feature => feature.Kind == LegacyDocUnsupportedFeatureKind.OleObject);
            Assert.Contains(document.LegacyDocImportDiagnostics, diagnostic => diagnostic.Code == "DOC-MACROS-PRESENT");
            Assert.Contains(document.LegacyDocImportDiagnostics, diagnostic => diagnostic.Code == "DOC-OLE-OBJECTS-PRESENT");
        }

        [Fact]
        public void LegacyDoc_NormalLoad_BlocksAutoSaveForLegacyDocProjection() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                File.WriteAllBytes(docPath, LegacyDocTestBuilder.CreateSimpleDoc("No autosave"));

                NotSupportedException exception = Assert.Throws<NotSupportedException>(() => WordDocument.Load(docPath, autoSave: true));

                Assert.Contains("Auto-save is not supported", exception.Message);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    document.AddParagraph("Zażółć gęślą jaźń");
                    document.AddParagraph("Second plain paragraph");

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                Assert.Equal(string.Empty, reloaded.FilePath);
                Assert.Empty(reloaded.LegacyDocUnsupportedFeatures);
                string[] paragraphs = reloaded.Paragraphs
                    .Select(paragraph => paragraph.Text)
                    .Where(text => !string.IsNullOrEmpty(text))
                    .ToArray();
                Assert.Equal(new[] { "Zażółć gęślą jaźń", "Second plain paragraph" }, paragraphs);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveStreamWithLegacyDocFormat_WritesNativeDocAndReloadsThroughLegacyReader() {
            using var stream = new MemoryStream();
            using (WordDocument document = WordDocument.Create()) {
                document.AddParagraph("Native DOC stream");
                document.Save(stream, new WordSaveOptions {
                    StreamFormat = WordStreamSaveFormat.LegacyDoc
                });
            }

            byte[] bytes = stream.ToArray();
            Assert.True(bytes.Length > 512);
            Assert.Equal(0xD0, bytes[0]);
            Assert.Equal(0xCF, bytes[1]);
            Assert.Equal(0x11, bytes[2]);
            Assert.Equal(0xE0, bytes[3]);

            stream.Seek(0, SeekOrigin.Begin);
            using WordDocument reloaded = WordDocument.Load(stream);

            Assert.True(reloaded.WasLoadedFromLegacyDoc);
            WordParagraph paragraph = Assert.Single(reloaded.Paragraphs, paragraph => !string.IsNullOrEmpty(paragraph.Text));
            Assert.Equal("Native DOC stream", paragraph.Text);
        }

        [Fact]
        public void LegacyDoc_SaveStreamWithDefaultFormat_KeepsOpenXmlStreamSave() {
            using var stream = new MemoryStream();
            using (WordDocument document = WordDocument.Create()) {
                document.AddParagraph("Default stream format");
                document.Save(stream, WordSaveOptions.None);
            }

            byte[] bytes = stream.ToArray();
            Assert.True(bytes.Length > 4);
            Assert.Equal((byte)'P', bytes[0]);
            Assert.Equal((byte)'K', bytes[1]);

            stream.Seek(0, SeekOrigin.Begin);
            using WordDocument reloaded = WordDocument.Load(stream);

            Assert.False(reloaded.WasLoadedFromLegacyDoc);
            WordParagraph paragraph = Assert.Single(reloaded.Paragraphs, paragraph => !string.IsNullOrEmpty(paragraph.Text));
            Assert.Equal("Default stream format", paragraph.Text);
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocPropertiesAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");
            DateTime created = new DateTime(2026, 6, 29, 10, 0, 0, DateTimeKind.Utc);
            DateTime modified = new DateTime(2026, 6, 29, 10, 30, 0, DateTimeKind.Utc);
            DateTime reviewedAt = new DateTime(2026, 6, 29, 11, 0, 0, DateTimeKind.Utc);

            try {
                using (WordDocument document = WordDocument.Create()) {
                    document.AddParagraph("Metadata native DOC");
                    document.BuiltinDocumentProperties.Title = "Native DOC Metadata Title";
                    document.BuiltinDocumentProperties.Subject = "Native DOC metadata subject";
                    document.BuiltinDocumentProperties.Creator = "OfficeIMO Native DOC";
                    document.BuiltinDocumentProperties.Keywords = "doc, metadata, native";
                    document.BuiltinDocumentProperties.Description = "Native DOC metadata comments";
                    document.BuiltinDocumentProperties.Category = "Native Category";
                    document.BuiltinDocumentProperties.Created = created;
                    document.BuiltinDocumentProperties.Modified = modified;
                    document.ApplicationProperties.Company = "EvotecIT";
                    document.ApplicationProperties.Manager = new Manager { Text = "Native Manager" };
                    document.CustomDocumentProperties["ReleaseStatus"] = new WordCustomProperty("Ready");
                    document.CustomDocumentProperties["Reviewed"] = new WordCustomProperty(true);
                    document.CustomDocumentProperties["Ticket"] = new WordCustomProperty(2004);
                    document.CustomDocumentProperties["Score"] = new WordCustomProperty(98.5d);
                    document.CustomDocumentProperties["ReviewedAt"] = new WordCustomProperty(reviewedAt);

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                Assert.Equal("Native DOC Metadata Title", reloaded.BuiltinDocumentProperties.Title);
                Assert.Equal("Native DOC metadata subject", reloaded.BuiltinDocumentProperties.Subject);
                Assert.Equal("OfficeIMO Native DOC", reloaded.BuiltinDocumentProperties.Creator);
                Assert.Equal("doc, metadata, native", reloaded.BuiltinDocumentProperties.Keywords);
                Assert.Equal("Native DOC metadata comments", reloaded.BuiltinDocumentProperties.Description);
                Assert.Equal("Native Category", reloaded.BuiltinDocumentProperties.Category);
                AssertSameInstant(created, reloaded.BuiltinDocumentProperties.Created);
                AssertSameInstant(modified, reloaded.BuiltinDocumentProperties.Modified);
                Assert.Equal("EvotecIT", reloaded.ApplicationProperties.Company);
                Assert.Equal("Native Manager", reloaded.ApplicationProperties.Manager?.Text);
                Assert.Equal("Ready", reloaded.CustomDocumentProperties["ReleaseStatus"].Text);
                Assert.True(reloaded.CustomDocumentProperties["Reviewed"].Bool);
                Assert.Equal(2004, reloaded.CustomDocumentProperties["Ticket"].NumberInteger);
                Assert.Equal(98.5d, reloaded.CustomDocumentProperties["Score"].NumberDouble);
                AssertSameInstant(reviewedAt, reloaded.CustomDocumentProperties["ReviewedAt"].Date);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocBoldItalicRunsAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    WordParagraph paragraph = document.AddParagraph();
                    paragraph.AddText("plain ");
                    paragraph.AddText("bold ").SetBold();
                    paragraph.AddText("italic").SetItalic();

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                WordParagraph[] runs = reloaded.Paragraphs.ToArray();
                Assert.Equal(3, runs.Length);
                Assert.Equal("plain ", runs[0].Text);
                Assert.False(runs[0].Bold);
                Assert.False(runs[0].Italic);
                Assert.Equal("bold ", runs[1].Text);
                Assert.True(runs[1].Bold);
                Assert.False(runs[1].Italic);
                Assert.Equal("italic", runs[2].Text);
                Assert.False(runs[2].Bold);
                Assert.True(runs[2].Italic);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocUnderlineSizeColorStrikeVerticalAndHighlightRunsAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    WordParagraph paragraph = document.AddParagraph();
                    paragraph.AddText("plain ");
                    paragraph.AddText("under ").SetUnderline(UnderlineValues.Single);
                    paragraph.AddText("sized ").SetFontSize(14);
                    paragraph.AddText("strike ").SetStrike();
                    paragraph.AddText("double ").SetDoubleStrike();
                    paragraph.AddText("outline ").SetOutline();
                    paragraph.AddText("shadow ").SetShadow();
                    paragraph.AddText("emboss ").SetEmboss();
                    WordParagraph imprint = paragraph.AddText("imprint ");
                    imprint._run!.RunProperties ??= new RunProperties();
                    imprint._run.RunProperties.Imprint = new Imprint();
                    WordParagraph hidden = paragraph.AddText("hidden ");
                    hidden._run!.RunProperties ??= new RunProperties();
                    hidden._run.RunProperties.Vanish = new Vanish();
                    paragraph.AddText("caps ").SetCapsStyle(CapsStyle.Caps);
                    paragraph.AddText("small ").SetSmallCaps();
                    paragraph.AddText("super ").SetSuperScript();
                    paragraph.AddText("sub ").SetSubScript();
                    paragraph.AddText("mark ").SetHighlight(HighlightColorValues.Yellow);
                    paragraph.AddText("color").SetColorHex("336699");

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                WordParagraph[] runs = reloaded.Paragraphs.ToArray();
                Assert.Equal(16, runs.Length);
                Assert.Equal("plain ", runs[0].Text);
                Assert.Null(runs[0].Underline);
                Assert.False(runs[0].Strike);
                Assert.False(runs[0].DoubleStrike);
                Assert.False(runs[0].Outline);
                Assert.False(runs[0].Shadow);
                Assert.False(runs[0].Emboss);
                Assert.Null(runs[0]._runProperties?.Imprint);
                Assert.Null(runs[0]._runProperties?.Vanish);
                Assert.Equal(CapsStyle.None, runs[0].CapsStyle);
                Assert.Null(runs[0].VerticalTextAlignment);
                Assert.Null(runs[0].Highlight);
                Assert.Equal("under ", runs[1].Text);
                Assert.Equal(UnderlineValues.Single, runs[1].Underline);
                Assert.Equal("sized ", runs[2].Text);
                Assert.Equal(14, runs[2].FontSize);
                Assert.Equal("strike ", runs[3].Text);
                Assert.True(runs[3].Strike);
                Assert.Equal("double ", runs[4].Text);
                Assert.True(runs[4].DoubleStrike);
                Assert.Equal("outline ", runs[5].Text);
                Assert.True(runs[5].Outline);
                Assert.Equal("shadow ", runs[6].Text);
                Assert.True(runs[6].Shadow);
                Assert.Equal("emboss ", runs[7].Text);
                Assert.True(runs[7].Emboss);
                Assert.Equal("imprint ", runs[8].Text);
                Assert.NotNull(runs[8]._runProperties?.Imprint);
                Assert.Equal("hidden ", runs[9].Text);
                Assert.NotNull(runs[9]._runProperties?.Vanish);
                Assert.Equal("caps ", runs[10].Text);
                Assert.Equal(CapsStyle.Caps, runs[10].CapsStyle);
                Assert.Equal("small ", runs[11].Text);
                Assert.Equal(CapsStyle.SmallCaps, runs[11].CapsStyle);
                Assert.Equal("super ", runs[12].Text);
                Assert.Equal(VerticalPositionValues.Superscript, runs[12].VerticalTextAlignment);
                Assert.Equal("sub ", runs[13].Text);
                Assert.Equal(VerticalPositionValues.Subscript, runs[13].VerticalTextAlignment);
                Assert.Equal("mark ", runs[14].Text);
                Assert.Equal(HighlightColorValues.Yellow, runs[14].Highlight);
                Assert.Equal("color", runs[15].Text);
                Assert.Equal("336699", runs[15].ColorHex);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocFontFamilyRunsAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    WordParagraph paragraph = document.AddParagraph();
                    paragraph.AddText("plain ");
                    paragraph.AddText("font").SetFontFamily("Courier New");

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                WordParagraph[] runs = reloaded.Paragraphs.ToArray();
                Assert.Equal(2, runs.Length);
                Assert.Equal("plain ", runs[0].Text);
                Assert.Null(runs[0].FontFamily);
                Assert.Equal("font", runs[1].Text);
                Assert.Equal("Courier New", runs[1].FontFamily);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocTabsAndReloadsAsWordTabRuns() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    WordParagraph paragraph = document.AddParagraph();
                    paragraph.AddText("Left");
                    paragraph.AddTab();
                    paragraph.AddText("Right");

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                Paragraph reloadedParagraph = Assert.Single(reloaded._wordprocessingDocument!.MainDocumentPart!.Document.Body!.Elements<Paragraph>());
                Assert.Equal(1, reloadedParagraph.Descendants<TabChar>().Count());
                Assert.DoesNotContain(reloadedParagraph.Descendants<Text>(), text => text.Text.Contains('\t'));
                Assert.Equal(new[] { "Left", "Right" }, reloadedParagraph.Descendants<Text>().Select(text => text.Text).ToArray());
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocBreaksAndReloadsAsWordBreakRuns() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    WordParagraph paragraph = document.AddParagraph();
                    paragraph.AddText("Line");
                    paragraph.AddBreak();
                    paragraph.AddText("Break");
                    paragraph.AddBreak(BreakValues.Page);
                    paragraph.AddText("Page");

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                Paragraph reloadedParagraph = Assert.Single(reloaded._wordprocessingDocument!.MainDocumentPart!.Document.Body!.Elements<Paragraph>());
                Break[] breaks = reloadedParagraph.Descendants<Break>().ToArray();
                Assert.Equal(2, breaks.Length);
                Assert.Null(breaks[0].Type);
                Assert.Equal(BreakValues.Page, breaks[1].Type!.Value);
                Assert.DoesNotContain(reloadedParagraph.Descendants<Text>(), text => text.Text.Contains('\v') || text.Text.Contains('\f'));
                Assert.Equal(new[] { "Line", "Break", "Page" }, reloadedParagraph.Descendants<Text>().Select(text => text.Text).ToArray());
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocParagraphAlignmentAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    document.AddParagraph("left");
                    document.AddParagraph("center").ParagraphAlignment = JustificationValues.Center;
                    document.AddParagraph("right").ParagraphAlignment = JustificationValues.Right;

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                WordParagraph[] paragraphs = reloaded.Paragraphs.ToArray();
                Assert.Equal(3, paragraphs.Length);
                Assert.Equal("left", paragraphs[0].Text);
                Assert.Null(paragraphs[0].ParagraphAlignment);
                Assert.Equal("center", paragraphs[1].Text);
                Assert.Equal(JustificationValues.Center, paragraphs[1].ParagraphAlignment);
                Assert.Equal("right", paragraphs[2].Text);
                Assert.Equal(JustificationValues.Right, paragraphs[2].ParagraphAlignment);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocParagraphSpacingAndIndentationAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    document.AddParagraph("plain");
                    WordParagraph formatted = document.AddParagraph("formatted");
                    formatted.LineSpacingBefore = 240;
                    formatted.LineSpacingAfter = 120;
                    formatted.LineSpacing = 360;
                    formatted.IndentationBefore = 720;
                    formatted.IndentationAfter = 360;
                    formatted.IndentationFirstLine = 240;
                    WordParagraph hanging = document.AddParagraph("hanging");
                    hanging.IndentationBefore = 720;
                    hanging.IndentationHanging = 360;

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                WordParagraph[] paragraphs = reloaded.Paragraphs.ToArray();
                Assert.Equal(3, paragraphs.Length);
                Assert.Equal("plain", paragraphs[0].Text);
                Assert.Null(paragraphs[0].LineSpacingBefore);
                Assert.Null(paragraphs[0].IndentationBefore);
                Assert.Equal("formatted", paragraphs[1].Text);
                Assert.Equal(240, paragraphs[1].LineSpacingBefore);
                Assert.Equal(120, paragraphs[1].LineSpacingAfter);
                Assert.Equal(360, paragraphs[1].LineSpacing);
                Assert.Equal(720, paragraphs[1].IndentationBefore);
                Assert.Equal(360, paragraphs[1].IndentationAfter);
                Assert.Equal(240, paragraphs[1].IndentationFirstLine);
                Assert.Equal("hanging", paragraphs[2].Text);
                Assert.Equal(720, paragraphs[2].IndentationBefore);
                Assert.Equal(360, paragraphs[2].IndentationHanging);
                Assert.Null(paragraphs[2].IndentationFirstLine);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocParagraphTabStopsAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    document.AddParagraph("plain");
                    WordParagraph paragraph = document.AddParagraph("tabs");
                    paragraph.AddTabStop(1440, TabStopValues.Left, TabStopLeaderCharValues.None);
                    paragraph.AddTabStop(2880, TabStopValues.Decimal, TabStopLeaderCharValues.Dot);
                    paragraph.AddTabStop(4320, TabStopValues.Right, TabStopLeaderCharValues.Underscore);
                    WordParagraph clear = document.AddParagraph("clear");
                    clear.AddTabStop(1440, TabStopValues.Clear, TabStopLeaderCharValues.None);
                    clear.AddTabStop(2160, TabStopValues.Bar, TabStopLeaderCharValues.None);

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                WordParagraph[] paragraphs = reloaded.Paragraphs.ToArray();
                Assert.Equal(3, paragraphs.Length);
                Assert.Equal("plain", paragraphs[0].Text);
                Assert.Empty(paragraphs[0].TabStops);
                Assert.Equal("tabs", paragraphs[1].Text);
                Assert.Equal(3, paragraphs[1].TabStops.Count);
                Assert.Equal(1440, paragraphs[1].TabStops[0].Position);
                Assert.Equal(TabStopValues.Left, paragraphs[1].TabStops[0].Alignment);
                Assert.Equal(TabStopLeaderCharValues.None, paragraphs[1].TabStops[0].Leader);
                Assert.Equal(2880, paragraphs[1].TabStops[1].Position);
                Assert.Equal(TabStopValues.Decimal, paragraphs[1].TabStops[1].Alignment);
                Assert.Equal(TabStopLeaderCharValues.Dot, paragraphs[1].TabStops[1].Leader);
                Assert.Equal(4320, paragraphs[1].TabStops[2].Position);
                Assert.Equal(TabStopValues.Right, paragraphs[1].TabStops[2].Alignment);
                Assert.Equal(TabStopLeaderCharValues.Underscore, paragraphs[1].TabStops[2].Leader);
                Assert.Equal("clear", paragraphs[2].Text);
                Assert.Equal(2, paragraphs[2].TabStops.Count);
                Assert.Equal(1440, paragraphs[2].TabStops[0].Position);
                Assert.Equal(TabStopValues.Clear, paragraphs[2].TabStops[0].Alignment);
                Assert.Equal(2160, paragraphs[2].TabStops[1].Position);
                Assert.Equal(TabStopValues.Bar, paragraphs[2].TabStops[1].Alignment);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocParagraphPaginationFlagsAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    document.AddParagraph("plain");
                    WordParagraph formatted = document.AddParagraph("pagination");
                    formatted.KeepLinesTogether = true;
                    formatted.KeepWithNext = true;
                    formatted.PageBreakBefore = true;
                    formatted.AvoidWidowAndOrphan = true;

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                WordParagraph[] paragraphs = reloaded.Paragraphs.ToArray();
                Assert.Equal(2, paragraphs.Length);
                Assert.Equal("plain", paragraphs[0].Text);
                Assert.False(paragraphs[0].KeepLinesTogether);
                Assert.False(paragraphs[0].KeepWithNext);
                Assert.False(paragraphs[0].PageBreakBefore);
                Assert.False(paragraphs[0].AvoidWidowAndOrphan);
                Assert.Equal("pagination", paragraphs[1].Text);
                Assert.True(paragraphs[1].KeepLinesTogether);
                Assert.True(paragraphs[1].KeepWithNext);
                Assert.True(paragraphs[1].PageBreakBefore);
                Assert.True(paragraphs[1].AvoidWidowAndOrphan);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocSimpleTableAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    WordTable table = document.AddTable(2, 2);
                    table.Rows[0].Cells[0].AddParagraph("A1", removeExistingParagraphs: true);
                    table.Rows[0].Cells[1].AddParagraph("B1", removeExistingParagraphs: true);
                    table.Rows[1].Cells[0].AddParagraph("A2", removeExistingParagraphs: true);
                    table.Rows[1].Cells[1].AddParagraph("B2", removeExistingParagraphs: true);

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                WordTable reloadedTable = Assert.Single(reloaded.Tables);
                Assert.Equal(2, reloadedTable.Rows.Count);
                Assert.Equal(2, reloadedTable.Rows[0].Cells.Count);
                Assert.Equal("A1", reloadedTable.Rows[0].Cells[0].Paragraphs[0].Text);
                Assert.Equal("B1", reloadedTable.Rows[0].Cells[1].Paragraphs[0].Text);
                Assert.Equal("A2", reloadedTable.Rows[1].Cells[0].Paragraphs[0].Text);
                Assert.Equal("B2", reloadedTable.Rows[1].Cells[1].Paragraphs[0].Text);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocFormattedTableCellRunsAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    WordTable table = document.AddTable(1, 2);
                    WordParagraph firstCellParagraph = table.Rows[0].Cells[0].AddParagraph(removeExistingParagraphs: true);
                    firstCellParagraph.AddText("A1").SetBold();
                    table.Rows[0].Cells[1].AddParagraph("B1", removeExistingParagraphs: true);

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                WordTable reloadedTable = Assert.Single(reloaded.Tables);
                WordParagraph firstCellRun = reloadedTable.Rows[0].Cells[0].Paragraphs[0];
                WordParagraph secondCellRun = reloadedTable.Rows[0].Cells[1].Paragraphs[0];
                Assert.Equal("A1", firstCellRun.Text);
                Assert.True(firstCellRun.Bold);
                Assert.Equal("B1", secondCellRun.Text);
                Assert.False(secondCellRun.Bold);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocTableCellParagraphFormattingAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    WordTable table = document.AddTable(1, 2);
                    WordParagraph formatted = table.Rows[0].Cells[0].AddParagraph("Centered", removeExistingParagraphs: true);
                    formatted.ParagraphAlignment = JustificationValues.Center;
                    formatted.LineSpacingAfter = 120;
                    formatted.IndentationBefore = 360;
                    table.Rows[0].Cells[1].AddParagraph("Plain", removeExistingParagraphs: true);

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                WordTable reloadedTable = Assert.Single(reloaded.Tables);
                WordParagraph formattedCellParagraph = reloadedTable.Rows[0].Cells[0].Paragraphs[0];
                WordParagraph plainCellParagraph = reloadedTable.Rows[0].Cells[1].Paragraphs[0];
                Assert.Equal("Centered", formattedCellParagraph.Text);
                Assert.Equal(JustificationValues.Center, formattedCellParagraph.ParagraphAlignment);
                Assert.Equal(120, formattedCellParagraph.LineSpacingAfter);
                Assert.Equal(360, formattedCellParagraph.IndentationBefore);
                Assert.Equal("Plain", plainCellParagraph.Text);
                Assert.Null(plainCellParagraph.ParagraphAlignment);
                Assert.Null(plainCellParagraph.LineSpacingAfter);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocTableMarkerParagraphFlags() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    WordTable table = document.AddTable(1, 2);
                    table.Rows[0].Cells[0].AddParagraph("A1", removeExistingParagraphs: true);
                    table.Rows[0].Cells[1].AddParagraph(string.Empty, removeExistingParagraphs: true);

                    document.Save(docPath);
                }

                byte[] wordDocumentStream = ReadCompoundStream(File.ReadAllBytes(docPath), "WordDocument");

                Assert.True(
                    ContainsBytePattern(wordDocumentStream, 0x16, 0x24, 0x01),
                    "Expected the native DOC paragraph property stream to contain sprmPFInTable.");
                Assert.True(
                    ContainsBytePattern(wordDocumentStream, 0x17, 0x24, 0x01),
                    "Expected the native DOC paragraph property stream to contain sprmPFTtp.");

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                WordTable reloadedTable = Assert.Single(reloaded.Tables);
                WordTableRow row = Assert.Single(reloadedTable.Rows);
                Assert.Equal(2, row.Cells.Count);
                Assert.Equal("A1", row.Cells[0].Paragraphs[0].Text);
                Assert.Equal(string.Empty, row.Cells[1].Paragraphs[0].Text);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocTableCellWidthsAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    WordTable table = document.AddTable(1, 2);
                    table.Rows[0].Cells[0].WidthType = TableWidthUnitValues.Dxa;
                    table.Rows[0].Cells[0].Width = 1440;
                    table.Rows[0].Cells[0].AddParagraph("Narrow", removeExistingParagraphs: true);
                    table.Rows[0].Cells[1].WidthType = TableWidthUnitValues.Dxa;
                    table.Rows[0].Cells[1].Width = 2880;
                    table.Rows[0].Cells[1].AddParagraph("Wide", removeExistingParagraphs: true);

                    document.Save(docPath);
                }

                byte[] wordDocumentStream = ReadCompoundStream(File.ReadAllBytes(docPath), "WordDocument");
                Assert.True(
                    ContainsBytePattern(wordDocumentStream, 0x08, 0xD6),
                    "Expected the native DOC paragraph property stream to contain sprmTDefTable.");

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                WordTable reloadedTable = Assert.Single(reloaded.Tables);
                WordTableRow row = Assert.Single(reloadedTable.Rows);
                Assert.Equal("Narrow", row.Cells[0].Paragraphs[0].Text);
                Assert.Equal(1440, row.Cells[0].Width);
                Assert.Equal(TableWidthUnitValues.Dxa, row.Cells[0].WidthType);
                Assert.Equal("Wide", row.Cells[1].Paragraphs[0].Text);
                Assert.Equal(2880, row.Cells[1].Width);
                Assert.Equal(TableWidthUnitValues.Dxa, row.Cells[1].WidthType);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_BlocksUnsupportedMergedTablesBeforeCreatingFile() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using WordDocument document = WordDocument.Create();
                WordTable table = document.AddTable(1, 2);
                table.Rows[0].Cells[0].AddParagraph("Merged", removeExistingParagraphs: true);
                table.Rows[0].Cells[0].MergeHorizontally(1);

                NotSupportedException exception = Assert.Throws<NotSupportedException>(() => document.Save(docPath));

                Assert.Contains("table cell property", exception.Message.ToLowerInvariant());
                Assert.False(File.Exists(docPath));
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_BlocksUnsupportedMultiParagraphTableCellsBeforeCreatingFile() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using WordDocument document = WordDocument.Create();
                WordTable table = document.AddTable(1, 1);
                table.Rows[0].Cells[0].AddParagraph("First", removeExistingParagraphs: true);
                table.Rows[0].Cells[0].AddParagraph("Second");

                NotSupportedException exception = Assert.Throws<NotSupportedException>(() => document.Save(docPath));

                Assert.Contains("one paragraph per cell", exception.Message.ToLowerInvariant());
                Assert.False(File.Exists(docPath));
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_BlocksUnsupportedNestedTablesBeforeCreatingFile() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using WordDocument document = WordDocument.Create();
                WordTable table = document.AddTable(1, 1);
                WordTable nestedTable = table.Rows[0].Cells[0].AddTable(1, 1);
                nestedTable.Rows[0].Cells[0].AddParagraph("Nested", removeExistingParagraphs: true);

                NotSupportedException exception = Assert.Throws<NotSupportedException>(() => document.Save(docPath));

                Assert.Contains("nested tables", exception.Message.ToLowerInvariant());
                Assert.False(File.Exists(docPath));
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_BlocksUnsupportedRunFormattingBeforeCreatingFile() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using WordDocument document = WordDocument.Create();
                WordParagraph formatted = document.AddParagraph("Formatted");
                formatted._run!.RunProperties ??= new RunProperties();
                formatted._run.RunProperties.Languages = new Languages { Val = "en-US" };

                NotSupportedException exception = Assert.Throws<NotSupportedException>(() => document.Save(docPath));

                Assert.Contains("unsupported run property", exception.Message.ToLowerInvariant());
                Assert.False(File.Exists(docPath));
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_BlocksRevisionTrackingSettingsBeforeCreatingFile() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using WordDocument document = WordDocument.Create();
                document.AddParagraph("Tracked settings");
                document.Settings.TrackRevisions = true;

                NotSupportedException exception = Assert.Throws<NotSupportedException>(() => document.Save(docPath));

                Assert.Contains("revision tracking", exception.Message.ToLowerInvariant());
                Assert.False(File.Exists(docPath));
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_BlocksTrackedRevisionMarkupBeforeCreatingFile() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using WordDocument document = WordDocument.Create();
                WordParagraph paragraph = document.AddParagraph("Review ");
                paragraph.AddInsertedText("inserted", "OfficeIMO");
                paragraph.AddDeletedText("deleted", "OfficeIMO");

                NotSupportedException exception = Assert.Throws<NotSupportedException>(() => document.Save(docPath));

                Assert.Contains("tracked revision markup", exception.Message.ToLowerInvariant());
                Assert.False(File.Exists(docPath));
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_BlocksCommentsBeforeCreatingFile() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using WordDocument document = WordDocument.Create();
                WordParagraph paragraph = document.AddParagraph("Commented");
                paragraph.AddComment("OfficeIMO", "OI", "Native DOC comments are not supported yet.");

                NotSupportedException exception = Assert.Throws<NotSupportedException>(() => document.Save(docPath));

                Assert.Contains("comments", exception.Message.ToLowerInvariant());
                Assert.False(File.Exists(docPath));
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocSectionPageSetupAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    document.AddParagraph("Landscape section");
                    document.PageSettings.PageSize = WordPageSize.Letter;
                    document.PageOrientation = PageOrientationValues.Landscape;
                    document.Sections[0].SetMargins(WordMargin.Narrow);
                    document.Margins.HeaderDistance = (DocumentFormat.OpenXml.UInt32Value)540U;
                    document.Margins.FooterDistance = (DocumentFormat.OpenXml.UInt32Value)900U;
                    document.Margins.Gutter = (DocumentFormat.OpenXml.UInt32Value)360U;

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                Assert.Equal("Landscape section", Assert.Single(reloaded.Paragraphs).Text);
                Assert.Equal(PageOrientationValues.Landscape, reloaded.PageOrientation);
                Assert.Equal((uint)15840, reloaded.PageSettings.Width?.Value);
                Assert.Equal((uint)12240, reloaded.PageSettings.Height?.Value);
                Assert.Equal(720, reloaded.Margins.Top);
                Assert.Equal((uint)720, reloaded.Margins.Right.Value);
                Assert.Equal(720, reloaded.Margins.Bottom);
                Assert.Equal((uint)720, reloaded.Margins.Left.Value);
                Assert.Equal((uint)540, reloaded.Margins.HeaderDistance.Value);
                Assert.Equal((uint)900, reloaded.Margins.FooterDistance.Value);
                Assert.Equal((uint)360, reloaded.Margins.Gutter.Value);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocMultipleSectionsAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    document.AddParagraph("Portrait section");
                    WordSection secondSection = document.AddSection(SectionMarkValues.NextPage);
                    secondSection.PageSettings.PageSize = WordPageSize.Letter;
                    secondSection.PageOrientation = PageOrientationValues.Landscape;
                    secondSection.SetMargins(WordMargin.Narrow);
                    secondSection.AddParagraph("Landscape section");

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                Assert.Equal(2, reloaded.Sections.Count);
                Assert.Equal("Portrait section", Assert.Single(reloaded.Sections[0].Paragraphs).Text);
                Assert.Equal("Landscape section", Assert.Single(reloaded.Sections[1].Paragraphs).Text);
                Assert.Equal(PageOrientationValues.Landscape, reloaded.Sections[1].PageOrientation);
                Assert.Equal((uint)15840, reloaded.Sections[1].PageSettings.Width!.Value);
                Assert.Equal((uint)12240, reloaded.Sections[1].PageSettings.Height!.Value);
                Assert.Equal(720, reloaded.Sections[1].Margins.Top);
                Assert.Equal((uint)720, reloaded.Sections[1].Margins.Right!.Value);
                Assert.Equal(720, reloaded.Sections[1].Margins.Bottom);
                Assert.Equal((uint)720, reloaded.Sections[1].Margins.Left!.Value);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocSectionBreakAfterTableAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    WordTable table = document.AddTable(1, 2);
                    table.Rows[0].Cells[0].AddParagraph("A1", removeExistingParagraphs: true);
                    table.Rows[0].Cells[1].AddParagraph("B1", removeExistingParagraphs: true);

                    WordSection secondSection = document.AddSection(SectionMarkValues.NextPage);
                    secondSection.PageSettings.PageSize = WordPageSize.Letter;
                    secondSection.PageOrientation = PageOrientationValues.Landscape;
                    secondSection.SetMargins(WordMargin.Narrow);
                    secondSection.AddParagraph("Landscape after table");

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                Assert.Equal(2, reloaded.Sections.Count);
                WordTable reloadedTable = Assert.Single(reloaded.Sections[0].Tables);
                WordTableRow row = Assert.Single(reloadedTable.Rows);
                Assert.Equal("A1", row.Cells[0].Paragraphs[0].Text);
                Assert.Equal("B1", row.Cells[1].Paragraphs[0].Text);
                Assert.Equal("Landscape after table", Assert.Single(reloaded.Sections[1].Paragraphs).Text);
                Assert.Equal(PageOrientationValues.Landscape, reloaded.Sections[1].PageOrientation);
                Assert.Equal((uint)15840, reloaded.Sections[1].PageSettings.Width!.Value);
                Assert.Equal((uint)12240, reloaded.Sections[1].PageSettings.Height!.Value);
                Assert.Equal(720, reloaded.Sections[1].Margins.Top);
                Assert.Equal((uint)720, reloaded.Sections[1].Margins.Right!.Value);
                Assert.Equal(720, reloaded.Sections[1].Margins.Bottom);
                Assert.Equal((uint)720, reloaded.Sections[1].Margins.Left!.Value);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Theory]
        [InlineData("continuous", "Continuous section")]
        [InlineData("nextColumn", "Next-column section")]
        [InlineData("nextPage", "Next-page section")]
        [InlineData("evenPage", "Even-page section")]
        [InlineData("oddPage", "Odd-page section")]
        public void LegacyDoc_SaveDocPath_WritesNativeDocSectionBreakTypeAndReloadsThroughLegacyReader(string sectionBreakTypeKey, string sectionText) {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");
            SectionMarkValues sectionBreakType = GetSectionMarkValue(sectionBreakTypeKey);

            try {
                using (WordDocument document = WordDocument.Create()) {
                    document.AddParagraph("Before continuous section");
                    WordSection secondSection = document.AddSection(sectionBreakType);
                    secondSection.AddParagraph(sectionText);

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                Assert.Equal(2, reloaded.Sections.Count);
                Assert.Equal("Before continuous section", Assert.Single(reloaded.Sections[0].Paragraphs).Text);
                Assert.Equal(sectionText, Assert.Single(reloaded.Sections[1].Paragraphs).Text);
                Assert.Equal(sectionBreakType, GetParagraphSectionType(reloaded));
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_BlocksUnsupportedSectionColumnsBeforeCreatingFile() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using WordDocument document = WordDocument.Create();
                document.AddParagraph("Columns");
                document.Sections[0].ColumnCount = 2;

                NotSupportedException exception = Assert.Throws<NotSupportedException>(() => document.Save(docPath));

                Assert.Contains("section property", exception.Message.ToLowerInvariant());
                Assert.False(File.Exists(docPath));
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_BlocksNativeDocSaveWhenImportedLegacyDocHasUnsupportedFeaturesBeforeCreatingFile() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using WordDocument document = WordDocument.Load(new MemoryStream(LegacyDocTestBuilder.CreateSimpleDocWithDataStream("Blocked")));

                NotSupportedException exception = Assert.Throws<NotSupportedException>(() => document.Save(docPath));

                Assert.Contains("imported from a legacy DOC", exception.Message);
                Assert.Contains("DOC-BINARY-DATA-STREAM-PRESENT", exception.Message);
                Assert.False(File.Exists(docPath));
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveStream_BlocksNativeDocSaveWhenImportedLegacyDocHasUnsupportedFeaturesBeforeWritingStream() {
            using WordDocument document = WordDocument.Load(new MemoryStream(LegacyDocTestBuilder.CreateSimpleDocWithDataStream("Blocked")));
            using var output = new MemoryStream(new byte[] { 1, 2, 3, 4 }, writable: true);

            NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
                document.Save(output, new WordSaveOptions {
                    StreamFormat = WordStreamSaveFormat.LegacyDoc
                }));

            Assert.Contains("imported from a legacy DOC", exception.Message);
            Assert.Contains("DOC-BINARY-DATA-STREAM-PRESENT", exception.Message);
            Assert.Equal(new byte[] { 1, 2, 3, 4 }, output.ToArray());
        }

        private static class LegacyDocTestBuilder {
            internal static byte[] CreateSimpleDoc(params string[] paragraphs) {
                string text = string.Join("\r", paragraphs) + "\r";
                byte[] wordDocumentStream = CreateWordDocumentStream(text);
                byte[] tableStream = CreateTableStream(text.Length);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateSimpleDocWithTable() {
                const string text = "A1\aB1\a\aA2\aB2\a\a\r";
                byte[] wordDocumentStream = CreateWordDocumentStream(text);
                byte[] tableStream = CreateTableStream(text.Length);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateUnicodeDocWithFormattedTableCell() {
                const string text = "A1\aB1\a\aA2\aB2\a\a\r";
                const int textOffset = 0x200;
                const int chpxFkpOffset = 0x400;
                byte[] wordDocumentStream = CreateUnicodeWordDocumentStreamWithFormattedTableCell(text, textOffset, chpxFkpOffset);
                byte[] tableStream = CreateUnicodeTableStreamWithCharacterBinTable(text.Length, textOffset, chpxFkpOffset / 512);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateUnicodeDocWithFormattedTableCellParagraph() {
                const string text = "A1\aB1\a\aA2\aB2\a\a\r";
                const int textOffset = 0x200;
                const int papxFkpOffset = 0x400;
                byte[] wordDocumentStream = CreateUnicodeWordDocumentStreamWithFormattedTableCellParagraph(text, textOffset, papxFkpOffset);
                byte[] tableStream = CreateUnicodeTableStreamWithParagraphBinTable(text.Length, textOffset, papxFkpOffset / 512);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateUnicodeDocWithExplicitTableMarkersAndTrailingEmptyCell() {
                const string text = "A1\a\a\a\r";
                const int textOffset = 0x200;
                const int papxFkpOffset = 0x400;
                byte[] wordDocumentStream = CreateUnicodeWordDocumentStreamWithExplicitTableMarkers(text, textOffset, papxFkpOffset);
                byte[] tableStream = CreateUnicodeTableStreamWithParagraphBinTable(text.Length, textOffset, papxFkpOffset / 512);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateUnicodeDocWithTableCellWidths() {
                const string text = "A1\aB1\a\a\r";
                const int textOffset = 0x200;
                const int papxFkpOffset = 0x400;
                byte[] wordDocumentStream = CreateUnicodeWordDocumentStreamWithExplicitTableMarkers(text, textOffset, papxFkpOffset, new[] { 1440, 2880 });
                byte[] tableStream = CreateUnicodeTableStreamWithParagraphBinTable(text.Length, textOffset, papxFkpOffset / 512);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateSimpleDocWithDocumentProperties(DateTime created, DateTime modified, params string[] paragraphs) {
                string text = string.Join("\r", paragraphs) + "\r";
                byte[] wordDocumentStream = CreateWordDocumentStream(text);
                byte[] tableStream = CreateTableStream(text.Length);
                byte[] summaryInformation = CreateSummaryInformationPropertySet(created, modified);
                byte[] documentSummaryInformation = CreateDocumentSummaryInformationPropertySet();

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                    WriteStream(root, "\u0005SummaryInformation", summaryInformation);
                    WriteStream(root, "\u0005DocumentSummaryInformation", documentSummaryInformation);
                }

                return package.ToArray();
            }

            internal static byte[] CreateSimpleDocWithUnsupportedFeatureStorage(params string[] paragraphs) {
                string text = string.Join("\r", paragraphs) + "\r";
                byte[] wordDocumentStream = CreateWordDocumentStream(text);
                byte[] tableStream = CreateTableStream(text.Length);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                    root.CreateStorage("_VBA_PROJECT_CUR");
                    root.CreateStorage("ObjectPool");
                }

                return package.ToArray();
            }

            internal static byte[] CreateSimpleDocWithActiveXAndEmbeddedPackageStorage(params string[] paragraphs) {
                string text = string.Join("\r", paragraphs) + "\r";
                byte[] wordDocumentStream = CreateWordDocumentStream(text);
                byte[] tableStream = CreateTableStream(text.Length);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                    root.CreateStorage("ActiveX");
                    var objectPool = root.CreateStorage("ObjectPool");
                    var packageStorage = objectPool.CreateStorage("OLEPackage");
                    using CfbStream nativePackage = packageStorage.CreateStream("\u0001Ole10Native");
                    nativePackage.Write(new byte[] { 1, 2, 3, 4 }, 0, 4);
                }

                return package.ToArray();
            }

            internal static byte[] CreateSimpleDocWithDataStream(params string[] paragraphs) {
                string text = string.Join("\r", paragraphs) + "\r";
                byte[] wordDocumentStream = CreateWordDocumentStream(text);
                byte[] tableStream = CreateTableStream(text.Length);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                    WriteStream(root, "Data", new byte[] { 1, 2, 3, 4 });
                }

                return package.ToArray();
            }

            internal static byte[] CreateSimpleDocWithFibFlags(ushort fibFlags, params string[] paragraphs) {
                string text = string.Join("\r", paragraphs) + "\r";
                byte[] wordDocumentStream = CreateWordDocumentStream(text, fibFlags: fibFlags);
                byte[] tableStream = CreateTableStream(text.Length);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateSimpleDocWithRevisionTrackingDop(uint dopSecondFlags, params string[] paragraphs) {
                string text = string.Join("\r", paragraphs) + "\r";
                const int dopOffset = 21;
                const int dopLength = 8;
                byte[] wordDocumentStream = CreateWordDocumentStream(text, fcDop: dopOffset, lcbDop: dopLength);
                byte[] tableStream = CreateTableStream(text.Length);
                Array.Resize(ref tableStream, dopOffset + dopLength);
                WriteUInt32(tableStream, dopOffset + 4, dopSecondFlags);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateSimpleDocWithFibVersion(ushort nFib, params string[] paragraphs) {
                string text = string.Join("\r", paragraphs) + "\r";
                byte[] wordDocumentStream = CreateWordDocumentStream(text, nFib: nFib);
                byte[] tableStream = CreateTableStream(text.Length);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateSimpleDocWithUnsupportedStoryCounts(params string[] paragraphs) {
                string text = string.Join("\r", paragraphs) + "\r";
                byte[] wordDocumentStream = CreateWordDocumentStream(
                    text,
                    ccpFtn: 3,
                    ccpHdd: 5,
                    ccpAtn: 7,
                    ccpEdn: 11,
                    ccpTxbx: 13,
                    ccpHdrTxbx: 17);
                byte[] tableStream = CreateTableStream(text.Length);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateSimpleDocWithUnsupportedHeaderFooterStoryText(string bodyParagraph, string headerFooterStory) {
                string bodyText = bodyParagraph + "\r";
                string storyText = headerFooterStory + "\r";
                string documentText = bodyText + storyText;
                byte[] wordDocumentStream = CreateWordDocumentStream(
                    documentText,
                    ccpTextOverride: bodyText.Length,
                    ccpHdd: storyText.Length);
                byte[] tableStream = CreateTableStream(documentText.Length);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateSimpleDocWithMultipleSectionDescriptors(params string[] paragraphs) {
                string text = string.Join("\r", paragraphs) + "\r";
                byte[] tableStream = CreateTableStream(text.Length);
                int fcPlcfSed = tableStream.Length;
                byte[] sectionDescriptorPlc = CreateTwoSectionDescriptorPlc(text.Length);
                Array.Resize(ref tableStream, tableStream.Length + sectionDescriptorPlc.Length);
                Buffer.BlockCopy(sectionDescriptorPlc, 0, tableStream, fcPlcfSed, sectionDescriptorPlc.Length);
                byte[] wordDocumentStream = CreateWordDocumentStream(
                    text,
                    fcPlcfSed: fcPlcfSed,
                    lcbPlcfSed: sectionDescriptorPlc.Length);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateSimpleDocWithTwoSectionPageSetup() {
                const string firstParagraph = "Portrait section";
                const string secondParagraph = "Landscape section";
                string text = firstParagraph + "\r" + secondParagraph + "\r";
                int firstSectionEnd = firstParagraph.Length + 1;
                const int firstSepxOffset = 0x300;
                const int secondSepxOffset = 0x340;

                byte[] tableStream = CreateTableStream(text.Length);
                int fcPlcfSed = tableStream.Length;
                byte[] sectionDescriptorPlc = CreateTwoSectionDescriptorPlc(
                    firstSectionEnd,
                    text.Length,
                    firstSepxOffset,
                    secondSepxOffset);
                Array.Resize(ref tableStream, tableStream.Length + sectionDescriptorPlc.Length);
                Buffer.BlockCopy(sectionDescriptorPlc, 0, tableStream, fcPlcfSed, sectionDescriptorPlc.Length);

                byte[] wordDocumentStream = CreateWordDocumentStream(
                    text,
                    fcPlcfSed: fcPlcfSed,
                    lcbPlcfSed: sectionDescriptorPlc.Length);
                WriteBytesAt(ref wordDocumentStream, firstSepxOffset, CreateSectionSepx());
                WriteBytesAt(
                    ref wordDocumentStream,
                    secondSepxOffset,
                    CreateSectionSepx(
                        orientation: 2,
                        pageWidth: 15840,
                        pageHeight: 12240,
                        marginLeft: 720,
                        marginRight: 720,
                        marginTop: 720,
                        marginBottom: 720));

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateSimpleDocWithSectionBreakKind(int sectionBreakOperand, string secondParagraph) {
                const string firstParagraph = "Before continuous section";
                string text = firstParagraph + "\r" + secondParagraph + "\r";
                int firstSectionEnd = firstParagraph.Length + 1;
                const int secondSepxOffset = 0x300;

                byte[] tableStream = CreateTableStream(text.Length);
                int fcPlcfSed = tableStream.Length;
                byte[] sectionDescriptorPlc = CreateTwoSectionDescriptorPlc(
                    firstSectionEnd,
                    text.Length,
                    0,
                    secondSepxOffset);
                Array.Resize(ref tableStream, tableStream.Length + sectionDescriptorPlc.Length);
                Buffer.BlockCopy(sectionDescriptorPlc, 0, tableStream, fcPlcfSed, sectionDescriptorPlc.Length);

                byte[] wordDocumentStream = CreateWordDocumentStream(
                    text,
                    fcPlcfSed: fcPlcfSed,
                    lcbPlcfSed: sectionDescriptorPlc.Length);
                WriteBytesAt(ref wordDocumentStream, secondSepxOffset, CreateSectionSepx(sectionBreakType: checked((byte)sectionBreakOperand)));

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateUnicodeDocWithDirectCharacterFormatting() {
                const string text = "plain bold italic\r";
                const int textOffset = 0x200;
                const int chpxFkpOffset = 0x400;
                byte[] wordDocumentStream = CreateUnicodeWordDocumentStreamWithDirectCharacterFormatting(text, textOffset, chpxFkpOffset);
                byte[] tableStream = CreateUnicodeTableStreamWithCharacterBinTable(text.Length, textOffset, chpxFkpOffset / 512);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateUnicodeDocWithExtendedDirectCharacterFormatting() {
                const string text = "plain under sized red strike double outline shadow emboss imprint hidden caps small super sub mark direct\r";
                const int textOffset = 0x200;
                const int chpxFkpOffset = 0x400;
                byte[] wordDocumentStream = CreateUnicodeWordDocumentStreamWithExtendedCharacterFormatting(text, textOffset, chpxFkpOffset);
                byte[] tableStream = CreateUnicodeTableStreamWithCharacterBinTable(text.Length, textOffset, chpxFkpOffset / 512);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateUnicodeDocWithFontFamilyFormatting() {
                const string text = "plain font\r";
                const int textOffset = 0x200;
                const int chpxFkpOffset = 0x400;
                byte[] fontTable = CreateFontTable("Courier New");
                byte[] wordDocumentStream = CreateUnicodeWordDocumentStreamWithFontFamilyFormatting(text, textOffset, chpxFkpOffset, fontTable.Length);
                byte[] tableStream = CreateUnicodeTableStreamWithCharacterBinTableAndFontTable(text.Length, textOffset, chpxFkpOffset / 512, fontTable);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateUnicodeDocWithParagraphAlignment() {
                const string text = "left\rcenter\rright\r";
                const int textOffset = 0x200;
                const int papxFkpOffset = 0x400;
                byte[] wordDocumentStream = CreateUnicodeWordDocumentStreamWithParagraphAlignment(text, textOffset, papxFkpOffset);
                byte[] tableStream = CreateUnicodeTableStreamWithParagraphBinTable(text.Length, textOffset, papxFkpOffset / 512);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateUnicodeDocWithParagraphSpacingAndIndentation() {
                const string text = "plain\rformatted\rhanging\r";
                const int textOffset = 0x200;
                const int papxFkpOffset = 0x400;
                byte[] wordDocumentStream = CreateUnicodeWordDocumentStreamWithParagraphSpacingAndIndentation(text, textOffset, papxFkpOffset);
                byte[] tableStream = CreateUnicodeTableStreamWithParagraphBinTable(text.Length, textOffset, papxFkpOffset / 512);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateUnicodeDocWithParagraphPaginationFlags() {
                const string text = "plain\rpagination\r";
                const int textOffset = 0x200;
                const int papxFkpOffset = 0x400;
                byte[] wordDocumentStream = CreateUnicodeWordDocumentStreamWithParagraphPaginationFlags(text, textOffset, papxFkpOffset);
                byte[] tableStream = CreateUnicodeTableStreamWithParagraphBinTable(text.Length, textOffset, papxFkpOffset / 512);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateUnicodeDocWithParagraphTabStops() {
                const string text = "plain\rtabs\rclear\r";
                const int textOffset = 0x200;
                const int papxFkpOffset = 0x400;
                byte[] wordDocumentStream = CreateUnicodeWordDocumentStreamWithParagraphTabStops(text, textOffset, papxFkpOffset);
                byte[] tableStream = CreateUnicodeTableStreamWithParagraphBinTable(text.Length, textOffset, papxFkpOffset / 512);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateUnicodeDocWithStyleLevelParagraphTabStops() {
                const string text = "style tabs\r";
                const int textOffset = 0x200;
                const int papxFkpOffset = 0x400;
                byte[] styleSheet = CreateStyleSheet(
                    CreateParagraphStyleRecord(0, 0x0FFF, "Normal"),
                    CreateParagraphStyleRecord(
                        0x0FFF,
                        0,
                        "Tab Style",
                        CreateStyleParagraphFormatting(CreateParagraphTabStopsSprm(
                            new[] { 3600 },
                            (1800, 1, 1)))));
                byte[] wordDocumentStream = CreateUnicodeWordDocumentStreamWithStyleIndex(text, textOffset, papxFkpOffset, styleSheet.Length, 1);
                byte[] tableStream = CreateUnicodeTableStreamWithParagraphBinTableAndStyleSheet(text.Length, textOffset, papxFkpOffset / 512, styleSheet);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateUnicodeDocWithStyleLevelCapsDoubleStrikeAndVerticalPosition() {
                const string text = "caps style\rsmall style\rsuper style\rsub style\r";
                const int textOffset = 0x200;
                const int papxFkpOffset = 0x400;
                byte[] styleSheet = CreateStyleSheet(
                    CreateParagraphStyleRecord(0, 0x0FFF, "Normal"),
                    CreateParagraphStyleRecord(
                        0x0FFF,
                        0,
                        "Caps Double",
                        Array.Empty<byte>(),
                        CreateStyleCharacterFormatting(
                            CreateCharacterSprm(0x083B, 1),
                            CreateCharacterSprm(0x2A53, 1),
                            CreateCharacterSprm(0x0838, 1),
                            CreateCharacterSprm(0x0839, 1),
                            CreateCharacterSprm(0x0858, 1),
                            CreateCharacterSprm(0x0854, 1),
                            CreateCharacterSprm(0x083C, 1))),
                    CreateParagraphStyleRecord(
                        0x0FFF,
                        0,
                        "Small Caps",
                        Array.Empty<byte>(),
                        CreateStyleCharacterFormatting(CreateCharacterSprm(0x083A, 1))),
                    CreateParagraphStyleRecord(
                        0x0FFF,
                        0,
                        "Super",
                        Array.Empty<byte>(),
                        CreateStyleCharacterFormatting(CreateCharacterSprm(0x2A48, 1))),
                    CreateParagraphStyleRecord(
                        0x0FFF,
                        0,
                        "Sub",
                        Array.Empty<byte>(),
                        CreateStyleCharacterFormatting(CreateCharacterSprm(0x2A48, 2))));
                byte[] wordDocumentStream = CreateUnicodeWordDocumentStreamWithStyleLevelCapsDoubleStrikeAndVerticalPosition(text, textOffset, papxFkpOffset, styleSheet.Length);
                byte[] tableStream = CreateUnicodeTableStreamWithParagraphBinTableAndStyleSheet(text.Length, textOffset, papxFkpOffset / 512, styleSheet);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateUnicodeDocWithBuiltInStyleLevelFormatting() {
                const string text = "built in heading\r";
                const int textOffset = 0x200;
                const int papxFkpOffset = 0x400;
                byte[] styleSheet = CreateStyleSheet(
                    CreateParagraphStyleRecord(0, 0x0FFF, "Normal"),
                    CreateParagraphStyleRecord(
                        1,
                        0,
                        "heading 1",
                        CreateStyleParagraphFormatting(
                            CreateParagraphSprm(0x2461, 1),
                            CreateParagraphSprm(0xA413, 0xF0, 0x00),
                            CreateParagraphSprm(0xA414, 0x78, 0x00),
                            CreateParagraphTabStopsSprm(
                                Array.Empty<int>(),
                                (1440, 0, 0),
                                (4320, 2, 3))),
                        CreateStyleCharacterFormatting(
                            CreateCharacterSprm(0x0835, 1),
                            CreateCharacterSprm(0x0838, 1),
                            CreateCharacterSprm(0x0839, 1),
                            CreateCharacterSprm(0x0858, 1),
                            CreateCharacterSprm(0x0854, 1),
                            CreateCharacterSprm(0x083C, 1),
                            CreateCharacterSprm(0x2A3E, 1),
                            CreateCharacterSprm(0x2A0C, 7),
                            CreateCharacterSprm(0x6870, 0x33, 0x66, 0x99, 0x00),
                            CreateCharacterSprm(0x4A43, 0x20, 0x00))));
                byte[] wordDocumentStream = CreateUnicodeWordDocumentStreamWithBuiltInStyleLevelFormatting(text, textOffset, papxFkpOffset, styleSheet.Length);
                byte[] tableStream = CreateUnicodeTableStreamWithParagraphBinTableAndStyleSheet(text.Length, textOffset, papxFkpOffset / 512, styleSheet);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateUnicodeDocWithInheritedBuiltInStyleFormatting() {
                const string text = "inherited heading\r";
                const int textOffset = 0x200;
                const int papxFkpOffset = 0x400;
                byte[] styleSheet = CreateStyleSheet(
                    CreateParagraphStyleRecord(0, 0x0FFF, "Normal"),
                    CreateParagraphStyleRecord(
                        1,
                        0,
                        "heading 1",
                        CreateStyleParagraphFormatting(CreateParagraphSprm(0x2461, 1)),
                        CreateStyleCharacterFormatting(
                            CreateCharacterSprm(0x0835, 1),
                            CreateCharacterSprm(0x6870, 0x33, 0x66, 0x99, 0x00))),
                    CreateParagraphStyleRecord(
                        0x0FFF,
                        1,
                        "Inherited Heading",
                        Array.Empty<byte>(),
                        CreateStyleCharacterFormatting(CreateCharacterSprm(0x0836, 1))));
                byte[] wordDocumentStream = CreateUnicodeWordDocumentStreamWithStyleIndex(text, textOffset, papxFkpOffset, styleSheet.Length, 2);
                byte[] tableStream = CreateUnicodeTableStreamWithParagraphBinTableAndStyleSheet(text.Length, textOffset, papxFkpOffset / 512, styleSheet);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateCompoundWithoutWordDocumentStream() {
                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "NotWordDocument", new byte[] { 1, 2, 3 });
                }

                return package.ToArray();
            }

            private static byte[] CreateWordDocumentStream(
                string text,
                int ccpFtn = 0,
                int ccpHdd = 0,
                int ccpAtn = 0,
                int ccpEdn = 0,
                int ccpTxbx = 0,
                int ccpHdrTxbx = 0,
                ushort nFib = 0x00D9,
                ushort fibFlags = 0x0200,
                int fcPlcfSed = 0,
                int lcbPlcfSed = 0,
                int fcDop = 0,
                int lcbDop = 0,
                int? ccpTextOverride = null) {
                const int fibLength = 0x1AA;
                const int textOffset = 0x200;
                byte[] textBytes = EncodeWindows1252(text);
                var stream = new byte[textOffset + textBytes.Length];
                WriteUInt16(stream, 0x00, 0xA5EC);
                WriteUInt16(stream, 0x02, nFib);
                WriteUInt16(stream, 0x0A, fibFlags);
                WriteInt32(stream, 0x4C, ccpTextOverride ?? text.Length);
                WriteInt32(stream, 0x50, ccpFtn);
                WriteInt32(stream, 0x54, ccpHdd);
                WriteInt32(stream, 0x5C, ccpAtn);
                WriteInt32(stream, 0x60, ccpEdn);
                WriteInt32(stream, 0x64, ccpTxbx);
                WriteInt32(stream, 0x68, ccpHdrTxbx);
                WriteInt32(stream, 0xCA, fcPlcfSed);
                WriteInt32(stream, 0xCE, lcbPlcfSed);
                WriteInt32(stream, 0x192, fcDop);
                WriteInt32(stream, 0x196, lcbDop);
                WriteInt32(stream, 0x1A2, 0);
                WriteInt32(stream, 0x1A6, 21);
                Buffer.BlockCopy(textBytes, 0, stream, textOffset, textBytes.Length);
                if (stream.Length < fibLength) {
                    Array.Resize(ref stream, fibLength);
                }

                return stream;
            }

            private static byte[] CreateTwoSectionDescriptorPlc(int characterCount) {
                return CreateTwoSectionDescriptorPlc(Math.Max(0, characterCount - 1), characterCount, 0, 0);
            }

            private static byte[] CreateTwoSectionDescriptorPlc(int firstSectionEnd, int characterCount, int firstSepxOffset, int secondSepxOffset) {
                var plc = new byte[36];
                WriteInt32(plc, 0, 0);
                WriteInt32(plc, 4, firstSectionEnd);
                WriteInt32(plc, 8, characterCount);
                WriteInt32(plc, 14, firstSepxOffset);
                WriteInt32(plc, 26, secondSepxOffset);
                return plc;
            }

            private static byte[] CreateSectionSepx(
                byte? sectionBreakType = null,
                byte? orientation = null,
                int? pageWidth = null,
                int? pageHeight = null,
                int? marginLeft = null,
                int? marginRight = null,
                int? marginTop = null,
                int? marginBottom = null) {
                var grpprl = new List<byte>();
                if (sectionBreakType != null) {
                    AddSingleByteSprm(grpprl, 0x3009, sectionBreakType.Value);
                }

                if (orientation != null) {
                    AddSingleByteSprm(grpprl, 0x301D, orientation.Value);
                }

                AddUInt16SprmIfPresent(grpprl, 0xB01F, pageWidth);
                AddUInt16SprmIfPresent(grpprl, 0xB020, pageHeight);
                AddUInt16SprmIfPresent(grpprl, 0xB021, marginLeft);
                AddUInt16SprmIfPresent(grpprl, 0xB022, marginRight);
                AddUInt16SprmIfPresent(grpprl, 0x9023, marginTop);
                AddUInt16SprmIfPresent(grpprl, 0x9024, marginBottom);

                var sepx = new byte[2 + grpprl.Count];
                WriteUInt16(sepx, 0, (ushort)grpprl.Count);
                grpprl.CopyTo(sepx, 2);
                return sepx;
            }

            private static void AddSingleByteSprm(List<byte> grpprl, ushort sprm, byte operand) {
                grpprl.Add((byte)(sprm & 0xFF));
                grpprl.Add((byte)(sprm >> 8));
                grpprl.Add(operand);
            }

            private static void AddUInt16SprmIfPresent(List<byte> grpprl, ushort sprm, int? operand) {
                if (operand == null) {
                    return;
                }

                grpprl.Add((byte)(sprm & 0xFF));
                grpprl.Add((byte)(sprm >> 8));
                grpprl.Add((byte)(operand.Value & 0xFF));
                grpprl.Add((byte)(operand.Value >> 8));
            }

            private static void WriteBytesAt(ref byte[] bytes, int offset, byte[] value) {
                if (bytes.Length < offset + value.Length) {
                    Array.Resize(ref bytes, offset + value.Length);
                }

                Buffer.BlockCopy(value, 0, bytes, offset, value.Length);
            }

            private static byte[] CreateUnicodeWordDocumentStreamWithFormattedTableCell(string text, int textOffset, int chpxFkpOffset) {
                const int fibLength = 0x1AA;
                byte[] textBytes = Encoding.Unicode.GetBytes(text);
                var stream = new byte[Math.Max(chpxFkpOffset + 512, textOffset + textBytes.Length)];
                WriteUInt16(stream, 0x00, 0xA5EC);
                WriteUInt16(stream, 0x02, 0x00D9);
                WriteUInt16(stream, 0x0A, 0x0200);
                WriteInt32(stream, 0x4C, text.Length);
                WriteInt32(stream, 0xFA, 21);
                WriteInt32(stream, 0xFE, 12);
                WriteInt32(stream, 0x1A2, 0);
                WriteInt32(stream, 0x1A6, 21);
                Buffer.BlockCopy(textBytes, 0, stream, textOffset, textBytes.Length);

                int firstCellEnd = textOffset + ("A1".Length * 2);
                int end = textOffset + (text.Length * 2);
                WriteChpxFkp(
                    stream,
                    chpxFkpOffset,
                    new[] { textOffset, firstCellEnd, end },
                    new Dictionary<int, byte[]> {
                        [0] = CreateSingleSprmChpx(0x0835, 1)
                    });

                if (stream.Length < fibLength) {
                    Array.Resize(ref stream, fibLength);
                }

                return stream;
            }

            private static byte[] CreateUnicodeWordDocumentStreamWithFormattedTableCellParagraph(string text, int textOffset, int papxFkpOffset) {
                const int fibLength = 0x1AA;
                byte[] textBytes = Encoding.Unicode.GetBytes(text);
                var stream = new byte[Math.Max(papxFkpOffset + 512, textOffset + textBytes.Length)];
                WriteUInt16(stream, 0x00, 0xA5EC);
                WriteUInt16(stream, 0x02, 0x00D9);
                WriteUInt16(stream, 0x0A, 0x0200);
                WriteInt32(stream, 0x4C, text.Length);
                WriteInt32(stream, 0x102, 21);
                WriteInt32(stream, 0x106, 12);
                WriteInt32(stream, 0x1A2, 0);
                WriteInt32(stream, 0x1A6, 21);
                Buffer.BlockCopy(textBytes, 0, stream, textOffset, textBytes.Length);

                int secondCellStart = textOffset + ("A1\a".Length * 2);
                int end = textOffset + (text.Length * 2);
                WritePapxFkp(
                    stream,
                    papxFkpOffset,
                    new[] { textOffset, secondCellStart, end },
                    new Dictionary<int, byte[]> {
                        [0] = CreateParagraphPropertiesPapx(
                            CreateParagraphSprm(0x2461, 1),
                            CreateParagraphSprm(0xA414, 0x78, 0x00),
                            CreateParagraphSprm(0x840F, 0x68, 0x01))
                    });

                if (stream.Length < fibLength) {
                    Array.Resize(ref stream, fibLength);
                }

                return stream;
            }

            private static byte[] CreateUnicodeWordDocumentStreamWithExplicitTableMarkers(string text, int textOffset, int papxFkpOffset, IReadOnlyList<int>? tableCellWidthsTwips = null) {
                const int fibLength = 0x1AA;
                byte[] textBytes = Encoding.Unicode.GetBytes(text);
                var stream = new byte[Math.Max(papxFkpOffset + 512, textOffset + textBytes.Length)];
                WriteUInt16(stream, 0x00, 0xA5EC);
                WriteUInt16(stream, 0x02, 0x00D9);
                WriteUInt16(stream, 0x0A, 0x0200);
                WriteInt32(stream, 0x4C, text.Length);
                WriteInt32(stream, 0x102, 21);
                WriteInt32(stream, 0x106, 12);
                WriteInt32(stream, 0x1A2, 0);
                WriteInt32(stream, 0x1A6, 21);
                Buffer.BlockCopy(textBytes, 0, stream, textOffset, textBytes.Length);

                int[] markerEnds = GetFirstMarkerEnds(text, textOffset, 3);
                int firstCellMarkerEnd = markerEnds[0];
                int secondCellMarkerEnd = markerEnds[1];
                int rowMarkerEnd = markerEnds[2];
                int end = textOffset + (text.Length * 2);
                byte[] tableCellPapx = CreateParagraphPropertiesPapx(CreateParagraphSprm(0x2416, 1));
                byte[] tableRowPapx = tableCellWidthsTwips == null
                    ? CreateParagraphPropertiesPapx(
                        CreateParagraphSprm(0x2416, 1),
                        CreateParagraphSprm(0x2417, 1))
                    : CreateParagraphPropertiesPapx(
                        CreateParagraphSprm(0x2416, 1),
                        CreateParagraphSprm(0x2417, 1),
                        CreateTableDefinitionSprm(tableCellWidthsTwips));
                WritePapxFkp(
                    stream,
                    papxFkpOffset,
                    new[] { textOffset, firstCellMarkerEnd, secondCellMarkerEnd, rowMarkerEnd, end },
                    new Dictionary<int, byte[]> {
                        [0] = tableCellPapx,
                        [1] = tableCellPapx,
                        [2] = tableRowPapx
                    });

                if (stream.Length < fibLength) {
                    Array.Resize(ref stream, fibLength);
                }

                return stream;
            }

            private static byte[] CreateUnicodeWordDocumentStreamWithDirectCharacterFormatting(string text, int textOffset, int chpxFkpOffset) {
                const int fibLength = 0x1AA;
                byte[] textBytes = System.Text.Encoding.Unicode.GetBytes(text);
                var stream = new byte[Math.Max(chpxFkpOffset + 512, textOffset + textBytes.Length)];
                WriteUInt16(stream, 0x00, 0xA5EC);
                WriteUInt16(stream, 0x02, 0x00D9);
                WriteUInt16(stream, 0x0A, 0x0200);
                WriteInt32(stream, 0x4C, text.Length);
                WriteInt32(stream, 0xFA, 21);
                WriteInt32(stream, 0xFE, 12);
                WriteInt32(stream, 0x1A2, 0);
                WriteInt32(stream, 0x1A6, 21);
                Buffer.BlockCopy(textBytes, 0, stream, textOffset, textBytes.Length);

                int boldStart = textOffset + ("plain ".Length * 2);
                int italicStart = boldStart + ("bold ".Length * 2);
                int paragraphMarkStart = italicStart + ("italic".Length * 2);
                int end = paragraphMarkStart + 2;
                WriteChpxFkp(
                    stream,
                    chpxFkpOffset,
                    new[] { textOffset, boldStart, italicStart, paragraphMarkStart, end },
                    boldRunIndex: 1,
                    italicRunIndex: 2);

                if (stream.Length < fibLength) {
                    Array.Resize(ref stream, fibLength);
                }

                return stream;
            }

            private static byte[] CreateUnicodeWordDocumentStreamWithExtendedCharacterFormatting(string text, int textOffset, int chpxFkpOffset) {
                const int fibLength = 0x1AA;
                byte[] textBytes = System.Text.Encoding.Unicode.GetBytes(text);
                var stream = new byte[Math.Max(chpxFkpOffset + 512, textOffset + textBytes.Length)];
                WriteUInt16(stream, 0x00, 0xA5EC);
                WriteUInt16(stream, 0x02, 0x00D9);
                WriteUInt16(stream, 0x0A, 0x0200);
                WriteInt32(stream, 0x4C, text.Length);
                WriteInt32(stream, 0xFA, 21);
                WriteInt32(stream, 0xFE, 12);
                WriteInt32(stream, 0x1A2, 0);
                WriteInt32(stream, 0x1A6, 21);
                Buffer.BlockCopy(textBytes, 0, stream, textOffset, textBytes.Length);

                int underStart = textOffset + ("plain ".Length * 2);
                int sizedStart = underStart + ("under ".Length * 2);
                int redStart = sizedStart + ("sized ".Length * 2);
                int strikeStart = redStart + ("red ".Length * 2);
                int doubleStrikeStart = strikeStart + ("strike ".Length * 2);
                int outlineStart = doubleStrikeStart + ("double ".Length * 2);
                int shadowStart = outlineStart + ("outline ".Length * 2);
                int embossStart = shadowStart + ("shadow ".Length * 2);
                int imprintStart = embossStart + ("emboss ".Length * 2);
                int hiddenStart = imprintStart + ("imprint ".Length * 2);
                int capsStart = hiddenStart + ("hidden ".Length * 2);
                int smallCapsStart = capsStart + ("caps ".Length * 2);
                int superStart = smallCapsStart + ("small ".Length * 2);
                int subStart = superStart + ("super ".Length * 2);
                int markStart = subStart + ("sub ".Length * 2);
                int directStart = markStart + ("mark ".Length * 2);
                int paragraphMarkStart = directStart + ("direct".Length * 2);
                int end = paragraphMarkStart + 2;
                WriteChpxFkp(
                    stream,
                    chpxFkpOffset,
                    new[] { textOffset, underStart, sizedStart, redStart, strikeStart, doubleStrikeStart, outlineStart, shadowStart, embossStart, imprintStart, hiddenStart, capsStart, smallCapsStart, superStart, subStart, markStart, directStart, paragraphMarkStart, end },
                    new Dictionary<int, byte[]> {
                        [1] = CreateSingleSprmChpx(0x2A3E, 1),
                        [2] = CreateSingleSprmChpx(0x4A43, 28, 0),
                        [3] = CreateSingleSprmChpx(0x2A42, 6),
                        [4] = CreateSingleSprmChpx(0x0837, 1),
                        [5] = CreateSingleSprmChpx(0x2A53, 1),
                        [6] = CreateSingleSprmChpx(0x0838, 1),
                        [7] = CreateSingleSprmChpx(0x0839, 1),
                        [8] = CreateSingleSprmChpx(0x0858, 1),
                        [9] = CreateSingleSprmChpx(0x0854, 1),
                        [10] = CreateSingleSprmChpx(0x083C, 1),
                        [11] = CreateSingleSprmChpx(0x083B, 1),
                        [12] = CreateSingleSprmChpx(0x083A, 1),
                        [13] = CreateSingleSprmChpx(0x2A48, 1),
                        [14] = CreateSingleSprmChpx(0x2A48, 2),
                        [15] = CreateSingleSprmChpx(0x2A0C, 7),
                        [16] = CreateSingleSprmChpx(0x6870, 0x33, 0x66, 0x99, 0)
                    });

                if (stream.Length < fibLength) {
                    Array.Resize(ref stream, fibLength);
                }

                return stream;
            }

            private static byte[] CreateUnicodeWordDocumentStreamWithFontFamilyFormatting(string text, int textOffset, int chpxFkpOffset, int fontTableLength) {
                const int fibLength = 0x1AA;
                byte[] textBytes = System.Text.Encoding.Unicode.GetBytes(text);
                var stream = new byte[Math.Max(chpxFkpOffset + 512, textOffset + textBytes.Length)];
                WriteUInt16(stream, 0x00, 0xA5EC);
                WriteUInt16(stream, 0x02, 0x00D9);
                WriteUInt16(stream, 0x0A, 0x0200);
                WriteInt32(stream, 0x4C, text.Length);
                WriteInt32(stream, 0xFA, 21);
                WriteInt32(stream, 0xFE, 12);
                WriteInt32(stream, 0x112, 33);
                WriteInt32(stream, 0x116, fontTableLength);
                WriteInt32(stream, 0x1A2, 0);
                WriteInt32(stream, 0x1A6, 21);
                Buffer.BlockCopy(textBytes, 0, stream, textOffset, textBytes.Length);

                int fontStart = textOffset + ("plain ".Length * 2);
                int paragraphMarkStart = fontStart + ("font".Length * 2);
                int end = paragraphMarkStart + 2;
                WriteChpxFkp(
                    stream,
                    chpxFkpOffset,
                    new[] { textOffset, fontStart, paragraphMarkStart, end },
                    new Dictionary<int, byte[]> {
                        [1] = CreateSingleSprmChpx(0x4A4F, 0, 0)
                    });

                if (stream.Length < fibLength) {
                    Array.Resize(ref stream, fibLength);
                }

                return stream;
            }

            private static byte[] CreateUnicodeWordDocumentStreamWithParagraphAlignment(string text, int textOffset, int papxFkpOffset) {
                const int fibLength = 0x1AA;
                byte[] textBytes = System.Text.Encoding.Unicode.GetBytes(text);
                var stream = new byte[Math.Max(papxFkpOffset + 512, textOffset + textBytes.Length)];
                WriteUInt16(stream, 0x00, 0xA5EC);
                WriteUInt16(stream, 0x02, 0x00D9);
                WriteUInt16(stream, 0x0A, 0x0200);
                WriteInt32(stream, 0x4C, text.Length);
                WriteInt32(stream, 0x102, 21);
                WriteInt32(stream, 0x106, 12);
                WriteInt32(stream, 0x1A2, 0);
                WriteInt32(stream, 0x1A6, 21);
                Buffer.BlockCopy(textBytes, 0, stream, textOffset, textBytes.Length);

                int centerStart = textOffset + ("left\r".Length * 2);
                int rightStart = centerStart + ("center\r".Length * 2);
                int end = rightStart + ("right\r".Length * 2);
                WritePapxFkp(
                    stream,
                    papxFkpOffset,
                    new[] { textOffset, centerStart, rightStart, end },
                    new Dictionary<int, byte[]> {
                        [1] = CreateParagraphAlignmentPapx(1),
                        [2] = CreateParagraphAlignmentPapx(2)
                    });

                if (stream.Length < fibLength) {
                    Array.Resize(ref stream, fibLength);
                }

                return stream;
            }

            private static byte[] CreateUnicodeWordDocumentStreamWithParagraphSpacingAndIndentation(string text, int textOffset, int papxFkpOffset) {
                const int fibLength = 0x1AA;
                byte[] textBytes = System.Text.Encoding.Unicode.GetBytes(text);
                var stream = new byte[Math.Max(papxFkpOffset + 512, textOffset + textBytes.Length)];
                WriteUInt16(stream, 0x00, 0xA5EC);
                WriteUInt16(stream, 0x02, 0x00D9);
                WriteUInt16(stream, 0x0A, 0x0200);
                WriteInt32(stream, 0x4C, text.Length);
                WriteInt32(stream, 0x102, 21);
                WriteInt32(stream, 0x106, 12);
                WriteInt32(stream, 0x1A2, 0);
                WriteInt32(stream, 0x1A6, 21);
                Buffer.BlockCopy(textBytes, 0, stream, textOffset, textBytes.Length);

                int formattedStart = textOffset + ("plain\r".Length * 2);
                int hangingStart = formattedStart + ("formatted\r".Length * 2);
                int end = hangingStart + ("hanging\r".Length * 2);
                WritePapxFkp(
                    stream,
                    papxFkpOffset,
                    new[] { textOffset, formattedStart, hangingStart, end },
                    new Dictionary<int, byte[]> {
                        [1] = CreateParagraphPropertiesPapx(
                            CreateParagraphSprm(0xA413, 0xF0, 0x00),
                            CreateParagraphSprm(0xA414, 0x78, 0x00),
                            CreateParagraphSprm(0x6412, 0x68, 0x01, 0x00, 0x00),
                            CreateParagraphSprm(0x840F, 0xD0, 0x02),
                            CreateParagraphSprm(0x840E, 0x68, 0x01),
                            CreateParagraphSprm(0x8411, 0xF0, 0x00)),
                        [2] = CreateParagraphPropertiesPapx(
                            CreateParagraphSprm(0x840F, 0xD0, 0x02),
                            CreateParagraphSprm(0x8411, 0x98, 0xFE))
                    });

                if (stream.Length < fibLength) {
                    Array.Resize(ref stream, fibLength);
                }

                return stream;
            }

            private static byte[] CreateUnicodeWordDocumentStreamWithParagraphPaginationFlags(string text, int textOffset, int papxFkpOffset) {
                const int fibLength = 0x1AA;
                byte[] textBytes = System.Text.Encoding.Unicode.GetBytes(text);
                var stream = new byte[Math.Max(papxFkpOffset + 512, textOffset + textBytes.Length)];
                WriteUInt16(stream, 0x00, 0xA5EC);
                WriteUInt16(stream, 0x02, 0x00D9);
                WriteUInt16(stream, 0x0A, 0x0200);
                WriteInt32(stream, 0x4C, text.Length);
                WriteInt32(stream, 0x102, 21);
                WriteInt32(stream, 0x106, 12);
                WriteInt32(stream, 0x1A2, 0);
                WriteInt32(stream, 0x1A6, 21);
                Buffer.BlockCopy(textBytes, 0, stream, textOffset, textBytes.Length);

                int formattedStart = textOffset + ("plain\r".Length * 2);
                int end = formattedStart + ("pagination\r".Length * 2);
                WritePapxFkp(
                    stream,
                    papxFkpOffset,
                    new[] { textOffset, formattedStart, end },
                    new Dictionary<int, byte[]> {
                        [1] = CreateParagraphPropertiesPapx(
                            CreateParagraphSprm(0x2405, 1),
                            CreateParagraphSprm(0x2406, 1),
                            CreateParagraphSprm(0x2407, 1),
                            CreateParagraphSprm(0x2431, 1))
                    });

                if (stream.Length < fibLength) {
                    Array.Resize(ref stream, fibLength);
                }

                return stream;
            }

            private static byte[] CreateUnicodeWordDocumentStreamWithParagraphTabStops(string text, int textOffset, int papxFkpOffset) {
                const int fibLength = 0x1AA;
                byte[] textBytes = System.Text.Encoding.Unicode.GetBytes(text);
                var stream = new byte[Math.Max(papxFkpOffset + 512, textOffset + textBytes.Length)];
                WriteUInt16(stream, 0x00, 0xA5EC);
                WriteUInt16(stream, 0x02, 0x00D9);
                WriteUInt16(stream, 0x0A, 0x0200);
                WriteInt32(stream, 0x4C, text.Length);
                WriteInt32(stream, 0x102, 21);
                WriteInt32(stream, 0x106, 12);
                WriteInt32(stream, 0x1A2, 0);
                WriteInt32(stream, 0x1A6, 21);
                Buffer.BlockCopy(textBytes, 0, stream, textOffset, textBytes.Length);

                int tabsStart = textOffset + ("plain\r".Length * 2);
                int clearStart = tabsStart + ("tabs\r".Length * 2);
                int end = clearStart + ("clear\r".Length * 2);
                WritePapxFkp(
                    stream,
                    papxFkpOffset,
                    new[] { textOffset, tabsStart, clearStart, end },
                    new Dictionary<int, byte[]> {
                        [1] = CreateParagraphPropertiesPapx(CreateParagraphTabStopsSprm(
                            Array.Empty<int>(),
                            (1440, 0, 0),
                            (2880, 3, 1),
                            (4320, 2, 3))),
                        [2] = CreateParagraphPropertiesPapx(CreateParagraphTabStopsSprm(
                            new[] { 1440 },
                            (2160, 4, 0)))
                    });

                if (stream.Length < fibLength) {
                    Array.Resize(ref stream, fibLength);
                }

                return stream;
            }

            private static byte[] CreateUnicodeWordDocumentStreamWithStyleLevelCapsDoubleStrikeAndVerticalPosition(string text, int textOffset, int papxFkpOffset, int styleSheetLength) {
                const int fibLength = 0x1AA;
                byte[] textBytes = System.Text.Encoding.Unicode.GetBytes(text);
                var stream = new byte[Math.Max(papxFkpOffset + 512, textOffset + textBytes.Length)];
                WriteUInt16(stream, 0x00, 0xA5EC);
                WriteUInt16(stream, 0x02, 0x00D9);
                WriteUInt16(stream, 0x0A, 0x0200);
                WriteInt32(stream, 0x4C, text.Length);
                WriteInt32(stream, 0xA2, 34);
                WriteInt32(stream, 0xA6, styleSheetLength);
                WriteInt32(stream, 0x102, 21);
                WriteInt32(stream, 0x106, 12);
                WriteInt32(stream, 0x1A2, 0);
                WriteInt32(stream, 0x1A6, 21);
                Buffer.BlockCopy(textBytes, 0, stream, textOffset, textBytes.Length);

                int secondParagraphStart = textOffset + ("caps style\r".Length * 2);
                int thirdParagraphStart = secondParagraphStart + ("small style\r".Length * 2);
                int fourthParagraphStart = thirdParagraphStart + ("super style\r".Length * 2);
                int end = fourthParagraphStart + ("sub style\r".Length * 2);
                WritePapxFkp(
                    stream,
                    papxFkpOffset,
                    new[] { textOffset, secondParagraphStart, thirdParagraphStart, fourthParagraphStart, end },
                    new Dictionary<int, byte[]> {
                        [0] = CreateParagraphPropertiesPapx(CreateParagraphSprm(0x4600, 1, 0)),
                        [1] = CreateParagraphPropertiesPapx(CreateParagraphSprm(0x4600, 2, 0)),
                        [2] = CreateParagraphPropertiesPapx(CreateParagraphSprm(0x4600, 3, 0)),
                        [3] = CreateParagraphPropertiesPapx(CreateParagraphSprm(0x4600, 4, 0))
                    });

                if (stream.Length < fibLength) {
                    Array.Resize(ref stream, fibLength);
                }

                return stream;
            }

            private static byte[] CreateUnicodeWordDocumentStreamWithBuiltInStyleLevelFormatting(string text, int textOffset, int papxFkpOffset, int styleSheetLength) {
                return CreateUnicodeWordDocumentStreamWithStyleIndex(text, textOffset, papxFkpOffset, styleSheetLength, 1);
            }

            private static byte[] CreateUnicodeWordDocumentStreamWithStyleIndex(string text, int textOffset, int papxFkpOffset, int styleSheetLength, ushort styleIndex) {
                const int fibLength = 0x1AA;
                byte[] textBytes = System.Text.Encoding.Unicode.GetBytes(text);
                var stream = new byte[Math.Max(papxFkpOffset + 512, textOffset + textBytes.Length)];
                WriteUInt16(stream, 0x00, 0xA5EC);
                WriteUInt16(stream, 0x02, 0x00D9);
                WriteUInt16(stream, 0x0A, 0x0200);
                WriteInt32(stream, 0x4C, text.Length);
                WriteInt32(stream, 0xA2, 34);
                WriteInt32(stream, 0xA6, styleSheetLength);
                WriteInt32(stream, 0x102, 21);
                WriteInt32(stream, 0x106, 12);
                WriteInt32(stream, 0x1A2, 0);
                WriteInt32(stream, 0x1A6, 21);
                Buffer.BlockCopy(textBytes, 0, stream, textOffset, textBytes.Length);

                WritePapxFkp(
                    stream,
                    papxFkpOffset,
                    new[] { textOffset, textOffset + textBytes.Length },
                    new Dictionary<int, byte[]> {
                        [0] = CreateParagraphPropertiesPapx(
                            CreateParagraphSprm(
                                0x4600,
                                (byte)(styleIndex & 0xFF),
                                (byte)(styleIndex >> 8)))
                    });

                if (stream.Length < fibLength) {
                    Array.Resize(ref stream, fibLength);
                }

                return stream;
            }

            private static byte[] CreateTableStream(int characterCount) {
                const int textOffset = 0x200;
                var table = new byte[21];
                table[0] = 0x02;
                WriteInt32(table, 1, 16);
                WriteInt32(table, 5, 0);
                WriteInt32(table, 9, characterCount);
                WriteUInt16(table, 13, 0);
                WriteUInt32(table, 15, 0x40000000U | ((uint)textOffset * 2U));
                WriteUInt16(table, 19, 0);
                return table;
            }

            private static byte[] CreateUnicodeTableStreamWithParagraphBinTable(int characterCount, int textOffset, int papxFkpPageNumber) {
                var table = new byte[33];
                table[0] = 0x02;
                WriteInt32(table, 1, 16);
                WriteInt32(table, 5, 0);
                WriteInt32(table, 9, characterCount);
                WriteUInt16(table, 13, 0);
                WriteUInt32(table, 15, unchecked((uint)textOffset));
                WriteUInt16(table, 19, 0);

                int papxPlcOffset = 21;
                WriteInt32(table, papxPlcOffset, textOffset);
                WriteInt32(table, papxPlcOffset + 4, textOffset + (characterCount * 2));
                WriteInt32(table, papxPlcOffset + 8, papxFkpPageNumber);
                return table;
            }

            private static byte[] CreateUnicodeTableStreamWithParagraphBinTableAndStyleSheet(int characterCount, int textOffset, int papxFkpPageNumber, byte[] styleSheet) {
                byte[] table = CreateUnicodeTableStreamWithParagraphBinTable(characterCount, textOffset, papxFkpPageNumber);
                Array.Resize(ref table, table.Length + 1 + styleSheet.Length);
                Buffer.BlockCopy(styleSheet, 0, table, 34, styleSheet.Length);
                return table;
            }

            private static byte[] CreateUnicodeTableStreamWithCharacterBinTableAndFontTable(int characterCount, int textOffset, int chpxFkpPageNumber, byte[] fontTable) {
                byte[] table = CreateUnicodeTableStreamWithCharacterBinTable(characterCount, textOffset, chpxFkpPageNumber);
                Array.Resize(ref table, table.Length + fontTable.Length);
                Buffer.BlockCopy(fontTable, 0, table, 33, fontTable.Length);
                return table;
            }

            private static byte[] CreateFontTable(params string[] fontFamilies) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, checked((ushort)fontFamilies.Length));
                WriteUInt16(stream, 0);
                foreach (string fontFamily in fontFamilies) {
                    byte[] ffn = CreateFfn(fontFamily);
                    stream.WriteByte(checked((byte)ffn.Length));
                    stream.Write(ffn, 0, ffn.Length);
                }

                return stream.ToArray();
            }

            private static byte[] CreateFfn(string fontFamily) {
                byte[] nameBytes = System.Text.Encoding.Unicode.GetBytes(fontFamily + '\0');
                var ffn = new byte[39 + nameBytes.Length];
                ffn[1] = 0x90;
                ffn[2] = 0x01;
                Buffer.BlockCopy(nameBytes, 0, ffn, 39, nameBytes.Length);
                return ffn;
            }

            private static byte[] CreateUnicodeTableStreamWithCharacterBinTable(int characterCount, int textOffset, int chpxFkpPageNumber) {
                var table = new byte[33];
                table[0] = 0x02;
                WriteInt32(table, 1, 16);
                WriteInt32(table, 5, 0);
                WriteInt32(table, 9, characterCount);
                WriteUInt16(table, 13, 0);
                WriteUInt32(table, 15, unchecked((uint)textOffset));
                WriteUInt16(table, 19, 0);

                int chpxPlcOffset = 21;
                WriteInt32(table, chpxPlcOffset, textOffset);
                WriteInt32(table, chpxPlcOffset + 4, textOffset + (characterCount * 2));
                WriteInt32(table, chpxPlcOffset + 8, chpxFkpPageNumber);
                return table;
            }

            private static byte[] CreateStyleSheet(params byte[][] styleRecords) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 4);
                WriteUInt16(stream, checked((ushort)styleRecords.Length));
                WriteUInt16(stream, 8);

                foreach (byte[] styleRecord in styleRecords) {
                    WriteUInt16(stream, checked((ushort)styleRecord.Length));
                    stream.Write(styleRecord, 0, styleRecord.Length);
                    if ((stream.Position & 1) != 0) {
                        stream.WriteByte(0);
                    }
                }

                return stream.ToArray();
            }

            private static byte[] CreateParagraphStyleRecord(ushort sti, ushort baseStyleIndex, string name, params byte[][] upxs) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, sti);
                WriteUInt16(stream, checked((ushort)((baseStyleIndex << 4) | 1)));
                WriteUInt16(stream, checked((ushort)upxs.Length));
                WriteUInt16(stream, 0);
                WriteXstz(stream, name);

                foreach (byte[] upx in upxs) {
                    WriteUInt16(stream, checked((ushort)upx.Length));
                    stream.Write(upx, 0, upx.Length);
                    if ((stream.Position & 1) != 0) {
                        stream.WriteByte(0);
                    }
                }

                return stream.ToArray();
            }

            private static byte[] CreateStyleParagraphFormatting(params byte[][] sprms) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0);
                foreach (byte[] sprm in sprms) {
                    stream.Write(sprm, 0, sprm.Length);
                }

                return stream.ToArray();
            }

            private static byte[] CreateStyleCharacterFormatting(params byte[][] sprms) {
                using var stream = new MemoryStream();
                foreach (byte[] sprm in sprms) {
                    stream.Write(sprm, 0, sprm.Length);
                }

                return stream.ToArray();
            }

            private static byte[] CreateCharacterSprm(ushort sprm, params byte[] operand) {
                return CreateParagraphSprm(sprm, operand);
            }

            private static void WriteXstz(Stream stream, string value) {
                WriteUInt16(stream, checked((ushort)value.Length));
                byte[] bytes = System.Text.Encoding.Unicode.GetBytes(value);
                stream.Write(bytes, 0, bytes.Length);
                WriteUInt16(stream, 0);
            }

            private static void WriteChpxFkp(byte[] stream, int fkpOffset, int[] fileCharacterPositions, int boldRunIndex, int italicRunIndex) {
                const int boldChpxOffset = 0xF0;
                const int italicChpxOffset = 0xF8;
                int runCount = fileCharacterPositions.Length - 1;
                for (int i = 0; i < fileCharacterPositions.Length; i++) {
                    WriteInt32(stream, fkpOffset + (i * 4), fileCharacterPositions[i]);
                }

                int rgbOffset = fkpOffset + (fileCharacterPositions.Length * 4);
                for (int i = 0; i < runCount; i++) {
                    if (i == boldRunIndex) {
                        stream[rgbOffset + i] = boldChpxOffset / 2;
                    } else if (i == italicRunIndex) {
                        stream[rgbOffset + i] = italicChpxOffset / 2;
                    }
                }

                WriteSingleToggleChpx(stream, fkpOffset + boldChpxOffset, 0x0835);
                WriteSingleToggleChpx(stream, fkpOffset + italicChpxOffset, 0x0836);
                stream[fkpOffset + 511] = checked((byte)runCount);
            }

            private static void WriteChpxFkp(byte[] stream, int fkpOffset, int[] fileCharacterPositions, IReadOnlyDictionary<int, byte[]> chpxByRunIndex) {
                int runCount = fileCharacterPositions.Length - 1;
                for (int i = 0; i < fileCharacterPositions.Length; i++) {
                    WriteInt32(stream, fkpOffset + (i * 4), fileCharacterPositions[i]);
                }

                int rgbOffset = fkpOffset + (fileCharacterPositions.Length * 4);
                int chpxOffset = 0xE0;
                for (int i = 0; i < runCount; i++) {
                    if (!chpxByRunIndex.TryGetValue(i, out byte[]? chpx)) {
                        continue;
                    }

                    chpxOffset = AlignToEven(chpxOffset);
                    stream[rgbOffset + i] = checked((byte)(chpxOffset / 2));
                    Buffer.BlockCopy(chpx, 0, stream, fkpOffset + chpxOffset, chpx.Length);
                    chpxOffset += chpx.Length;
                }

                stream[fkpOffset + 511] = checked((byte)runCount);
            }

            private static void WritePapxFkp(byte[] stream, int fkpOffset, int[] fileParagraphPositions, IReadOnlyDictionary<int, byte[]> papxByParagraphIndex) {
                const int bxLength = 13;
                int paragraphCount = fileParagraphPositions.Length - 1;
                for (int i = 0; i < fileParagraphPositions.Length; i++) {
                    WriteInt32(stream, fkpOffset + (i * 4), fileParagraphPositions[i]);
                }

                int rgbxOffset = fkpOffset + (fileParagraphPositions.Length * 4);
                int papxOffset = 0x180;
                for (int i = 0; i < paragraphCount; i++) {
                    if (!papxByParagraphIndex.TryGetValue(i, out byte[]? papx)) {
                        continue;
                    }

                    papxOffset = AlignToEven(papxOffset);
                    stream[rgbxOffset + (i * bxLength)] = checked((byte)(papxOffset / 2));
                    Buffer.BlockCopy(papx, 0, stream, fkpOffset + papxOffset, papx.Length);
                    papxOffset += papx.Length;
                }

                stream[fkpOffset + 511] = checked((byte)paragraphCount);
            }

            private static void WriteSingleToggleChpx(byte[] stream, int offset, ushort sprm) {
                stream[offset] = 3;
                WriteUInt16(stream, offset + 1, sprm);
                stream[offset + 3] = 1;
            }

            private static byte[] CreateSingleSprmChpx(ushort sprm, params byte[] operand) {
                var chpx = new byte[3 + operand.Length];
                chpx[0] = checked((byte)(2 + operand.Length));
                WriteUInt16(chpx, 1, sprm);
                Buffer.BlockCopy(operand, 0, chpx, 3, operand.Length);
                return chpx;
            }

            private static byte[] CreateParagraphAlignmentPapx(byte alignment) {
                return CreateParagraphPropertiesPapx(CreateParagraphSprm(0x2461, alignment));
            }

            private static byte[] CreateParagraphPropertiesPapx(params byte[][] sprms) {
                var grpprl = new List<byte> {
                    0,
                    0
                };

                foreach (byte[] sprm in sprms) {
                    grpprl.AddRange(sprm);
                }

                if (grpprl.Count % 2 != 0) {
                    grpprl.Add(0);
                }

                var papx = new byte[grpprl.Count + 2];
                papx[0] = 0;
                papx[1] = checked((byte)(grpprl.Count / 2));
                grpprl.CopyTo(papx, 2);
                return papx;
            }

            private static byte[] CreateParagraphSprm(ushort sprm, params byte[] operand) {
                var bytes = new byte[2 + operand.Length];
                WriteUInt16(bytes, 0, sprm);
                Buffer.BlockCopy(operand, 0, bytes, 2, operand.Length);
                return bytes;
            }

            private static int[] GetFirstMarkerEnds(string text, int textOffset, int markerCount) {
                var markerEnds = new List<int>(markerCount);
                for (int index = 0; index < text.Length && markerEnds.Count < markerCount; index++) {
                    if (text[index] == '\a') {
                        markerEnds.Add(textOffset + ((index + 1) * 2));
                    }
                }

                if (markerEnds.Count != markerCount) {
                    throw new InvalidOperationException("The synthetic DOC table fixture does not contain the expected table markers.");
                }

                return markerEnds.ToArray();
            }

            private static byte[] CreateParagraphTabStopsSprm(int[] clearPositions, params (int Position, byte Alignment, byte Leader)[] addedTabStops) {
                var operand = new List<byte>();
                operand.Add(checked((byte)clearPositions.Length));
                foreach (int position in clearPositions) {
                    AddInt16(operand, position);
                }

                operand.Add(checked((byte)addedTabStops.Length));
                foreach ((int Position, byte Alignment, byte Leader) tabStop in addedTabStops) {
                    AddInt16(operand, tabStop.Position);
                }

                foreach ((int Position, byte Alignment, byte Leader) tabStop in addedTabStops) {
                    operand.Add((byte)(tabStop.Alignment | (tabStop.Leader << 3)));
                }

                if (operand.Count > byte.MaxValue) {
                    throw new InvalidOperationException("Test tab-stop operand is too large.");
                }

                return CreateParagraphSprm(0xC60D, new[] { checked((byte)operand.Count) }.Concat(operand).ToArray());
            }

            private static byte[] CreateTableDefinitionSprm(IReadOnlyList<int> cellWidthsTwips) {
                var remainder = new List<byte>();
                remainder.Add(checked((byte)cellWidthsTwips.Count));
                AddInt16(remainder, 0);
                int edge = 0;
                foreach (int width in cellWidthsTwips) {
                    edge = checked(edge + width);
                    AddInt16(remainder, edge);
                }

                for (int index = 0; index < cellWidthsTwips.Count; index++) {
                    for (int byteIndex = 0; byteIndex < 20; byteIndex++) {
                        remainder.Add(0);
                    }
                }

                int cb = checked(remainder.Count + 1);
                var operand = new List<byte> {
                    (byte)(cb & 0xFF),
                    (byte)(cb >> 8)
                };
                operand.AddRange(remainder);
                return CreateParagraphSprm(0xD608, operand.ToArray());
            }

            private static void AddInt16(List<byte> bytes, int value) {
                short signed = checked((short)value);
                bytes.Add((byte)(signed & 0xFF));
                bytes.Add((byte)(signed >> 8));
            }

            private static int AlignToEven(int value) {
                return value % 2 == 0 ? value : value + 1;
            }

            private static byte[] CreateSummaryInformationPropertySet(DateTime created, DateTime modified) {
                var properties = new List<OleTestProperty> {
                    OleTestProperty.Int16(1, 1200),
                    OleTestProperty.String(2, "Legacy DOC Metadata Title"),
                    OleTestProperty.String(3, "Legacy DOC metadata subject"),
                    OleTestProperty.String(4, "OfficeIMO Legacy Import"),
                    OleTestProperty.String(5, "doc, metadata, officeimo"),
                    OleTestProperty.String(6, "OLE SummaryInformation comments"),
                    OleTestProperty.FileTime(12, created),
                    OleTestProperty.FileTime(13, modified)
                };

                return CreateOlePropertySet(CreateOlePropertySection(properties));
            }

            private static byte[] CreateDocumentSummaryInformationPropertySet() {
                var documentProperties = new List<OleTestProperty> {
                    OleTestProperty.Int16(1, 1200),
                    OleTestProperty.String(2, "Legacy Category"),
                    OleTestProperty.String(14, "Document Manager"),
                    OleTestProperty.String(15, "EvotecIT")
                };
                var customProperties = new List<OleTestProperty> {
                    OleTestProperty.Int16(1, 1200),
                    OleTestProperty.Dictionary(0, new Dictionary<uint, string> {
                        [2] = "ReleaseStatus",
                        [3] = "Reviewed",
                        [4] = "Ticket"
                    }),
                    OleTestProperty.String(2, "Ready"),
                    OleTestProperty.Boolean(3, true),
                    OleTestProperty.Int32(4, 2003)
                };

                return CreateOlePropertySet(CreateOlePropertySection(documentProperties), CreateOlePropertySection(customProperties));
            }

            private static byte[] CreateOlePropertySet(params byte[][] sections) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0xfffe);
                WriteUInt16(stream, 0);
                WriteUInt32(stream, 0);
                stream.Write(new byte[16], 0, 16);
                WriteUInt32(stream, checked((uint)sections.Length));

                int sectionOffset = 28 + sections.Length * 20;
                foreach (byte[] section in sections) {
                    stream.Write(new byte[16], 0, 16);
                    WriteUInt32(stream, checked((uint)sectionOffset));
                    sectionOffset += section.Length;
                }

                foreach (byte[] section in sections) {
                    stream.Write(section, 0, section.Length);
                }

                return stream.ToArray();
            }

            private static byte[] CreateOlePropertySection(IReadOnlyList<OleTestProperty> properties) {
                using var values = new MemoryStream();
                var offsets = new List<uint>(properties.Count);
                foreach (OleTestProperty property in properties) {
                    offsets.Add(checked((uint)(8 + properties.Count * 8 + values.Length)));
                    values.Write(property.ValueBytes, 0, property.ValueBytes.Length);
                    PadToInt32(values);
                }

                using var stream = new MemoryStream();
                WriteUInt32(stream, checked((uint)(8 + properties.Count * 8 + values.Length)));
                WriteUInt32(stream, checked((uint)properties.Count));
                for (int i = 0; i < properties.Count; i++) {
                    WriteUInt32(stream, properties[i].PropertyId);
                    WriteUInt32(stream, offsets[i]);
                }

                byte[] valueBytes = values.ToArray();
                stream.Write(valueBytes, 0, valueBytes.Length);
                return stream.ToArray();
            }

            private static void WriteStream(RootStorage root, string name, byte[] bytes) {
                using CfbStream stream = root.CreateStream(name);
                stream.Write(bytes, 0, bytes.Length);
            }

            private static byte[] EncodeWindows1252(string text) {
                var bytes = new byte[text.Length];
                for (int i = 0; i < text.Length; i++) {
                    char character = text[i];
                    bytes[i] = character <= 0x7F || (character >= 0xA0 && character <= 0xFF)
                        ? (byte)character
                        : (byte)'?';
                }

                return bytes;
            }

            private static void PadToInt32(Stream stream) {
                while (stream.Position % 4 != 0) {
                    stream.WriteByte(0);
                }
            }

            private static void WriteUInt16(Stream stream, ushort value) {
                stream.WriteByte((byte)(value & 0xff));
                stream.WriteByte((byte)((value >> 8) & 0xff));
            }

            private static void WriteUInt32(Stream stream, uint value) {
                stream.WriteByte((byte)(value & 0xff));
                stream.WriteByte((byte)((value >> 8) & 0xff));
                stream.WriteByte((byte)((value >> 16) & 0xff));
                stream.WriteByte((byte)((value >> 24) & 0xff));
            }

            private static void WriteUInt64(Stream stream, ulong value) {
                WriteUInt32(stream, unchecked((uint)(value & 0xffffffffUL)));
                WriteUInt32(stream, unchecked((uint)(value >> 32)));
            }

            private static void WriteUInt16(byte[] bytes, int offset, ushort value) {
                bytes[offset] = (byte)value;
                bytes[offset + 1] = (byte)(value >> 8);
            }

            private static void WriteInt32(byte[] bytes, int offset, int value) {
                bytes[offset] = (byte)value;
                bytes[offset + 1] = (byte)(value >> 8);
                bytes[offset + 2] = (byte)(value >> 16);
                bytes[offset + 3] = (byte)(value >> 24);
            }

            private static void WriteUInt32(byte[] bytes, int offset, uint value) {
                bytes[offset] = (byte)value;
                bytes[offset + 1] = (byte)(value >> 8);
                bytes[offset + 2] = (byte)(value >> 16);
                bytes[offset + 3] = (byte)(value >> 24);
            }

            private readonly struct OleTestProperty {
                private OleTestProperty(uint propertyId, byte[] valueBytes) {
                    PropertyId = propertyId;
                    ValueBytes = valueBytes;
                }

                internal uint PropertyId { get; }

                internal byte[] ValueBytes { get; }

                internal static OleTestProperty Int16(uint id, short value) {
                    using var stream = new MemoryStream();
                    WriteUInt16(stream, 0x0002);
                    WriteUInt16(stream, 0);
                    WriteUInt16(stream, unchecked((ushort)value));
                    WriteUInt16(stream, 0);
                    return new OleTestProperty(id, stream.ToArray());
                }

                internal static OleTestProperty Int32(uint id, int value) {
                    using var stream = new MemoryStream();
                    WriteUInt16(stream, 0x0003);
                    WriteUInt16(stream, 0);
                    WriteUInt32(stream, unchecked((uint)value));
                    return new OleTestProperty(id, stream.ToArray());
                }

                internal static OleTestProperty Boolean(uint id, bool value) {
                    using var stream = new MemoryStream();
                    WriteUInt16(stream, 0x000b);
                    WriteUInt16(stream, 0);
                    WriteUInt16(stream, value ? (ushort)0xffff : (ushort)0);
                    WriteUInt16(stream, 0);
                    return new OleTestProperty(id, stream.ToArray());
                }

                internal static OleTestProperty FileTime(uint id, DateTime value) {
                    using var stream = new MemoryStream();
                    WriteUInt16(stream, 0x0040);
                    WriteUInt16(stream, 0);
                    WriteUInt64(stream, unchecked((ulong)value.ToUniversalTime().ToFileTimeUtc()));
                    return new OleTestProperty(id, stream.ToArray());
                }

                internal static OleTestProperty String(uint id, string value) {
                    using var stream = new MemoryStream();
                    WriteUInt16(stream, 0x001f);
                    WriteUInt16(stream, 0);
                    WriteUInt32(stream, checked((uint)(value.Length + 1)));
                    byte[] bytes = System.Text.Encoding.Unicode.GetBytes(value + '\0');
                    stream.Write(bytes, 0, bytes.Length);
                    PadToInt32(stream);
                    return new OleTestProperty(id, stream.ToArray());
                }

                internal static OleTestProperty Dictionary(uint id, IReadOnlyDictionary<uint, string> names) {
                    using var stream = new MemoryStream();
                    WriteUInt32(stream, checked((uint)names.Count));
                    foreach (KeyValuePair<uint, string> name in names.OrderBy(entry => entry.Key)) {
                        WriteUInt32(stream, name.Key);
                        WriteUInt32(stream, checked((uint)(name.Value.Length + 1)));
                        byte[] bytes = System.Text.Encoding.Unicode.GetBytes(name.Value + '\0');
                        stream.Write(bytes, 0, bytes.Length);
                        PadToInt32(stream);
                    }

                    return new OleTestProperty(id, stream.ToArray());
                }
            }
        }

        private static void AssertSameInstant(DateTime expected, DateTime? actual) {
            Assert.NotNull(actual);
            Assert.Equal(expected.ToUniversalTime(), actual.Value.ToUniversalTime());
        }

        private static string NormalizeLegacyDocBaselineText(string text) {
            return text.Replace("\r\n", "\n").Replace('\r', '\n').TrimEnd() + "\n";
        }

        private static string GetRelativePath(string relativeTo, string path) {
            Uri baseUri = new Uri(AppendDirectorySeparator(Path.GetFullPath(relativeTo)));
            Uri pathUri = new Uri(Path.GetFullPath(path));
            string relative = Uri.UnescapeDataString(baseUri.MakeRelativeUri(pathUri).ToString());
            return relative.Replace('/', Path.DirectorySeparatorChar);
        }

        private static string AppendDirectorySeparator(string path) {
            if (path.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)
                || path.EndsWith(Path.AltDirectorySeparatorChar.ToString(), StringComparison.Ordinal)) {
                return path;
            }

            return path + Path.DirectorySeparatorChar;
        }

        private static string GetWordTestsProjectRoot() {
            var directory = new DirectoryInfo(AppContext.BaseDirectory);
            while (directory != null) {
                if (File.Exists(Path.Combine(directory.FullName, "OfficeIMO.Tests.csproj"))) {
                    return directory.FullName;
                }

                directory = directory.Parent;
            }

            return AppContext.BaseDirectory;
        }

        private static void DeleteIfExists(string path) {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }

        private static byte[] ReadCompoundStream(byte[] compoundBytes, string streamName) {
            Assert.True(
                OfficeCompoundFileReader.TryRead(compoundBytes, out OfficeCompoundFile? compoundFile, out string? error),
                error);
            Assert.True(compoundFile!.Streams.TryGetValue(streamName, out byte[]? stream), $"Compound stream '{streamName}' was not found.");
            return stream!;
        }

        private static bool ContainsBytePattern(byte[] bytes, params byte[] pattern) {
            for (int offset = 0; offset <= bytes.Length - pattern.Length; offset++) {
                bool match = true;
                for (int index = 0; index < pattern.Length; index++) {
                    if (bytes[offset + index] != pattern[index]) {
                        match = false;
                        break;
                    }
                }

                if (match) {
                    return true;
                }
            }

            return false;
        }

        private static SectionMarkValues? GetParagraphSectionType(WordDocument document) {
            return document._wordprocessingDocument.MainDocumentPart!.Document.Body!
                .Elements<Paragraph>()
                .Select(paragraph => paragraph.ParagraphProperties?.SectionProperties?.GetFirstChild<SectionType>()?.Val?.Value)
                .FirstOrDefault(value => value != null);
        }

        private static SectionMarkValues GetSectionMarkValue(string key) {
            switch (key) {
                case "continuous":
                    return SectionMarkValues.Continuous;
                case "nextColumn":
                    return SectionMarkValues.NextColumn;
                case "nextPage":
                    return SectionMarkValues.NextPage;
                case "evenPage":
                    return SectionMarkValues.EvenPage;
                case "oddPage":
                    return SectionMarkValues.OddPage;
                default:
                    throw new ArgumentOutOfRangeException(nameof(key), key, "Unsupported section mark test key.");
            }
        }
    }
}
