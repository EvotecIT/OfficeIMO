using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfParsedParagraphStateTests {
    private const string Styles = @"{\stylesheet{\s0 Normal;}{\s1\sl360\slmult0\tqr\tldot\tx4320 SpacedTabs;}}";

    [Theory]
    [InlineData("none")]
    [InlineData("unused")]
    [InlineData("selected")]
    public void Parsed_Implicit_Spacing_Is_Single_With_Or_Without_A_Stylesheet(string stylesheet) {
        string header = stylesheet == "none" ? string.Empty : Styles;
        string selection = stylesheet == "selected" ? @"\s1" : string.Empty;
        RtfDocument source = RtfDocument.Read(@"{\rtf1\ansi" + header + @"\pard" + selection + @" Plain\par}").Document;
        RtfParagraph paragraph = Assert.Single(source.Paragraphs);
        RtfParagraph effective = source.ResolveParagraphFormatting(paragraph);

        Assert.Null(paragraph.LineSpacingTwips);
        Assert.Null(paragraph.LineSpacingMultiple);
        Assert.Empty(paragraph.TabStops);
        Assert.Null(paragraph.SpaceBeforeTwips);
        Assert.Null(paragraph.SpaceAfterTwips);
        Assert.Equal(0, effective.SpaceBeforeTwips);
        Assert.Equal(0, effective.SpaceAfterTwips);
        Assert.Equal(240, effective.LineSpacingTwips);
        Assert.True(effective.LineSpacingMultiple);
        Assert.Empty(effective.TabStops);
        PdfCore.RichParagraphBlock block = Assert.IsType<PdfCore.RichParagraphBlock>(Assert.Single(source.ToPdf().Blocks));
        PdfCore.PdfParagraphStyle pdfStyle = Assert.IsType<PdfCore.PdfParagraphStyle>(block.Style);
        PdfCore.PdfLineSpacing pdfSpacing = Assert.IsType<PdfCore.PdfLineSpacing>(pdfStyle.LineSpacing);
        Assert.Equal(PdfCore.PdfLineSpacingRule.Multiple, pdfSpacing.Rule);

        using WordDocument word = source.ToWordDocument();
        using WordDocument reopened = WordDocument.Load(new MemoryStream(word.ToBytes()));
        Paragraph persisted = Assert.Single(reopened._document.Body!.Elements<Paragraph>());
        AssertWordSpacing(persisted, "240", LineSpacingRuleValues.Auto);
        if (stylesheet == "selected") {
            Assert.Equal("RtfP1", persisted.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
            AssertWordTab(Assert.Single(persisted.ParagraphProperties!.Tabs!.Elements<TabStop>()), 4320, TabStopValues.Clear);
        } else {
            Assert.Empty(persisted.ParagraphProperties?.Tabs?.Elements<TabStop>() ?? Enumerable.Empty<TabStop>());
        }
        Assert.Null(paragraph.LineSpacingTwips);
        Assert.Null(paragraph.LineSpacingMultiple);
        Assert.Empty(paragraph.TabStops);
    }

    [Fact]
    public void Authored_Styles_Inherit_While_Parsed_Paragraph_State_Remains_Authoritative() {
        RtfDocument authored = RtfDocument.Create();
        RtfStyle authoredStyle = authored.AddStyle(1, "SpacedTabs");
        authoredStyle.LineSpacingTwips = 360;
        authoredStyle.LineSpacingMultiple = false;
        authoredStyle.SpaceBeforeTwips = 120;
        authoredStyle.SpaceAfterTwips = 240;
        authoredStyle.AddTabStop(4320, RtfTabAlignment.Right, RtfTabLeader.Dots);
        RtfParagraph authoredParagraph = authored.AddParagraph("Plain").SetStyle(1);
        RtfDocument parsed = RtfDocument.Read(@"{\rtf1\ansi" + Styles + @"\pard\s1 Plain\par}").Document;
        RtfParagraph parsedParagraph = Assert.Single(parsed.Paragraphs);

        RtfParagraph authoredEffective = authored.ResolveParagraphFormatting(authoredParagraph);
        Assert.Equal(360, authoredEffective.LineSpacingTwips);
        Assert.False(authoredEffective.LineSpacingMultiple);
        Assert.Equal(120, authoredEffective.SpaceBeforeTwips);
        Assert.Equal(240, authoredEffective.SpaceAfterTwips);
        Assert.Equal(4320, Assert.Single(authoredEffective.TabStops).PositionTwips);
        Assert.Equal(240, parsed.ResolveParagraphFormatting(parsedParagraph).LineSpacingTwips);
        Assert.Empty(parsed.ResolveParagraphFormatting(parsedParagraph).TabStops);
        authoredStyle.LineSpacingTwips = 720;
        authoredStyle.SpaceAfterTwips = 480;
        Assert.Equal(480, authored.ResolveParagraphFormatting(authoredParagraph).SpaceAfterTwips);
        Assert.Single(parsed.Styles, style => style.Id == 1).LineSpacingTwips = 720;
        Assert.Equal(720, authored.ResolveParagraphFormatting(authoredParagraph).LineSpacingTwips);
        Assert.Equal(240, parsed.ResolveParagraphFormatting(parsedParagraph).LineSpacingTwips);
        Assert.Null(authoredParagraph.LineSpacingTwips);
        Assert.Null(authoredParagraph.SpaceBeforeTwips);
        Assert.Null(authoredParagraph.SpaceAfterTwips);
        Assert.Null(parsedParagraph.LineSpacingTwips);
        using WordDocument authoredWord = authored.ToWordDocument();
        using WordDocument reopened = WordDocument.Load(new MemoryStream(authoredWord.ToBytes()));
        SpacingBetweenLines persisted = Assert.IsType<SpacingBetweenLines>(Assert.Single(reopened._document.Body!.Elements<Paragraph>()).ParagraphProperties?.SpacingBetweenLines);
        Assert.Equal("120", persisted.Before?.Value);
        Assert.Equal("480", persisted.After?.Value);
    }

    [Theory]
    [InlineData(@"\sl-360\slmult0", -360, false, WordLineSpacingRule.Exact)]
    [InlineData(@"\sl360\slmult0", 360, false, WordLineSpacingRule.AtLeast)]
    [InlineData(@"\sl360\slmult1", 360, true, WordLineSpacingRule.Auto)]
    public void Parsed_Direct_Line_And_Tab_Controls_Override_Style_Declarations(
        string controls, int line, bool multiple, WordLineSpacingRule rule) {
        RtfDocument source = RtfDocument.Read(@"{\rtf1\ansi" + Styles + @"\pard\s1" + controls +
            @"\tqdec\tlhyph\tx2880 Direct\tab 12.5\par}").Document;
        RtfParagraph paragraph = Assert.Single(source.Paragraphs);
        RtfParagraph effective = source.ResolveParagraphFormatting(paragraph);
        Assert.Equal(line, paragraph.LineSpacingTwips);
        Assert.Equal(multiple, paragraph.LineSpacingMultiple);
        Assert.Equal(line, effective.LineSpacingTwips);
        Assert.Equal(multiple, effective.LineSpacingMultiple);
        RtfTabStop tab = Assert.Single(effective.TabStops);
        Assert.Equal(2880, tab.PositionTwips);
        Assert.Equal(RtfTabAlignment.Decimal, tab.Alignment);
        Assert.Equal(RtfTabLeader.Hyphen, tab.Leader);

        using WordDocument word = source.ToWordDocument();
        using WordDocument reopened = WordDocument.Load(new MemoryStream(word.ToBytes()));
        Paragraph persisted = Assert.Single(reopened._document.Body!.Elements<Paragraph>());
        LineSpacingRuleValues wordRule = rule == WordLineSpacingRule.Exact ? LineSpacingRuleValues.Exact
            : rule == WordLineSpacingRule.AtLeast ? LineSpacingRuleValues.AtLeast : LineSpacingRuleValues.Auto;
        AssertWordSpacing(persisted, "360", wordRule);
        TabStop[] tabs = Assert.IsType<Tabs>(persisted.ParagraphProperties?.Tabs).Elements<TabStop>().ToArray();
        Assert.Equal(2, tabs.Length);
        AssertWordTab(Assert.Single(tabs, stop => stop.Position?.Value == 4320), 4320, TabStopValues.Clear);
        TabStop active = Assert.Single(tabs, stop => stop.Position?.Value == 2880);
        AssertWordTab(active, 2880, TabStopValues.Decimal);
        Assert.Equal(TabStopLeaderCharValues.Hyphen, active.Leader?.Value);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Parsed_Default_State_Survives_Clone_Html_Metadata_And_Materialized_Rtf(bool hasStyles) {
        string styleControls = hasStyles ? Styles + @"\pard\s1" : @"\pard";
        RtfDocument source = RtfDocument.Read(@"{\rtf1\ansi" + styleControls + @" Plain\par}").Document;
        RtfDocument cloned = source.Clone();
        AssertParsedState(cloned);
        string html = source.ToHtml(new RtfToHtmlOptions { IncludeRoundTripMetadata = true, FragmentOnly = false });
        RtfDocument fromHtml = HtmlConversionDocument.Parse(html).ToRtfDocument();
        AssertParsedState(fromHtml);
        if (hasStyles) {
            Assert.Equal(360, Assert.Single(fromHtml.Styles, style => style.Id == 1).LineSpacingTwips);
            Assert.Equal(4320, Assert.Single(Assert.Single(fromHtml.Styles, style => style.Id == 1).TabStops).PositionTwips);
        }

        RtfDocument materialized = RtfDocument.Read(source.ToRtf()).Document;
        RtfParagraph written = Assert.Single(materialized.Paragraphs);
        Assert.Equal(240, written.LineSpacingTwips);
        Assert.True(written.LineSpacingMultiple);
        Assert.Equal(0, written.SpaceBeforeTwips);
        Assert.Equal(0, written.SpaceAfterTwips);
        Assert.Empty(materialized.ResolveParagraphFormatting(written).TabStops);
        AssertParsedState(source);

        static void AssertParsedState(RtfDocument document) {
            RtfParagraph raw = Assert.Single(document.Paragraphs);
            Assert.Null(raw.LineSpacingTwips);
            Assert.Null(raw.LineSpacingMultiple);
            Assert.Empty(raw.TabStops);
            Assert.Null(raw.SpaceBeforeTwips);
            Assert.Null(raw.SpaceAfterTwips);
            RtfParagraph effective = document.ResolveParagraphFormatting(raw);
            Assert.Equal(240, effective.LineSpacingTwips);
            Assert.True(effective.LineSpacingMultiple);
            Assert.Equal(0, effective.SpaceBeforeTwips);
            Assert.Equal(0, effective.SpaceAfterTwips);
            Assert.Empty(effective.TabStops);
        }
    }

    [Fact]
    public void Parsed_Default_Normalization_Is_Not_Reported_As_Lost_Style_Inheritance() {
        RtfDocument parsed = RtfDocument.Read(@"{\rtf1\ansi\pard Plain\par}").Document;
        parsed.ToRtfResult().RequireNoLoss();
        RtfDocument parsedStyleZero = RtfDocument.Read(@"{\rtf1\ansi{\stylesheet{\s0\sb0\sa0 Normal;}}\pard Plain\par}").Document;
        Assert.Contains(parsedStyleZero.ToRtfResult().Report.Diagnostics,
            diagnostic => diagnostic.Code == "RtfNormalizationFormattingInheritanceMaterialized");
        RtfDocument authored = RtfDocument.Create();
        authored.AddStyle(1, "Minimum").LineSpacingTwips = 360;
        authored.AddParagraph("Styled").SetStyle(1);

        var output = authored.ToRtfResult();

        Assert.Contains(output.Report.Diagnostics, diagnostic => diagnostic.Code == "RtfNormalizationFormattingInheritanceMaterialized");
    }

    [Fact]
    public void Parsed_Default_State_Reaches_Cell_And_Footnote_Paragraphs() {
        RtfDocument source = RtfDocument.Read(@"{\rtf1\ansi\ftnbj" + Styles +
            @"\pard Body{\footnote\pard\s1 Note\par}\par\trowd\cellx2400\pard\intbl\s1 Cell\cell\row}").Document;
        RtfParagraph cell = Assert.Single(Assert.Single(source.Blocks.OfType<RtfTable>()).Rows[0].Cells[0].Paragraphs);
        RtfNote parsedNote = Assert.Single(source.Notes);
        Assert.Collection(parsedNote.Paragraphs,
            paragraph => Assert.Equal("Note", paragraph.ToPlainText()),
            paragraph => Assert.Empty(paragraph.Inlines));
        RtfParagraph note = parsedNote.Paragraphs[0];
        Assert.Equal(240, source.ResolveParagraphFormatting(cell).LineSpacingTwips);
        Assert.Equal(240, source.ResolveParagraphFormatting(note).LineSpacingTwips);
        Assert.Equal(0, source.ResolveParagraphFormatting(cell).SpaceBeforeTwips);
        Assert.Equal(0, source.ResolveParagraphFormatting(cell).SpaceAfterTwips);
        Assert.Equal(0, source.ResolveParagraphFormatting(note).SpaceBeforeTwips);
        Assert.Equal(0, source.ResolveParagraphFormatting(note).SpaceAfterTwips);
        Assert.Empty(source.ResolveParagraphFormatting(cell).TabStops);
        Assert.Empty(source.ResolveParagraphFormatting(note).TabStops);
        RtfParagraph terminal = source.ResolveParagraphFormatting(parsedNote.Paragraphs[1]);
        Assert.Equal(240, terminal.LineSpacingTwips);
        Assert.True(terminal.LineSpacingMultiple);
        Assert.Equal(0, terminal.SpaceBeforeTwips);
        Assert.Equal(0, terminal.SpaceAfterTwips);
        Assert.Empty(terminal.TabStops);
        using WordDocument word = source.ToWordDocument();
        using WordDocument reopened = WordDocument.Load(new MemoryStream(word.ToBytes()));
        var main = reopened._wordprocessingDocument.MainDocumentPart!;
        Document document = Assert.IsType<Document>(main.Document);
        Paragraph persistedCell = Assert.Single(document.Descendants<Paragraph>(), paragraph => paragraph.InnerText == "Cell");
        Paragraph persistedNote = Assert.Single(main.FootnotesPart!.Footnotes!.Descendants<Paragraph>(), paragraph => paragraph.InnerText == "Note");
        Footnote ordinaryNote = Assert.Single(main.FootnotesPart.Footnotes.Elements<Footnote>(), item => item.Id!.Value > 0);
        Paragraph[] persistedNoteParagraphs = ordinaryNote.Elements<Paragraph>().ToArray();
        Assert.Equal(2, persistedNoteParagraphs.Length);
        Assert.Empty(persistedNoteParagraphs[1].Descendants<Text>());
        AssertWordSpacing(persistedCell, "240", LineSpacingRuleValues.Auto);
        AssertWordSpacing(persistedNote, "240", LineSpacingRuleValues.Auto);
        AssertWordSpacing(persistedNoteParagraphs[1], "240", LineSpacingRuleValues.Auto);
        AssertWordTab(Assert.Single(persistedCell.ParagraphProperties!.Tabs!.Elements<TabStop>()), 4320, TabStopValues.Clear);
        AssertWordTab(Assert.Single(persistedNote.ParagraphProperties!.Tabs!.Elements<TabStop>()), 4320, TabStopValues.Clear);
    }

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void Parsed_Automatic_Spacing_Remains_Explicit_While_Other_Default_Gaps_Are_Zero(bool beforeAuto, bool afterAuto) {
        string controls = (beforeAuto ? @"\sbauto1" : @"\sbauto0") + (afterAuto ? @"\saauto1" : @"\saauto0");
        RtfDocument source = RtfDocument.Read(@"{\rtf1\ansi\pard" + controls + @" Plain\par}").Document;
        RtfParagraph paragraph = Assert.Single(source.Paragraphs);
        RtfParagraph effective = source.ResolveParagraphFormatting(paragraph);
        Assert.Null(paragraph.SpaceBeforeTwips);
        Assert.Null(paragraph.SpaceAfterTwips);
        Assert.Equal(beforeAuto, effective.SpaceBeforeAuto);
        Assert.Equal(afterAuto, effective.SpaceAfterAuto);
        Assert.Equal(beforeAuto ? (int?)null : 0, effective.SpaceBeforeTwips);
        Assert.Equal(afterAuto ? (int?)null : 0, effective.SpaceAfterTwips);
        using WordDocument word = source.ToWordDocument();
        using WordDocument reopened = WordDocument.Load(new MemoryStream(word.ToBytes()));
        SpacingBetweenLines spacing = Assert.IsType<SpacingBetweenLines>(Assert.Single(reopened._document.Body!.Elements<Paragraph>()).ParagraphProperties?.SpacingBetweenLines);
        Assert.Equal(beforeAuto, spacing.BeforeAutoSpacing?.Value);
        Assert.Equal(afterAuto, spacing.AfterAutoSpacing?.Value);
        Assert.Equal(beforeAuto ? null : "0", spacing.Before?.Value);
        Assert.Equal(afterAuto ? null : "0", spacing.After?.Value);
        Assert.Null(paragraph.SpaceBeforeTwips);
        Assert.Null(paragraph.SpaceAfterTwips);
    }

    private static void AssertWordSpacing(Paragraph paragraph, string line, LineSpacingRuleValues rule) {
        SpacingBetweenLines spacing = Assert.IsType<SpacingBetweenLines>(paragraph.ParagraphProperties?.SpacingBetweenLines);
        Assert.Equal(line, spacing.Line?.Value);
        Assert.Equal(rule, spacing.LineRule?.Value);
        Assert.Equal("0", spacing.Before?.Value);
        Assert.Equal("0", spacing.After?.Value);
    }

    private static void AssertWordTab(TabStop stop, int position, TabStopValues alignment) {
        Assert.Equal(position, stop.Position?.Value);
        Assert.Equal(alignment, stop.Val?.Value);
    }
}
