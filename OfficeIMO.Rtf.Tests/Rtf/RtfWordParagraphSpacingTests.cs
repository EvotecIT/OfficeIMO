using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfWordParagraphSpacingTests {
    [Theory]
    [InlineData(360, false, WordLineSpacingRule.AtLeast)]
    [InlineData(-360, false, WordLineSpacingRule.Exact)]
    [InlineData(360, true, WordLineSpacingRule.Auto)]
    [InlineData(360, null, WordLineSpacingRule.AtLeast)]
    [InlineData(-360, null, WordLineSpacingRule.Exact)]
    public void Rtf_To_Word_Persists_Signed_Line_Spacing_And_Units(int line, bool? multiple, WordLineSpacingRule rule) {
        RtfDocument source = RtfDocument.Create();
        source.AddParagraph("Spacing").SetLineSpacing(line, multiple);

        using WordDocument word = source.ToWordDocument();
        using WordDocument reopened = WordDocument.Load(new MemoryStream(word.ToBytes()));

        AssertSpacing(Assert.Single(reopened._document.Body!.Elements<Paragraph>()), Math.Abs(line), rule);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Theory]
    [InlineData(WordLineSpacingRule.AtLeast, 360, false)]
    [InlineData(WordLineSpacingRule.Exact, -360, false)]
    [InlineData(WordLineSpacingRule.Auto, 360, true)]
    [InlineData(null, 360, true)]
    public void Word_To_Rtf_Persists_Line_Spacing_Rule_In_Sign_And_Multiple(
        WordLineSpacingRule? rule, int expectedLine, bool expectedMultiple) {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph("Spacing");
        paragraph.LineSpacing = (rule) switch { global::OfficeIMO.Word.WordLineSpacingRule.Exact => global::OfficeIMO.Word.WordLineSpacing.Exact(global::OfficeIMO.OfficeLength.Emus(checked((long)(360) * 635))), global::OfficeIMO.Word.WordLineSpacingRule.AtLeast => global::OfficeIMO.Word.WordLineSpacing.AtLeast(global::OfficeIMO.OfficeLength.Emus(checked((long)(360) * 635))), _ => global::OfficeIMO.Word.WordLineSpacing.Multiple((360) / 240D) };

        if (!rule.HasValue) {
            paragraph.NativeLineSpacingRule = null;
            Assert.Null(paragraph._paragraph.ParagraphProperties!.SpacingBetweenLines!.LineRule);
        }

        string rtf = word.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfDocument parsed = RtfDocument.Read(rtf).Document;
        RtfParagraph mapped = Assert.Single(parsed.Paragraphs);

        Assert.Equal(expectedLine, mapped.LineSpacingTwips);
        Assert.Equal(expectedMultiple, mapped.LineSpacingMultiple);
        using WordDocument roundTrip = parsed.ToWordDocument();
        using WordDocument reopened = WordDocument.Load(new MemoryStream(roundTrip.ToBytes()));
        AssertSpacing(Assert.Single(reopened._document.Body!.Elements<Paragraph>()), 360, rule ?? WordLineSpacingRule.Auto);
    }

    [Theory]
    [InlineData(360, false, WordLineSpacingRule.AtLeast)]
    [InlineData(-360, false, WordLineSpacingRule.Exact)]
    [InlineData(480, true, WordLineSpacingRule.Auto)]
    public void Word_Rtf_Styles_Preserve_Signed_Spacing_And_BasedOn_Inheritance(
        int line, bool multiple, WordLineSpacingRule rule) {
        RtfDocument source = RtfDocument.Create();
        RtfStyle parent = source.AddStyle(7, "Spacing parent");
        parent.LineSpacingTwips = line;
        parent.LineSpacingMultiple = multiple;
        source.AddStyle(8, "Inherited spacing").BasedOnStyleId = 7;
        source.AddParagraph("Styled").SetStyle(8);

        using WordDocument word = source.ToWordDocument();
        using WordDocument reopened = WordDocument.Load(new MemoryStream(word.ToBytes()));
        Styles styles = reopened._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        Style mappedParent = Assert.Single(styles.Elements<Style>(), style => style.StyleId?.Value == "RtfP7");
        Style mappedChild = Assert.Single(styles.Elements<Style>(), style => style.StyleId?.Value == "RtfP8");
        AssertSpacing(mappedParent.StyleParagraphProperties!.SpacingBetweenLines!, Math.Abs(line), rule);
        Assert.Null(mappedChild.StyleParagraphProperties?.SpacingBetweenLines);
        Assert.Equal("RtfP7", mappedChild.BasedOn?.Val?.Value);
        AssertSpacing(Assert.Single(reopened._document.Body!.Elements<Paragraph>()), Math.Abs(line), rule);

        RtfDocument returned = RtfDocument.Read(reopened.ToRtf(new RtfWriteOptions { IncludeGenerator = false })).Document;
        RtfStyle returnedParent = Assert.Single(returned.Styles, style => style.Name == "RtfP7");
        RtfStyle returnedChild = Assert.Single(returned.Styles, style => style.Name == "RtfP8");
        Assert.Equal(line, returnedParent.LineSpacingTwips);
        Assert.Equal(multiple, returnedParent.LineSpacingMultiple);
        Assert.Null(returnedChild.LineSpacingTwips);
        Assert.Null(returnedChild.LineSpacingMultiple);
        Assert.Equal(returnedParent.Id, returnedChild.BasedOnStyleId);
    }

    [Fact]
    public void Word_To_Rtf_Styles_Use_Automatic_Units_When_Line_Rule_Is_Omitted() {
        using WordDocument word = WordDocument.Create();
        Style style = new Style { Type = StyleValues.Paragraph, StyleId = "AutomaticSpacing" };
        style.StyleParagraphProperties = new StyleParagraphProperties(new SpacingBetweenLines { Line = "360" });
        word._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(style);
        word.AddParagraph("Styled").SetStyleId("AutomaticSpacing");

        RtfStyle mapped = Assert.Single(word.ToRtfDocument().Styles, item => item.Name == "AutomaticSpacing");

        Assert.Equal(360, mapped.LineSpacingTwips);
        Assert.True(mapped.LineSpacingMultiple);
    }

    [Fact]
    public void Word_Rtf_Bridge_Leaves_Undeclared_Spacing_Absent() {
        RtfDocument source = RtfDocument.Create();
        source.AddStyle(7, "Spacing omitted");
        source.AddParagraph("Plain").SetStyle(7);

        using WordDocument word = source.ToWordDocument();
        using WordDocument reopened = WordDocument.Load(new MemoryStream(word.ToBytes()));
        Assert.Null(Assert.Single(reopened._document.Body!.Elements<Paragraph>()).ParagraphProperties?.SpacingBetweenLines);
        Style style = Assert.Single(reopened._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Elements<Style>(),
            item => item.StyleId?.Value == "RtfP7");
        Assert.Null(style.StyleParagraphProperties?.SpacingBetweenLines);
        RtfDocument returned = reopened.ToRtfDocument();
        Assert.Null(Assert.Single(returned.Paragraphs).LineSpacingTwips);
        Assert.Null(Assert.Single(returned.Paragraphs).LineSpacingMultiple);
        RtfStyle returnedStyle = Assert.Single(returned.Styles, item => item.Name == "RtfP7");
        Assert.Null(returnedStyle.LineSpacingTwips);
        Assert.Null(returnedStyle.LineSpacingMultiple);
    }

    [Fact]
    public void Word_Rule_Without_Line_Does_Not_Change_Inherited_Spacing_Units() {
        using WordDocument word = WordDocument.Create();
        var parent = new Style { Type = StyleValues.Paragraph, StyleId = "ExactParent" };
        parent.StyleParagraphProperties = new StyleParagraphProperties(new SpacingBetweenLines {
            Line = "360", LineRule = LineSpacingRuleValues.Exact
        });
        var child = new Style { Type = StyleValues.Paragraph, StyleId = "RuleOnlyChild" };
        child.BasedOn = new BasedOn { Val = "ExactParent" };
        child.StyleParagraphProperties = new StyleParagraphProperties(new SpacingBetweenLines { LineRule = LineSpacingRuleValues.Auto });
        word._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(parent, child);
        WordParagraph paragraph = word.AddParagraph("Inherited").SetStyleId("RuleOnlyChild");
        paragraph.NativeLineSpacingRule = WordLineSpacingRule.AtLeast;

        RtfDocument authored = word.ToRtfDocument();
        RtfStyle authoredChild = Assert.Single(authored.Styles, item => item.Name == "RuleOnlyChild");
        Assert.Null(authoredChild.LineSpacingTwips);
        Assert.Null(authoredChild.LineSpacingMultiple);
        Assert.Null(Assert.Single(authored.Paragraphs).LineSpacingTwips);
        Assert.Null(Assert.Single(authored.Paragraphs).LineSpacingMultiple);

        // The default writer materializes effective style formatting in body
        // controls; it must retain the inherited exact value and twip units.
        RtfDocument mapped = RtfDocument.Read(authored.ToRtf(new RtfWriteOptions { IncludeGenerator = false })).Document;
        RtfStyle mappedChild = Assert.Single(mapped.Styles, item => item.Name == "RuleOnlyChild");
        Assert.Null(mappedChild.LineSpacingTwips);
        Assert.Null(mappedChild.LineSpacingMultiple);
        Assert.Equal(-360, Assert.Single(mapped.Paragraphs).LineSpacingTwips);
        Assert.False(Assert.Single(mapped.Paragraphs).LineSpacingMultiple);
        using WordDocument returned = mapped.ToWordDocument();
        using WordDocument reopened = WordDocument.Load(new MemoryStream(returned.ToBytes()));
        AssertSpacing(Assert.Single(reopened._document.Body!.Elements<Paragraph>()), 360, WordLineSpacingRule.Exact);
    }

    [Fact]
    public void Rtf_To_Word_Rejects_Unrepresentable_Exact_Spacing_Magnitude() {
        RtfDocument source = RtfDocument.Create();
        source.AddParagraph("Spacing").SetLineSpacing(int.MinValue, false);

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => source.ToWordDocument());

        Assert.Contains("line-spacing magnitude", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Word_Rtf_Signed_Spacing_Reaches_Tables_Running_Stories_And_Both_Note_Kinds() {
        RtfDocument source = RtfDocument.Create();
        source.AddParagraph("Body").SetLineSpacing(360, false);
        RtfTableCell cell = source.AddTable(1, 1).Rows[0].Cells[0];
        cell.AddParagraph("Cell").SetLineSpacing(-360, false);
        cell.AddTable(1, 1).Rows[0].Cells[0].AddParagraph("Nested").SetLineSpacing(480, true);
        source.AddHeader().AddParagraph("Header").SetLineSpacing(360, false);
        source.AddFooter().AddParagraph("Footer").SetLineSpacing(-360, false);
        var footnote = new RtfNote(RtfNoteKind.Footnote);
        footnote.AddParagraph("Footnote").SetLineSpacing(360, false);
        var endnote = new RtfNote(RtfNoteKind.Endnote);
        endnote.AddParagraph("Endnote").SetLineSpacing(-360, false);
        source.AddParagraph("Footnote anchor").AddText(" F").SetNote(footnote);
        source.AddParagraph("Endnote anchor").AddText(" E").SetNote(endnote);

        using WordDocument word = source.ToWordDocument();
        using WordDocument reopened = WordDocument.Load(new MemoryStream(word.ToBytes()));
        var main = reopened._wordprocessingDocument.MainDocumentPart!;
        Document mainDocument = Assert.IsType<Document>(main.Document);
        AssertSpacing(Assert.Single(mainDocument.Descendants<Paragraph>(), item => item.InnerText == "Body"), 360, WordLineSpacingRule.AtLeast);
        AssertSpacing(Assert.Single(mainDocument.Descendants<Paragraph>(), item => item.InnerText == "Cell"), 360, WordLineSpacingRule.Exact);
        AssertSpacing(Assert.Single(mainDocument.Descendants<Paragraph>(), item => item.InnerText == "Nested"), 480, WordLineSpacingRule.Auto);
        AssertSpacing(Assert.Single(main.HeaderParts.SelectMany(part => Assert.IsType<Header>(part.Header).Descendants<Paragraph>()), item => item.InnerText == "Header"), 360, WordLineSpacingRule.AtLeast);
        AssertSpacing(Assert.Single(main.FooterParts.SelectMany(part => Assert.IsType<Footer>(part.Footer).Descendants<Paragraph>()), item => item.InnerText == "Footer"), 360, WordLineSpacingRule.Exact);
        AssertSpacing(Assert.Single(main.FootnotesPart!.Footnotes!.Descendants<Paragraph>(), item => item.InnerText == "Footnote"), 360, WordLineSpacingRule.AtLeast);
        AssertSpacing(Assert.Single(main.EndnotesPart!.Endnotes!.Descendants<Paragraph>(), item => item.InnerText == "Endnote"), 360, WordLineSpacingRule.Exact);

        RtfDocument returned = RtfDocument.Read(reopened.ToRtf(new RtfWriteOptions { IncludeGenerator = false })).Document;
        RtfTable returnedTable = Assert.Single(returned.Blocks.OfType<RtfTable>());
        AssertRawSpacing(Assert.Single(returned.Paragraphs, item => item.ToPlainText() == "Body"), 360, false);
        AssertRawSpacing(Assert.Single(returnedTable.Rows[0].Cells[0].Paragraphs, item => item.ToPlainText() == "Cell"), -360, false);
        RtfTable nested = Assert.Single(returnedTable.Rows[0].Cells[0].Blocks.OfType<RtfTable>());
        AssertRawSpacing(Assert.Single(nested.Rows[0].Cells[0].Paragraphs, item => item.ToPlainText() == "Nested"), 480, true);
        AssertRawSpacing(Assert.Single(returned.HeaderFooters, item => item.Kind == RtfHeaderFooterKind.Header).Paragraphs.Single(), 360, false);
        AssertRawSpacing(Assert.Single(returned.HeaderFooters, item => item.Kind == RtfHeaderFooterKind.Footer).Paragraphs.Single(), -360, false);
        RtfRun returnedFootnote = Assert.Single(Assert.Single(returned.Paragraphs, item => item.ToPlainText() == "Footnote anchor F").Runs, item => item.Note != null);
        RtfRun returnedEndnote = Assert.Single(Assert.Single(returned.Paragraphs, item => item.ToPlainText() == "Endnote anchor E").Runs, item => item.Note != null);
        AssertRawSpacing(Assert.Single(returnedFootnote.Note!.Paragraphs), 360, false);
        AssertRawSpacing(Assert.Single(returnedEndnote.Note!.Paragraphs), -360, false);
    }

    private static void AssertRawSpacing(RtfParagraph paragraph, int line, bool multiple) {
        Assert.Equal(line, paragraph.LineSpacingTwips);
        Assert.Equal(multiple, paragraph.LineSpacingMultiple);
    }

    private static void AssertSpacing(Paragraph paragraph, int line, WordLineSpacingRule rule) =>
        AssertSpacing(Assert.IsType<SpacingBetweenLines>(paragraph.ParagraphProperties?.SpacingBetweenLines), line, rule);

    private static void AssertSpacing(SpacingBetweenLines spacing, int line, WordLineSpacingRule rule) {
        Assert.Equal(line.ToString(System.Globalization.CultureInfo.InvariantCulture), spacing.Line?.Value);
        LineSpacingRuleValues expectedRule = rule == WordLineSpacingRule.Auto ? LineSpacingRuleValues.Auto
            : rule == WordLineSpacingRule.Exact ? LineSpacingRuleValues.Exact : LineSpacingRuleValues.AtLeast;
        Assert.Equal(expectedRule, spacing.LineRule?.Value);
    }
}
