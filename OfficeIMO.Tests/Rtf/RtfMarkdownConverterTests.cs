using OfficeIMO.Markdown;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Markdown;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfMarkdownConverterTests {
    [Fact]
    public void RtfDocumentToMarkdownPreservesCoreInlineBlocksListsAndTables() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.AddText("Hello ");
        paragraph.AddText("bold").SetBold();
        paragraph.AddText(" and ");
        paragraph.AddText("italic").SetItalic();
        paragraph.AddText(" with ");
        paragraph.AddText("strike").SetStrike();
        paragraph.AddText(" plus ");
        paragraph.AddText("link").SetHyperlink(new Uri("https://evotec.xyz"));

        document.AddParagraph("One").SetList(kind: RtfListKind.Bullet);
        document.AddParagraph("Two").SetList(kind: RtfListKind.Bullet);

        RtfTable table = document.AddTable(2, 2);
        table.Rows[0].Cells[0].AddParagraph("Name");
        table.Rows[0].Cells[1].AddParagraph("Value");
        table.Rows[1].Cells[0].AddParagraph("Alpha");
        table.Rows[1].Cells[1].AddParagraph("Beta");

        string markdown = document.ToMarkdown();

        Assert.Contains("**bold**", markdown);
        Assert.Contains("*italic*", markdown);
        Assert.Contains("~~strike~~", markdown);
        Assert.Contains("[link](https://evotec.xyz/)", markdown);
        Assert.Contains("- One", markdown);
        Assert.Contains("| Name | Value |", markdown);
        Assert.Contains("| Alpha | Beta |", markdown);
    }

    [Fact]
    public void MarkdownToRtfDocumentPreservesHeadingsInlinesListsAndTables() {
        string markdown = """
            # Title

            Hello **bold** and _italic_ with [link](https://evotec.xyz).

            - One
            - Two

            | Name | Value |
            | --- | --- |
            | Alpha | Beta |
            """;

        RtfDocument document = markdown.ToRtfDocumentFromMarkdown();

        Assert.Equal("Title", document.Paragraphs[0].ToPlainText());
        Assert.Equal(0, document.Paragraphs[0].OutlineLevel);
        Assert.Contains(document.Paragraphs[1].Runs, run => run.Text == "bold" && run.Bold);
        Assert.Contains(document.Paragraphs[1].Runs, run => run.Text == "italic" && run.Italic);
        Assert.Contains(document.Paragraphs[1].Runs, run => run.Text == "link" && run.Hyperlink != null);
        Assert.Contains(document.Paragraphs, paragraph => paragraph.ListKind == RtfListKind.Bullet && paragraph.ToPlainText() == "One");
        Assert.Contains(document.Blocks, block => block is RtfTable);
    }

    [Fact]
    public void MarkdownRtfMarkdownRoundTripKeepsReadableSemanticMarkdown() {
        string markdown = """
            ## Overview

            This is **important** and includes `code`.
            """;

        RtfDocument document = markdown.ToRtfDocumentFromMarkdown();
        string roundTripMarkdown = document.ToMarkdown();

        Assert.Contains("## Overview", roundTripMarkdown);
        Assert.Contains("**important**", roundTripMarkdown);
        Assert.Contains("code", roundTripMarkdown);
    }

    [Fact]
    public void MarkdownImagesEmitDiagnosticWhenBinaryPayloadIsNotProvided() {
        var options = new MarkdownToRtfOptions();

        RtfDocument document = "![Logo](logo.png)".ToRtfDocumentFromMarkdown(options);

        Assert.Contains(document.Paragraphs, paragraph => paragraph.ToPlainText().Contains("[Image: Logo]", StringComparison.Ordinal));
        Assert.Contains(options.Diagnostics, diagnostic => diagnostic.Code == "MDRTF003");
    }

    [Fact]
    public void RtfDocumentToMarkdownEscapesLiteralHtmlAndMarkdownText() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("<u>literal</u>");
        RtfTable table = document.AddTable(2, 1);
        table.Rows[0].RepeatHeader = true;
        table.Rows[0].Cells[0].AddParagraph("Header");
        table.Rows[1].Cells[0].AddParagraph("**not bold**");

        string markdown = document.ToMarkdown();
        RtfDocument roundTrip = markdown.ToRtfDocumentFromMarkdown();

        Assert.Contains("&lt;u&gt;literal&lt;/u&gt;", markdown, StringComparison.Ordinal);
        Assert.Contains(@"\*\*not bold\*\*", markdown, StringComparison.Ordinal);
        Assert.Contains(roundTrip.Paragraphs, paragraph => paragraph.ToPlainText() == "<u>literal</u>" && paragraph.Runs.All(run => run.UnderlineStyle == RtfUnderlineStyle.None));
        RtfTable roundTripTable = Assert.IsType<RtfTable>(roundTrip.Blocks.OfType<RtfTable>().Single());
        Assert.Equal("**not bold**", roundTripTable.Rows[1].Cells[0].Paragraphs[0].ToPlainText());
        Assert.DoesNotContain(roundTripTable.Rows[1].Cells[0].Paragraphs[0].Runs, run => run.Bold);
    }

    [Fact]
    public void RtfDocumentToMarkdownPreservesCombinedFormattingAndLinks() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.AddText("link").SetBold().SetHyperlink(new Uri("https://evotec.xyz"));
        paragraph.AddText(" ");
        paragraph.AddText("underlined").SetBold().SetUnderline(RtfUnderlineStyle.Single);

        string markdown = document.ToMarkdown();
        RtfDocument roundTrip = markdown.ToRtfDocumentFromMarkdown();

        Assert.Contains("[**link**](https://evotec.xyz/)", markdown, StringComparison.Ordinal);
        Assert.Contains("**<u>underlined</u>**", markdown, StringComparison.Ordinal);
        Assert.Contains(roundTrip.Paragraphs[0].Runs, run => run.Text == "link" && run.Bold && run.Hyperlink != null);
        Assert.Contains(roundTrip.Paragraphs[0].Runs, run => run.Text == "underlined" && run.Bold && run.UnderlineStyle != RtfUnderlineStyle.None);
    }

    [Fact]
    public void RtfDocumentToMarkdownPreservesUnderlineWithVerticalPosition() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph().AddText("raised").SetUnderline(RtfUnderlineStyle.Single).VerticalPosition = RtfVerticalPosition.Superscript;

        string markdown = document.ToMarkdown();
        RtfDocument roundTrip = markdown.ToRtfDocumentFromMarkdown();

        Assert.Contains("<u><sup>raised</sup></u>", markdown, StringComparison.Ordinal);
        Assert.Contains(roundTrip.Paragraphs[0].Runs, run =>
            run.Text == "raised" &&
            run.UnderlineStyle != RtfUnderlineStyle.None &&
            run.VerticalPosition == RtfVerticalPosition.Superscript);
    }

    [Fact]
    public void MarkdownToRtfDocumentPreservesHtmlInlineFormattingTags() {
        RtfDocument document = "<u>under</u> <sup>up</sup> <sub>down</sub>".ToRtfDocumentFromMarkdown();

        RtfParagraph paragraph = Assert.Single(document.Paragraphs);
        Assert.Contains(paragraph.Runs, run => run.Text == "under" && run.UnderlineStyle != RtfUnderlineStyle.None);
        Assert.Contains(paragraph.Runs, run => run.Text == "up" && run.VerticalPosition == RtfVerticalPosition.Superscript);
        Assert.Contains(paragraph.Runs, run => run.Text == "down" && run.VerticalPosition == RtfVerticalPosition.Subscript);
    }

    [Fact]
    public void MarkdownToRtfDocumentOmitsHtmlCommentBlocksByDefault() {
        var options = new MarkdownToRtfOptions();

        RtfDocument document = """
            <!-- hidden -->

            Visible
            """.ToRtfDocumentFromMarkdown(options);

        RtfParagraph paragraph = Assert.Single(document.Paragraphs);
        Assert.Equal("Visible", paragraph.ToPlainText());
        Assert.Contains(options.Diagnostics, diagnostic => diagnostic.Code == "MDRTF004");
    }

    [Fact]
    public void MarkdownToRtfDocumentKeepsEntitiesLiteralInsideCodeSpans() {
        RtfDocument document = "`&lt;tag&gt;` &lt;tag&gt;".ToRtfDocumentFromMarkdown();

        RtfParagraph paragraph = Assert.Single(document.Paragraphs);

        Assert.Contains(paragraph.Runs, run => run.Text == "&lt;tag&gt;" && run.FontId.HasValue);
        Assert.Contains(paragraph.Runs, run => run.Text == " <tag>" && !run.FontId.HasValue);
    }

    [Fact]
    public void MarkdownToRtfDocumentPreservesOrderedStartsNestedListsAndTableInlines() {
        string markdown = """
            5. Parent
               - Child

            | Name | Value |
            | --- | --- |
            | **Bold** | [Link](https://evotec.xyz) |
            """;

        RtfDocument document = markdown.ToRtfDocumentFromMarkdown();

        RtfParagraph parent = Assert.Single(document.Paragraphs, paragraph => paragraph.ToPlainText() == "Parent");
        RtfParagraph child = Assert.Single(document.Paragraphs, paragraph => paragraph.ToPlainText() == "Child");
        Assert.Equal(RtfListKind.Decimal, parent.ListKind);
        Assert.Equal(0, parent.ListLevel);
        Assert.Equal(RtfListKind.Bullet, child.ListKind);
        Assert.Equal(1, child.ListLevel);
        RtfListDefinition definition = Assert.Single(document.ListDefinitions, item => item.Id == parent.ListDefinitionId);
        Assert.Equal(5, definition.Levels[0].StartAt);

        RtfTable table = Assert.IsType<RtfTable>(document.Blocks.OfType<RtfTable>().Single());
        Assert.Contains(table.Rows[1].Cells[0].Paragraphs[0].Runs, run => run.Text == "Bold" && run.Bold);
        Assert.Contains(table.Rows[1].Cells[1].Paragraphs[0].Runs, run => run.Text == "Link" && run.Hyperlink != null);
    }

    [Fact]
    public void MarkdownToRtfDocumentKeepsNestedListsInSameListDefinition() {
        RtfDocument document = """
            3. Parent
               - Child
            """.ToRtfDocumentFromMarkdown();

        RtfParagraph parent = Assert.Single(document.Paragraphs, paragraph => paragraph.ToPlainText() == "Parent");
        RtfParagraph child = Assert.Single(document.Paragraphs, paragraph => paragraph.ToPlainText() == "Child");

        Assert.Equal(parent.ListId, child.ListId);
        Assert.Equal(parent.ListDefinitionId, child.ListDefinitionId);
        Assert.Equal(0, parent.ListLevel);
        Assert.Equal(1, child.ListLevel);
        RtfListDefinition definition = Assert.Single(document.ListDefinitions, item => item.Id == parent.ListDefinitionId);
        Assert.Equal(RtfListKind.Decimal, definition.Levels[0].Kind);
        Assert.Equal(3, definition.Levels[0].StartAt);
        Assert.Equal(RtfListKind.Bullet, definition.Levels[1].Kind);
    }

    [Fact]
    public void MarkdownToRtfDocumentAppliesNestedOrderedStarts() {
        RtfDocument document = """
            1. Parent
               5. Child
            """.ToRtfDocumentFromMarkdown();

        RtfParagraph parent = Assert.Single(document.Paragraphs, paragraph => paragraph.ToPlainText() == "Parent");
        RtfParagraph child = Assert.Single(document.Paragraphs, paragraph => paragraph.ToPlainText() == "Child");
        RtfListDefinition parentDefinition = Assert.Single(document.ListDefinitions, item => item.Id == parent.ListDefinitionId);
        RtfListDefinition childDefinition = Assert.Single(document.ListDefinitions, item => item.Id == child.ListDefinitionId);
        RtfListOverride childOverride = Assert.Single(document.ListOverrides, item => item.Id == child.ListId);

        Assert.NotEqual(parent.ListId, child.ListId);
        Assert.NotEqual(parent.ListDefinitionId, child.ListDefinitionId);
        Assert.Equal(1, child.ListLevel);
        Assert.Equal(1, parentDefinition.Levels[0].StartAt);
        Assert.Equal(5, childDefinition.Levels[1].StartAt);
        Assert.Equal(5, childOverride.LevelOverrides[1].StartAt);
    }

    [Fact]
    public void MarkdownToRtfDocumentWritesNoOpOverridesForSkippedListLevels() {
        string rtf = """
            1. Parent
               5. Child
            """.ToRtfFromMarkdown();

        Assert.Contains(@"\listoverridecount0", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\listoverridecount2", rtf, StringComparison.Ordinal);
        Assert.Equal(2, CountOccurrences(rtf, @"{\lfolevel"));
        Assert.Contains(@"{\lfolevel\listoverridestartat1\levelstartat5}", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownToRtfDocumentAppliesTableColumnAlignmentsToCellParagraphs() {
        RtfDocument document = """
            | Name | Count | Status |
            | :--- | ---: | :---: |
            | Alpha | 42 | Ready |
            """.ToRtfDocumentFromMarkdown();

        RtfTable table = Assert.IsType<RtfTable>(document.Blocks.OfType<RtfTable>().Single());

        Assert.Equal(RtfTextAlignment.Left, table.Rows[1].Cells[0].Paragraphs[0].Alignment);
        Assert.Equal(RtfTextAlignment.Right, table.Rows[1].Cells[1].Paragraphs[0].Alignment);
        Assert.Equal(RtfTextAlignment.Center, table.Rows[1].Cells[2].Paragraphs[0].Alignment);
    }

    [Fact]
    public void RtfDocumentToMarkdownPreservesTableColumnAlignments() {
        RtfDocument document = RtfDocument.Create();
        RtfTable table = document.AddTable(2, 2);
        table.Rows[0].RepeatHeader = true;
        table.Rows[0].Cells[0].AddParagraph("Name").SetAlignment(RtfTextAlignment.Center);
        table.Rows[0].Cells[1].AddParagraph("Count").SetAlignment(RtfTextAlignment.Right);
        table.Rows[1].Cells[0].AddParagraph("Alpha").SetAlignment(RtfTextAlignment.Center);
        table.Rows[1].Cells[1].AddParagraph("42").SetAlignment(RtfTextAlignment.Right);

        string markdown = document.ToMarkdown();
        RtfDocument roundTrip = markdown.ToRtfDocumentFromMarkdown();
        RtfTable roundTripTable = Assert.IsType<RtfTable>(roundTrip.Blocks.OfType<RtfTable>().Single());

        Assert.Contains("| :---: | ---: |", markdown, StringComparison.Ordinal);
        Assert.Equal(RtfTextAlignment.Center, roundTripTable.Rows[1].Cells[0].Paragraphs[0].Alignment);
        Assert.Equal(RtfTextAlignment.Right, roundTripTable.Rows[1].Cells[1].Paragraphs[0].Alignment);
    }

    [Fact]
    public void RtfDocumentToMarkdownPreservesTaskListItems() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("[x] Done").SetList(kind: RtfListKind.Bullet);
        document.AddParagraph("[ ] Todo").SetList(kind: RtfListKind.Bullet);

        string markdown = document.ToMarkdown();
        MarkdownDoc parsed = MarkdownReader.Parse(markdown);
        UnorderedListBlock list = Assert.IsType<UnorderedListBlock>(Assert.Single(parsed.Blocks));

        Assert.Contains("- [x] Done", markdown, StringComparison.Ordinal);
        Assert.Contains("- [ ] Todo", markdown, StringComparison.Ordinal);
        Assert.True(list.Items[0].IsTask);
        Assert.True(list.Items[0].Checked);
        Assert.True(list.Items[1].IsTask);
        Assert.False(list.Items[1].Checked);
        Assert.Equal("Done", ExtractPlainText(list.Items[0].Content));
        Assert.Equal("Todo", ExtractPlainText(list.Items[1].Content));
    }

    [Fact]
    public void RtfDocumentToMarkdownKeepsLiteralBlockMarkersAsText() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("# Heading\n- item\n1. item\n---");

        string markdown = document.ToMarkdown();
        RtfDocument roundTrip = markdown.ToRtfDocumentFromMarkdown();

        Assert.Contains(@"\# Heading", markdown, StringComparison.Ordinal);
        Assert.Contains(@"\- item", markdown, StringComparison.Ordinal);
        Assert.Contains(@"1\. item", markdown, StringComparison.Ordinal);
        Assert.Contains(@"\---", markdown, StringComparison.Ordinal);
        Assert.All(roundTrip.Paragraphs, paragraph => {
            Assert.Equal(RtfListKind.None, paragraph.ListKind);
            Assert.Null(paragraph.OutlineLevel);
        });
    }

    [Fact]
    public void RtfDocumentToMarkdownKeepsLiteralTildeAndHighlightMarkersParseable() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("~~not strike~~ ==not mark==");

        string markdown = document.ToMarkdown();
        RtfDocument roundTrip = markdown.ToRtfDocumentFromMarkdown();

        Assert.Contains(@"\~\~not strike\~\~", markdown, StringComparison.Ordinal);
        Assert.Contains(@"\=\=not mark\=\=", markdown, StringComparison.Ordinal);
        Assert.Equal("~~not strike~~ ==not mark==", roundTrip.Paragraphs[0].ToPlainText());
        Assert.DoesNotContain(roundTrip.Paragraphs[0].Runs, run => run.Strike || run.DoubleStrike || run.HighlightColorIndex.HasValue);
    }

    [Fact]
    public void RtfDocumentToMarkdownEncodesLiteralEntityAmpersands() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("&lt; &#42;");

        string markdown = document.ToMarkdown();
        RtfDocument roundTrip = markdown.ToRtfDocumentFromMarkdown();

        Assert.Contains("&amp;lt; &amp;#42;", markdown, StringComparison.Ordinal);
        Assert.Equal("&lt; &#42;", roundTrip.Paragraphs[0].ToPlainText());
    }

    [Fact]
    public void RtfDocumentToMarkdownRoundTripPreservesTableCellHtmlFormatting() {
        RtfDocument document = RtfDocument.Create();
        RtfTable table = document.AddTable(2, 3);
        table.Rows[0].RepeatHeader = true;
        table.Rows[0].Cells[0].AddParagraph("Under");
        table.Rows[0].Cells[1].AddParagraph("Super");
        table.Rows[0].Cells[2].AddParagraph("Sub");
        table.Rows[1].Cells[0].AddParagraph().AddText("under").SetUnderline(RtfUnderlineStyle.Single);
        table.Rows[1].Cells[1].AddParagraph().AddText("up").VerticalPosition = RtfVerticalPosition.Superscript;
        table.Rows[1].Cells[2].AddParagraph().AddText("down").VerticalPosition = RtfVerticalPosition.Subscript;

        string markdown = document.ToMarkdown();
        RtfDocument roundTrip = markdown.ToRtfDocumentFromMarkdown();
        RtfTable roundTripTable = Assert.IsType<RtfTable>(roundTrip.Blocks.OfType<RtfTable>().Single());

        Assert.Contains("<u>under</u>", markdown, StringComparison.Ordinal);
        Assert.Contains("<sup>up</sup>", markdown, StringComparison.Ordinal);
        Assert.Contains("<sub>down</sub>", markdown, StringComparison.Ordinal);
        Assert.Contains(roundTripTable.Rows[1].Cells[0].Paragraphs[0].Runs, run => run.Text == "under" && run.UnderlineStyle != RtfUnderlineStyle.None);
        Assert.Contains(roundTripTable.Rows[1].Cells[1].Paragraphs[0].Runs, run => run.Text == "up" && run.VerticalPosition == RtfVerticalPosition.Superscript);
        Assert.Contains(roundTripTable.Rows[1].Cells[2].Paragraphs[0].Runs, run => run.Text == "down" && run.VerticalPosition == RtfVerticalPosition.Subscript);
    }

    [Fact]
    public void RtfDocumentToMarkdownRoundTripPreservesLiteralEntityTextInTableCells() {
        RtfDocument document = RtfDocument.Create();
        RtfTable table = document.AddTable(2, 1);
        table.Rows[0].RepeatHeader = true;
        table.Rows[0].Cells[0].AddParagraph("Value");
        table.Rows[1].Cells[0].AddParagraph("&lt; &#42;");

        string markdown = document.ToMarkdown();
        RtfDocument roundTrip = markdown.ToRtfDocumentFromMarkdown();
        RtfTable roundTripTable = Assert.IsType<RtfTable>(roundTrip.Blocks.OfType<RtfTable>().Single());

        Assert.Contains("| &lt; &#42; |", markdown, StringComparison.Ordinal);
        Assert.Equal("&lt; &#42;", roundTripTable.Rows[1].Cells[0].Paragraphs[0].ToPlainText());
    }

    [Fact]
    public void RtfDocumentToMarkdownRoundTripPreservesDefinitionLikeTableCellText() {
        RtfDocument document = RtfDocument.Create();
        RtfTable table = document.AddTable(2, 1);
        table.Rows[0].RepeatHeader = true;
        table.Rows[0].Cells[0].AddParagraph("Value");
        table.Rows[1].Cells[0].AddParagraph("Key: value");

        string markdown = document.ToMarkdown();
        RtfDocument roundTrip = markdown.ToRtfDocumentFromMarkdown();
        RtfTable roundTripTable = Assert.IsType<RtfTable>(roundTrip.Blocks.OfType<RtfTable>().Single());

        Assert.Contains("| Key: value |", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain(@"Key\: value", markdown, StringComparison.Ordinal);
        Assert.Equal("Key: value", roundTripTable.Rows[1].Cells[0].Paragraphs[0].ToPlainText());
    }

    [Fact]
    public void MarkdownToRtfDocumentPreservesDefinitionListTermsAndInlineFormatting() {
        RtfDocument document = "Term: **Definition**".ToRtfDocumentFromMarkdown();

        RtfParagraph paragraph = Assert.Single(document.Paragraphs);
        Assert.Equal("Term: Definition", paragraph.ToPlainText());
        Assert.Contains(paragraph.Runs, run => run.Text == "Definition" && run.Bold);
    }

    [Fact]
    public void MarkdownToRtfDocumentClearsStaleNestedOrderedStarts() {
        string markdown = """
            1. A
               5. first
            2. B
               1. second
            """;

        RtfDocument serialized = RtfDocument.Read(markdown.ToRtfFromMarkdown()).Document;
        string roundTripMarkdown = serialized.ToMarkdown().Replace("\r\n", "\n");

        Assert.Contains("5. first", roundTripMarkdown, StringComparison.Ordinal);
        Assert.Contains("1. second", roundTripMarkdown, StringComparison.Ordinal);
        Assert.DoesNotContain("5. second", roundTripMarkdown, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownToRtfDocumentPreservesAdditionalListParagraphsTableHeadersAndEscapedEntities() {
        var list = new UnorderedListBlock();
        var item = ListItem.Text("Lead");
        item.AdditionalParagraphs.Add(new InlineSequence().Text("Continuation"));
        list.Items.Add(item);
        MarkdownDoc markdownDoc = MarkdownDoc.Create().Add(list);

        RtfDocument listDocument = markdownDoc.ToRtfDocument();

        Assert.Contains(listDocument.Paragraphs, paragraph => paragraph.ToPlainText() == "Lead" && paragraph.ListKind == RtfListKind.Bullet);
        Assert.Contains(listDocument.Paragraphs, paragraph => paragraph.ToPlainText() == "Continuation" && paragraph.ListKind == RtfListKind.None);

        RtfDocument tableDocument = """
            | Name | Value |
            | --- | --- |
            | Alpha | Beta |
            """.ToRtfDocumentFromMarkdown();
        RtfTable table = Assert.IsType<RtfTable>(tableDocument.Blocks.OfType<RtfTable>().Single());
        Assert.True(table.Rows[0].RepeatHeader);
        string tableRoundTripMarkdown = tableDocument.ToMarkdown();
        Assert.Contains("| Name | Value |", tableRoundTripMarkdown, StringComparison.Ordinal);
        Assert.Contains("| --- | --- |", tableRoundTripMarkdown, StringComparison.Ordinal);

        RtfDocument escaped = @"&amp;lt; &amp;#42;".ToRtfDocumentFromMarkdown();
        Assert.Equal("&lt; &#42;", escaped.Paragraphs[0].ToPlainText());
    }

    [Fact]
    public void RtfDocumentToMarkdownKeepsOneRowTableParseable() {
        RtfDocument document = RtfDocument.Create();
        RtfTable table = document.AddTable(1, 1);
        table.Rows[0].Cells[0].AddParagraph("Only row");

        string markdown = document.ToMarkdown();
        RtfDocument roundTrip = markdown.ToRtfDocumentFromMarkdown();

        Assert.Equal("""
            <!-- OfficeIMO:RTF:HeaderlessSingleRowTable -->
            | Only row |
            """.Replace("\r\n", "\n").Trim(), markdown.Replace("\r\n", "\n").Trim());
        RtfTable roundTripTable = Assert.IsType<RtfTable>(roundTrip.Blocks.OfType<RtfTable>().Single());
        Assert.Single(roundTripTable.Rows);
        Assert.False(roundTripTable.Rows[0].RepeatHeader);
        Assert.Equal("Only row", roundTripTable.Rows[0].Cells[0].Paragraphs[0].ToPlainText());
    }

    [Fact]
    public void MarkdownToRtfDocumentReadsOneRowTableMarkerWhenHtmlBlocksAreDisabled() {
        var options = new MarkdownToRtfOptions {
            ReaderOptions = new MarkdownReaderOptions {
                HtmlBlocks = false
            }
        };

        RtfDocument document = """
            <!-- OfficeIMO:RTF:HeaderlessSingleRowTable -->
            | Only row |
            """.ToRtfDocumentFromMarkdown(options);

        RtfTable table = Assert.IsType<RtfTable>(document.Blocks.OfType<RtfTable>().Single());
        Assert.Single(table.Rows);
        Assert.False(table.Rows[0].RepeatHeader);
        Assert.Equal("Only row", table.Rows[0].Cells[0].Paragraphs[0].ToPlainText());
        Assert.DoesNotContain(document.Paragraphs, paragraph => paragraph.ToPlainText().Contains("OfficeIMO:RTF", StringComparison.Ordinal));
    }

    [Fact]
    public void RtfDocumentToMarkdownIgnoresDisabledListStartOverrides() {
        RtfDocument document = RtfDocument.Create();
        RtfListDefinition definition = document.AddListDefinition(100);
        definition.AddLevel(RtfListKind.Decimal).StartAt = 3;
        RtfListOverride listOverride = document.AddListOverride(7, 100);
        RtfListLevelOverride levelOverride = listOverride.AddLevelOverride();
        levelOverride.OverrideStartAt = false;
        levelOverride.StartAt = 9;
        document.AddParagraph("Three").SetList(listId: 7, level: 0, kind: RtfListKind.Decimal).ListDefinitionId = 100;

        string markdown = document.ToMarkdown();

        Assert.Contains("3. Three", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("9. Three", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void RtfDocumentToMarkdownDoesNotAbsorbOrdinaryIndentedParagraphAfterList() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Item").SetList(kind: RtfListKind.Bullet);
        document.AddParagraph("Indented standalone").SetIndentation(leftTwips: 720);

        string markdown = document.ToMarkdown();
        MarkdownDoc parsed = MarkdownReader.Parse(markdown);

        Assert.Collection(parsed.Blocks,
            block => {
                UnorderedListBlock list = Assert.IsType<UnorderedListBlock>(block);
                Assert.Single(list.Items);
                Assert.Empty(list.Items[0].AdditionalParagraphs);
            },
            block => {
                ParagraphBlock paragraph = Assert.IsType<ParagraphBlock>(block);
                Assert.Equal("Indented standalone", ExtractPlainText(paragraph.Inlines));
            });
    }

    [Fact]
    public void RtfMarkdownRoundTripKeepsNestedListChildBlocksInsideListItem() {
        string markdown = """
            - Lead

              > Quoted

              ```
              code
              ```

              | A | B |
              | --- | --- |
              | C | D |

            - Next
            """;

        RtfDocument document = markdown.ToRtfDocumentFromMarkdown();
        RtfDocument serialized = RtfDocument.Read(document.ToRtf()).Document;
        string roundTripMarkdown = serialized.ToMarkdown().Replace("\r\n", "\n");
        MarkdownDoc parsed = MarkdownReader.Parse(roundTripMarkdown);
        UnorderedListBlock list = Assert.IsType<UnorderedListBlock>(Assert.Single(parsed.Blocks));

        Assert.Equal(2, list.Items.Count);
        Assert.Equal("Lead", ExtractPlainText(list.Items[0].Content));
        Assert.Equal("Next", ExtractPlainText(list.Items[1].Content));
        Assert.Contains(list.Items[0].AdditionalParagraphs, paragraph => ExtractPlainText(paragraph).Contains("Quoted", StringComparison.Ordinal));
        Assert.Contains(list.Items[0].AdditionalParagraphs, paragraph => ExtractPlainText(paragraph).Contains("code", StringComparison.Ordinal));
        Assert.Contains(list.Items[0].AdditionalParagraphs, paragraph => ExtractPlainText(paragraph).Contains("| A | B |", StringComparison.Ordinal));
        Assert.DoesNotContain(parsed.Blocks.Skip(1), block => block is QuoteBlock || block is TableBlock || block is CodeBlock);
    }

    [Fact]
    public void RtfDocumentToMarkdownReportsRunAttachedNotes() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph().AddFootnote("1", "Footnote body");
        var options = new RtfToMarkdownOptions();

        string markdown = document.ToMarkdown(options);

        Assert.Contains("1", markdown, StringComparison.Ordinal);
        Assert.Contains(options.Diagnostics, diagnostic => diagnostic.Code == "RTFMD012");
    }

    [Fact]
    public void RtfDocumentToMarkdownReportsGeneratedTextWithoutFallback() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph().AddPageNumber();
        var options = new RtfToMarkdownOptions();

        string markdown = document.ToMarkdown(options);

        Assert.Contains("RTF generated text omitted", markdown, StringComparison.Ordinal);
        Assert.Contains(options.Diagnostics, diagnostic => diagnostic.Code == "RTFMD013");
    }

    [Fact]
    public void RtfDocumentToMarkdownKeepsMixedKindNestedListTogether() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Parent").SetList(listId: 7, level: 0, kind: RtfListKind.Decimal);
        document.AddParagraph("Child").SetList(listId: 7, level: 1, kind: RtfListKind.Bullet);
        document.AddParagraph("Next").SetList(listId: 7, level: 0, kind: RtfListKind.Decimal);

        string markdown = document.ToMarkdown().Replace("\r\n", "\n");
        RtfDocument roundTrip = markdown.ToRtfDocumentFromMarkdown();

        Assert.Contains("1. Parent\n\n   - Child\n\n2. Next", markdown, StringComparison.Ordinal);
        Assert.Contains(roundTrip.Paragraphs, paragraph => paragraph.ToPlainText() == "Parent" && paragraph.ListKind == RtfListKind.Decimal && paragraph.ListLevel == 0);
        Assert.Contains(roundTrip.Paragraphs, paragraph => paragraph.ToPlainText() == "Child" && paragraph.ListKind == RtfListKind.Bullet && paragraph.ListLevel == 1);
        Assert.Contains(roundTrip.Paragraphs, paragraph => paragraph.ToPlainText() == "Next" && paragraph.ListKind == RtfListKind.Decimal && paragraph.ListLevel == 0);
    }

    [Fact]
    public void RtfDocumentToMarkdownPreservesListStartsAndDoesNotPromoteDataRows() {
        RtfDocument document = RtfDocument.Create();
        RtfListDefinition definition = document.AddListDefinition(100);
        definition.AddLevel(RtfListKind.Decimal).StartAt = 5;
        document.AddListOverride(3, 100);
        document.AddParagraph("Five").SetList(listId: 3, level: 0, kind: RtfListKind.Decimal).ListDefinitionId = 100;
        document.AddParagraph("Six").SetList(listId: 3, level: 0, kind: RtfListKind.Decimal).ListDefinitionId = 100;
        document.AddParagraph("Restart").SetList(listId: 4, level: 0, kind: RtfListKind.Decimal);

        RtfTable table = document.AddTable(2, 1);
        table.Rows[0].Cells[0].AddParagraph("Data one");
        table.Rows[1].Cells[0].AddParagraph("Data two");

        string markdown = document.ToMarkdown();

        Assert.Contains("5. Five", markdown, StringComparison.Ordinal);
        Assert.Contains("6. Six", markdown, StringComparison.Ordinal);
        Assert.Contains("1. Restart", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("| --- |", markdown, StringComparison.Ordinal);
        Assert.Contains("| Data one |", markdown, StringComparison.Ordinal);
        Assert.Contains("| Data two |", markdown, StringComparison.Ordinal);
    }

    private static string ExtractPlainText(IPlainTextMarkdownInline inline) {
        var builder = new System.Text.StringBuilder();
        inline.AppendPlainText(builder);
        return builder.ToString();
    }

    private static int CountOccurrences(string value, string needle) {
        int count = 0;
        int index = 0;
        while ((index = value.IndexOf(needle, index, StringComparison.Ordinal)) >= 0) {
            count++;
            index += needle.Length;
        }

        return count;
    }
}
