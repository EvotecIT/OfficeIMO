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
    public void RtfDocumentToMarkdownKeepsOneRowHeaderlessTableParseable() {
        RtfDocument document = RtfDocument.Create();
        RtfTable table = document.AddTable(1, 1);
        table.Rows[0].Cells[0].AddParagraph("Only row");

        string markdown = document.ToMarkdown();
        RtfDocument roundTrip = markdown.ToRtfDocumentFromMarkdown();

        Assert.Equal("| Only row |", markdown.Trim());
        RtfTable roundTripTable = Assert.IsType<RtfTable>(roundTrip.Blocks.OfType<RtfTable>().Single());
        Assert.Single(roundTripTable.Rows);
        Assert.Equal("Only row", roundTripTable.Rows[0].Cells[0].Paragraphs[0].ToPlainText());
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
}
