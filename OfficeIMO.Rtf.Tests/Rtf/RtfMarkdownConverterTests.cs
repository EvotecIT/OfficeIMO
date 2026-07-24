using OfficeIMO.Markdown;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Markdown;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfMarkdownConverterTests {
    [Fact]
    public async Task BidirectionalBridgeExposesBytesStreamsAndRealIoAsyncSaves() {
        RtfDocument rtf = RtfDocument.Create();
        rtf.AddParagraph("RTF to Markdown");

        Assert.Equal("RTF to Markdown", Encoding.UTF8.GetString(rtf.ToMarkdownBytes()).Trim());
        using MemoryStream markdownStream = rtf.ToMarkdownStream();
        Assert.Equal(0, markdownStream.Position);
        using var savedMarkdown = new MemoryStream();
        await rtf.SaveAsMarkdownAsync(savedMarkdown);
        Assert.Contains("RTF to Markdown", Encoding.UTF8.GetString(savedMarkdown.ToArray()), StringComparison.Ordinal);

        MarkdownDoc markdown = MarkdownDoc.Create().P("Markdown to RTF");
        byte[] rtfBytes = markdown.ToRtfBytes();
        Assert.StartsWith("{\\rtf", Encoding.UTF8.GetString(rtfBytes), StringComparison.Ordinal);
        using MemoryStream rtfStream = markdown.ToRtfStream();
        Assert.Equal(0, rtfStream.Position);
        using var savedRtf = new MemoryStream();
        await markdown.SaveAsRtfAsync(savedRtf);
        Assert.Equal(rtfBytes, savedRtf.ToArray());

        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            markdown.SaveAsRtfAsync(new MemoryStream(), cancellationToken: cancellation.Token));
    }

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

        RtfDocument document = MarkdownReader.Parse(markdown).ToRtfDocument();

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

        RtfDocument document = MarkdownReader.Parse(markdown).ToRtfDocument();
        string roundTripMarkdown = document.ToMarkdown();

        Assert.Contains("## Overview", roundTripMarkdown);
        Assert.Contains("**important**", roundTripMarkdown);
        Assert.Contains("code", roundTripMarkdown);
    }

    [Fact]
    public void MarkdownImagesEmitDiagnosticWhenBinaryPayloadIsNotProvided() {
        var options = new MarkdownToRtfOptions();

        RtfConversionResult<RtfDocument> result = MarkdownReader.Parse("![Logo](logo.png)").ToRtfDocumentResult(options);
        RtfDocument document = result.Value;

        Assert.Contains(document.Paragraphs, paragraph => paragraph.ToPlainText().Contains("[Image: Logo]", StringComparison.Ordinal));
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "MDRTF003");
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
        RtfDocument roundTrip = MarkdownReader.Parse(markdown).ToRtfDocument();

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
        RtfDocument roundTrip = MarkdownReader.Parse(markdown).ToRtfDocument();

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
        RtfDocument roundTrip = MarkdownReader.Parse(markdown).ToRtfDocument();

        Assert.Contains("<u><sup>raised</sup></u>", markdown, StringComparison.Ordinal);
        Assert.Contains(roundTrip.Paragraphs[0].Runs, run =>
            run.Text == "raised" &&
            run.UnderlineStyle != RtfUnderlineStyle.None &&
            run.VerticalPosition == RtfVerticalPosition.Superscript);
    }

    [Fact]
    public void MarkdownToRtfDocumentPreservesHtmlInlineFormattingTags() {
        RtfDocument document = MarkdownReader.Parse("<u>under</u> <sup>up</sup> <sub>down</sub>").ToRtfDocument();

        RtfParagraph paragraph = Assert.Single(document.Paragraphs);
        Assert.Contains(paragraph.Runs, run => run.Text == "under" && run.UnderlineStyle != RtfUnderlineStyle.None);
        Assert.Contains(paragraph.Runs, run => run.Text == "up" && run.VerticalPosition == RtfVerticalPosition.Superscript);
        Assert.Contains(paragraph.Runs, run => run.Text == "down" && run.VerticalPosition == RtfVerticalPosition.Subscript);
    }

    [Fact]
    public void MarkdownToRtfDocumentRejectsUnsupportedHtmlNestedInSupportedWrapper() {
        var options = new MarkdownToRtfOptions();

        RtfConversionResult<RtfDocument> result = MarkdownReader.Parse("<u><span>x</span></u>").ToRtfDocumentResult(options);
        RtfDocument document = result.Value;

        Assert.Empty(document.Paragraphs);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "MDRTF004");
    }

    [Fact]
    public void MarkdownToRtfDocumentPreservesDecodedEntityTextInsideHtmlWrappers() {
        RtfDocument document = MarkdownReader.Parse("<u>&amp;lt;</u>").ToRtfDocument();

        RtfParagraph paragraph = Assert.Single(document.Paragraphs);
        RtfRun run = Assert.Single(paragraph.Runs);

        Assert.Equal("&lt;", run.Text);
        Assert.NotEqual(RtfUnderlineStyle.None, run.UnderlineStyle);
    }

    [Fact]
    public void MarkdownToRtfDocumentOmitsHtmlCommentBlocksByDefault() {
        var options = new MarkdownToRtfOptions();

        RtfConversionResult<RtfDocument> result = MarkdownReader.Parse("""
            <!-- hidden -->

            Visible
            """).ToRtfDocumentResult(options);
        RtfDocument document = result.Value;

        RtfParagraph paragraph = Assert.Single(document.Paragraphs);
        Assert.Equal("Visible", paragraph.ToPlainText());
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "MDRTF004");
    }

    [Fact]
    public void MarkdownToRtfDocumentKeepsEntitiesLiteralInsideCodeSpans() {
        RtfDocument document = MarkdownReader.Parse("`&lt;tag&gt;` &lt;tag&gt;").ToRtfDocument();

        RtfParagraph paragraph = Assert.Single(document.Paragraphs);

        Assert.Contains(paragraph.Runs, run => run.Text == "&lt;tag&gt;" && run.FontId.HasValue);
        Assert.Contains(paragraph.Runs, run => run.Text == " <tag>" && !run.FontId.HasValue);
    }

    [Fact]
    public void MarkdownToRtfDocument_Renders_SoftBreak_As_Space() {
        RtfDocument document = MarkdownReader.Parse("Alpha\nBeta").ToRtfDocument();

        RtfParagraph paragraph = Assert.Single(document.Paragraphs);

        Assert.Equal("Alpha Beta", paragraph.ToPlainText());
        Assert.DoesNotContain(paragraph.Inlines, inline => inline is RtfBreak);
    }

    [Fact]
    public void MarkdownRtfMarkdownRoundTripKeepsFencedCodeBlocks() {
        string markdown = """
            ```csharp
            # not heading
            - not list
            ```
            """;

        RtfDocument serialized = RtfDocument.Read(MarkdownReader.Parse(markdown).ToRtf()).Document;
        string roundTripMarkdown = serialized.ToMarkdown().Replace("\r\n", "\n");
        CodeBlock code = Assert.IsType<CodeBlock>(Assert.Single(MarkdownReader.Parse(roundTripMarkdown).Blocks));

        Assert.Contains("```csharp\n# not heading\n- not list\n```", roundTripMarkdown, StringComparison.Ordinal);
        Assert.Equal("csharp", code.Language);
        Assert.Equal("# not heading\n- not list", code.Content);
    }

    [Fact]
    public void MarkdownRtfMarkdownRoundTripKeepsFullFencedCodeInfoString() {
        string markdown = """
            ```c++ title="demo"
            int main() {}
            ```
            """;

        RtfDocument serialized = RtfDocument.Read(MarkdownReader.Parse(markdown).ToRtf()).Document;
        string roundTripMarkdown = serialized.ToMarkdown().Replace("\r\n", "\n");
        CodeBlock code = Assert.IsType<CodeBlock>(Assert.Single(MarkdownReader.Parse(roundTripMarkdown).Blocks));

        Assert.Contains("```c++ title=\"demo\"", roundTripMarkdown, StringComparison.Ordinal);
        Assert.Equal("c++ title=\"demo\"", code.InfoString);
        Assert.Equal("c++", code.Language);
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

        RtfDocument document = MarkdownReader.Parse(markdown).ToRtfDocument();

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
    public void MarkdownToRtfDocument_Does_Not_Drop_Repeated_List_Item_Paragraph_After_Nested_Block() {
        string markdown = """
            - repeat

              > quote

              repeat
            """;

        RtfDocument document = MarkdownReader.Parse(markdown).ToRtfDocument();
        var plainText = document.Paragraphs.Select(paragraph => paragraph.ToPlainText()).ToArray();

        Assert.Equal(2, plainText.Count(text => text == "repeat"));
        Assert.Contains("quote", plainText);
    }

    [Fact]
    public void MarkdownToRtfDocumentKeepsNestedListsInSameListDefinition() {
        RtfDocument document = MarkdownReader.Parse("""
            3. Parent
               - Child
            """).ToRtfDocument();

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
        RtfDocument document = MarkdownReader.Parse("""
            1. Parent
               5. Child
            """).ToRtfDocument();

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
        string rtf = MarkdownReader.Parse("""
            1. Parent
               5. Child
            """).ToRtf();

        Assert.Contains(@"\listoverridecount0", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\listoverridecount2", rtf, StringComparison.Ordinal);
        Assert.Equal(2, CountOccurrences(rtf, @"{\lfolevel"));
        Assert.Contains(@"{\lfolevel\listoverridestartat1\levelstartat5}", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownToRtfDocumentAppliesTableColumnAlignmentsToCellParagraphs() {
        RtfDocument document = MarkdownReader.Parse("""
            | Name | Count | Status |
            | :--- | ---: | :---: |
            | Alpha | 42 | Ready |
            """).ToRtfDocument();

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
        RtfDocument roundTrip = MarkdownReader.Parse(markdown).ToRtfDocument();
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
    public void RtfDocumentToMarkdownKeepsTaskMarkerPrefixesLiteralWithoutWhitespace() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("[x]ylophone").SetList(kind: RtfListKind.Bullet);

        string markdown = document.ToMarkdown();
        MarkdownDoc parsed = MarkdownReader.Parse(markdown);
        UnorderedListBlock list = Assert.IsType<UnorderedListBlock>(Assert.Single(parsed.Blocks));

        Assert.False(list.Items[0].IsTask);
        Assert.Equal("[x]ylophone", ExtractPlainText(list.Items[0].Content));
    }

    [Fact]
    public void RtfDocumentToMarkdownKeepsLiteralBlockMarkersAsText() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("# Heading\n- item\n1. item\n---\nTerm: Definition");

        string markdown = document.ToMarkdown();
        RtfDocument roundTrip = MarkdownReader.Parse(markdown).ToRtfDocument();

        Assert.Contains(@"\# Heading", markdown, StringComparison.Ordinal);
        Assert.Contains(@"\- item", markdown, StringComparison.Ordinal);
        Assert.Contains(@"1\. item", markdown, StringComparison.Ordinal);
        Assert.Contains(@"\---", markdown, StringComparison.Ordinal);
        Assert.Contains("Term&#58; Definition", markdown, StringComparison.Ordinal);
        Assert.All(roundTrip.Paragraphs, paragraph => {
            Assert.Equal(RtfListKind.None, paragraph.ListKind);
            Assert.Null(paragraph.OutlineLevel);
        });
        Assert.Contains(roundTrip.Paragraphs, paragraph => paragraph.ToPlainText().Contains("Term: Definition", StringComparison.Ordinal));
    }

    [Fact]
    public void RtfDocumentToMarkdownKeepsLiteralTildeAndHighlightMarkersParseable() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("~~not strike~~ ==not mark==");

        string markdown = document.ToMarkdown();
        RtfDocument roundTrip = MarkdownReader.Parse(markdown).ToRtfDocument();

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
        RtfDocument roundTrip = MarkdownReader.Parse(markdown).ToRtfDocument();

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
        RtfDocument roundTrip = MarkdownReader.Parse(markdown).ToRtfDocument();
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
        RtfDocument roundTrip = MarkdownReader.Parse(markdown).ToRtfDocument();
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
        RtfDocument roundTrip = MarkdownReader.Parse(markdown).ToRtfDocument();
        RtfTable roundTripTable = Assert.IsType<RtfTable>(roundTrip.Blocks.OfType<RtfTable>().Single());

        Assert.Contains("| Key: value |", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain(@"Key\: value", markdown, StringComparison.Ordinal);
        Assert.Equal("Key: value", roundTripTable.Rows[1].Cells[0].Paragraphs[0].ToPlainText());
    }

    [Fact]
    public void MarkdownToRtfDocumentPreservesDefinitionListTermsAndInlineFormatting() {
        RtfDocument document = MarkdownReader.Parse("Term: **Definition**").ToRtfDocument();

        RtfParagraph paragraph = Assert.Single(document.Paragraphs);
        Assert.Equal("Term: Definition", paragraph.ToPlainText());
        Assert.Contains(paragraph.Runs, run => run.Text == "Definition" && run.Bold);
    }

    [Fact]
    public void MarkdownToRtfDocumentPreservesFootnotesAsNotes() {
        string markdown = """
            Text[^1]

            [^1]: Note **bold**
            """;

        RtfDocument document = MarkdownReader.Parse(markdown).ToRtfDocument();

        RtfParagraph paragraph = Assert.Single(document.Paragraphs);
        RtfGeneratedText reference = Assert.Single(paragraph.Inlines.OfType<RtfGeneratedText>());
        Assert.NotNull(reference.Note);
        Assert.Equal(RtfNoteKind.Footnote, reference.Note!.Kind);
        Assert.Equal("Note bold", reference.Note.ToPlainText());
        Assert.Contains(reference.Note.Paragraphs[0].Runs, run => run.Text == "bold" && run.Bold);
    }

    [Fact]
    public void MarkdownToRtfDocumentPreservesRecursiveFootnoteReferenceAsText() {
        var options = new MarkdownToRtfOptions();
        string markdown = """
            Text[^1]

            [^1]: see [^1]
            """;

        RtfConversionResult<RtfDocument> result = MarkdownReader.Parse(markdown).ToRtfDocumentResult(options);
        RtfDocument document = result.Value;

        RtfGeneratedText reference = Assert.Single(document.Paragraphs[0].Inlines.OfType<RtfGeneratedText>());
        Assert.NotNull(reference.Note);
        Assert.Equal("see [^1]", reference.Note!.ToPlainText());
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "MDRTF020");
    }

    [Fact]
    public void MarkdownToRtfDocumentClearsStaleNestedOrderedStarts() {
        string markdown = """
            1. A
               5. first
            2. B
               1. second
            """;

        RtfDocument serialized = RtfDocument.Read(MarkdownReader.Parse(markdown).ToRtf()).Document;
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

        RtfDocument tableDocument = MarkdownReader.Parse("""
            | Name | Value |
            | --- | --- |
            | Alpha | Beta |
            """).ToRtfDocument();
        RtfTable table = Assert.IsType<RtfTable>(tableDocument.Blocks.OfType<RtfTable>().Single());
        Assert.True(table.Rows[0].RepeatHeader);
        string tableRoundTripMarkdown = tableDocument.ToMarkdown();
        Assert.Contains("| Name | Value |", tableRoundTripMarkdown, StringComparison.Ordinal);
        Assert.Contains("| --- | --- |", tableRoundTripMarkdown, StringComparison.Ordinal);

        RtfDocument escaped = MarkdownReader.Parse(@"&amp;lt; &amp;#42;").ToRtfDocument();
        Assert.Equal("&lt; &#42;", escaped.Paragraphs[0].ToPlainText());
    }

    [Fact]
    public void RtfDocumentToMarkdownKeepsOneRowTableParseable() {
        RtfDocument document = RtfDocument.Create();
        RtfTable table = document.AddTable(1, 1);
        table.Rows[0].Cells[0].AddParagraph("Only row");

        string markdown = document.ToMarkdown();
        RtfDocument roundTrip = MarkdownReader.Parse(markdown).ToRtfDocument();

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

        RtfDocument document = MarkdownReader.Parse("""
            <!-- OfficeIMO:RTF:HeaderlessSingleRowTable -->
            | Only row |
            """, options.ReaderOptions).ToRtfDocument(options);

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

        RtfDocument document = MarkdownReader.Parse(markdown).ToRtfDocument();
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
    public void RtfDocumentToMarkdownConvertsRunAttachedNotesToFootnotes() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph().AddFootnote("1", "Footnote body");
        var options = new RtfToMarkdownOptions();

        RtfConversionResult<string> result = document.ToMarkdownResult(options);
        string markdown = result.Value.Replace("\r\n", "\n");

        Assert.Contains("<sup>1</sup>[^fn1]", markdown, StringComparison.Ordinal);
        Assert.Contains("[^fn1]: Footnote body", markdown, StringComparison.Ordinal);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "RTFMD015");
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic => diagnostic.Code == "RTFMD012");
        result.RequireNoLoss();
    }

    [Fact]
    public void RtfDocumentToMarkdownPreserves_Rich_MultiParagraph_Note_Content() {
        RtfDocument document = RtfDocument.Create();
        var note = new RtfNote(RtfNoteKind.Footnote);
        RtfParagraph first = note.AddParagraph("Rich ");
        first.AddText("bold").SetBold();
        first.AddText(" and ");
        first.AddText("link").SetHyperlink(new Uri("https://example.test/note"));
        RtfParagraph second = note.AddParagraph("Second ");
        second.AddText("paragraph").SetItalic();
        document.AddParagraph("Body").AddNoteReference(note, "1");
        var options = new RtfToMarkdownOptions();

        RtfConversionResult<string> result = document.ToMarkdownResult(options);
        string markdown = result.Value.Replace("\r\n", "\n");
        MarkdownDoc parsed = MarkdownReader.Parse(markdown);
        FootnoteDefinitionBlock definition = Assert.IsType<FootnoteDefinitionBlock>(Assert.Single(parsed.Blocks, block => block is FootnoteDefinitionBlock));

        Assert.Contains("[^fn1]: Rich **bold** and [link](https://example.test/note)", markdown, StringComparison.Ordinal);
        Assert.Contains("Second *paragraph*", markdown, StringComparison.Ordinal);
        Assert.Equal(2, definition.ParagraphBlocks.Count);
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic => diagnostic.Action == RtfConversionAction.Omitted || diagnostic.Action == RtfConversionAction.Flattened);
        result.RequireNoLoss();
    }

    [Fact]
    public void RtfDocumentToMarkdownOmitsNotesAttachedToHiddenRuns() {
        RtfDocument document = RtfDocument.Create();
        RtfRun hiddenReference = document.AddParagraph().AddFootnote("1", "Hidden footnote body");
        hiddenReference.Hidden = true;

        string markdown = document.ToMarkdown(new RtfToMarkdownOptions { IncludeHiddenText = false });

        Assert.DoesNotContain("fn1", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("Hidden footnote body", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void RtfDocumentToMarkdownReportsGeneratedTextWithoutFallback() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph().AddPageNumber();
        var options = new RtfToMarkdownOptions();

        RtfConversionResult<string> result = document.ToMarkdownResult(options);
        string markdown = result.Value;

        Assert.Contains("RTF generated text omitted", markdown, StringComparison.Ordinal);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "RTFMD013");
    }

    [Fact]
    public void RtfDocumentToMarkdownReportsHeaderFooterOmission() {
        RtfDocument document = RtfDocument.Create();
        document.AddHeader().AddParagraph("Header");
        document.AddFooter().AddParagraph("Footer");
        document.AddParagraph("Body");
        var options = new RtfToMarkdownOptions();

        RtfConversionResult<string> result = document.ToMarkdownResult(options);
        string markdown = result.Value;

        Assert.Contains("Body", markdown, StringComparison.Ordinal);
        Assert.Contains("RTF header/footer content omitted", markdown, StringComparison.Ordinal);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "RTFMD014");
    }

    [Fact]
    public void RtfDocumentToMarkdownUsesParseStableInlineOmissionMarker() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph("Before ");
        paragraph.AddObject(RtfObjectKind.Embedded, new byte[] { 1, 2 });
        paragraph.AddText(" after");

        string markdown = document.ToMarkdown();
        RtfDocument roundTrip = MarkdownReader.Parse(markdown).ToRtfDocument();

        Assert.DoesNotContain("<!-- RTF object inline omitted", markdown, StringComparison.Ordinal);
        Assert.Contains("\\[RTF object inline omitted\\]", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("<!--", roundTrip.Paragraphs[0].ToPlainText(), StringComparison.Ordinal);
        Assert.Contains("[RTF object inline omitted]", roundTrip.Paragraphs[0].ToPlainText(), StringComparison.Ordinal);
    }

    [Fact]
    public void RtfDocumentToMarkdownKeepsMixedKindNestedListTogether() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Parent").SetList(listId: 7, level: 0, kind: RtfListKind.Decimal);
        document.AddParagraph("Child").SetList(listId: 7, level: 1, kind: RtfListKind.Bullet);
        document.AddParagraph("Next").SetList(listId: 7, level: 0, kind: RtfListKind.Decimal);

        string markdown = document.ToMarkdown().Replace("\r\n", "\n");
        RtfDocument roundTrip = MarkdownReader.Parse(markdown).ToRtfDocument();

        Assert.Contains("1. Parent\n\n   - Child\n\n2. Next", markdown, StringComparison.Ordinal);
        Assert.Contains(roundTrip.Paragraphs, paragraph => paragraph.ToPlainText() == "Parent" && paragraph.ListKind == RtfListKind.Decimal && paragraph.ListLevel == 0);
        Assert.Contains(roundTrip.Paragraphs, paragraph => paragraph.ToPlainText() == "Child" && paragraph.ListKind == RtfListKind.Bullet && paragraph.ListLevel == 1);
        Assert.Contains(roundTrip.Paragraphs, paragraph => paragraph.ToPlainText() == "Next" && paragraph.ListKind == RtfListKind.Decimal && paragraph.ListLevel == 0);
    }

    [Fact]
    public void RtfDocumentToMarkdownSplitsRestartedNestedOrderedLists() {
        RtfDocument document = RtfDocument.Create();
        RtfListDefinition parentDefinition = document.AddListDefinition(100);
        parentDefinition.AddLevel(RtfListKind.Decimal).StartAt = 1;
        document.AddListOverride(10, 100);
        RtfListDefinition firstChildDefinition = document.AddListDefinition(200);
        firstChildDefinition.AddLevel(RtfListKind.Decimal).StartAt = 1;
        firstChildDefinition.AddLevel(RtfListKind.Decimal).StartAt = 5;
        document.AddListOverride(20, 200);
        RtfListDefinition secondChildDefinition = document.AddListDefinition(300);
        secondChildDefinition.AddLevel(RtfListKind.Decimal).StartAt = 1;
        secondChildDefinition.AddLevel(RtfListKind.Decimal).StartAt = 1;
        document.AddListOverride(30, 300);

        document.AddParagraph("Parent A").SetList(listId: 10, level: 0, kind: RtfListKind.Decimal).ListDefinitionId = 100;
        document.AddParagraph("Five").SetList(listId: 20, level: 1, kind: RtfListKind.Decimal).ListDefinitionId = 200;
        document.AddParagraph("Parent B").SetList(listId: 10, level: 0, kind: RtfListKind.Decimal).ListDefinitionId = 100;
        document.AddParagraph("One").SetList(listId: 30, level: 1, kind: RtfListKind.Decimal).ListDefinitionId = 300;

        string markdown = document.ToMarkdown().Replace("\r\n", "\n");

        Assert.Contains("1. Parent A\n\n   5. Five\n\n2. Parent B\n\n   1. One", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("   6. One", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void RtfDocumentToMarkdownResolvesManyDistinctListIdsWithoutRepeatedTableScans() {
        const int listCount = 1024;
        RtfDocument document = RtfDocument.Create();
        for (int index = 1; index <= listCount; index++) {
            int definitionId = 10_000 + index;
            RtfListDefinition definition = document.AddListDefinition(definitionId);
            definition.AddLevel(RtfListKind.Decimal).StartAt = index;
            document.AddListOverride(index, definitionId);
            document.AddParagraph("Item " + index)
                .SetList(listId: index, level: 0, kind: RtfListKind.Decimal)
                .ListDefinitionId = definitionId;
        }

        string markdown = document.ToMarkdown();

        Assert.Contains("1. Item 1", markdown, StringComparison.Ordinal);
        Assert.Contains(listCount + ". Item " + listCount, markdown, StringComparison.Ordinal);
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

    [Fact]
    public void RtfDocumentToMarkdownEncodesSpacedHyperlinkDestinations() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph().AddText("file").SetHyperlink(new Uri("docs/My File.docx", UriKind.Relative));

        string markdown = document.ToMarkdown();
        RtfDocument roundTrip = MarkdownReader.Parse(markdown).ToRtfDocument();

        Assert.Contains("[file](docs/My%20File.docx)", markdown, StringComparison.Ordinal);
        Assert.Contains(roundTrip.Paragraphs[0].Runs, run => run.Text == "file" && run.Hyperlink != null);
    }

    [Fact]
    public void RtfDocumentToMarkdownEncodesFactoryImagePaths() {
        RtfDocument document = RtfDocument.Create();
        document.AddImage(RtfImageFormat.Png, new byte[] { 0x89, 0x50 }).Description = "Logo";
        var options = new RtfToMarkdownOptions {
            ImagePathFactory = (_, _) => "images/My File.png"
        };

        string markdown = document.ToMarkdown(options);
        MarkdownDoc parsed = MarkdownReader.Parse(markdown);
        ImageBlock image = Assert.IsType<ImageBlock>(Assert.Single(parsed.Blocks));

        Assert.Contains("![Logo](images/My%20File.png)", markdown, StringComparison.Ordinal);
        Assert.Equal("images/My%20File.png", image.Path);
    }

    [Fact]
    public void RtfDocumentToMarkdownExportsBlockAndInlineImagePayloads() {
        RtfDocument document = RtfDocument.Create();
        RtfImage block = document.AddImage(RtfImageFormat.Png, new byte[] { 0x89, 0x50 });
        RtfImage inline = document.AddParagraph().AddImage(RtfImageFormat.Jpeg, new byte[] { 0xFF, 0xD8 });
        var exported = new List<(RtfImage Image, int Index, string Path)>();
        var options = new RtfToMarkdownOptions {
            ImagePathFactory = (image, index) => "media/image " + index + "." + image.Format.ToString().ToLowerInvariant(),
            ImageExporter = (image, index, path) => exported.Add((image, index, path))
        };

        string markdown = document.ToMarkdown(options);

        Assert.Equal(2, exported.Count);
        Assert.Equal((block, 0, "media/image 0.png"), exported[0]);
        Assert.Equal((inline, 1, "media/image 1.jpeg"), exported[1]);
        Assert.Contains("media/image%200.png", markdown, StringComparison.Ordinal);
        Assert.Contains("media/image%201.jpeg", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void ReusingOptionsDoesNotLeakDiagnosticsAcrossConversions() {
        var options = new MarkdownToRtfOptions();

        RtfConversionResult<RtfDocument> lossy = MarkdownReader.Parse("![Logo](logo.png)").ToRtfDocumentResult(options);
        RtfConversionResult<RtfDocument> clean = MarkdownReader.Parse("Plain text").ToRtfDocumentResult(options);

        Assert.Contains(lossy.Report.Diagnostics, diagnostic => diagnostic.Code == "MDRTF003");
        Assert.Empty(clean.Report.Diagnostics);
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
