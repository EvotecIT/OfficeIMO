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
}
