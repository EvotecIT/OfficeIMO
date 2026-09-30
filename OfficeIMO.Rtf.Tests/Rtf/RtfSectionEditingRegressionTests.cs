using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfSectionEditingRegressionTests {
    [Fact]
    public void Document_Additions_Belong_To_The_Last_Section_And_Survive_Writing() {
        RtfDocument document = RtfDocument.Read(@"{\rtf1\ansi\sectd Existing\par}").Document;
        RtfParagraph added = document.AddParagraph("Added");
        RtfTable table = document.AddTable(1, 1);
        table.Rows[0].Cells[0].AddParagraph("Cell");
        RtfImage image = document.AddImage(RtfImageFormat.Png, new byte[] { 1 });
        RtfObject embedded = document.AddObject(RtfObjectKind.Embedded, new byte[] { 2 });
        RtfShape shape = document.AddShape();
        RtfSection section = Assert.Single(document.Sections);
        Assert.Equal(document.Blocks, section.Blocks);
        Assert.Contains(added, section.Blocks);
        Assert.Contains(table, section.Blocks);
        Assert.Contains(image, section.Blocks);
        Assert.Contains(embedded, section.Blocks);
        Assert.Contains(shape, section.Blocks);
        RtfDocument reparsed = RtfDocument.Read(document.ToRtf(), RtfReadOptions.CreateCompatibilityProfile()).Document;
        Assert.Contains(reparsed.Paragraphs, paragraph => paragraph.ToPlainText() == "Added");
        Assert.Single(reparsed.Blocks.OfType<RtfTable>());
        Assert.Single(reparsed.Sections);
    }

    [Fact]
    public void Adding_A_Section_Preserves_Preceding_Unsectioned_Content() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph before = document.AddParagraph("Before");
        RtfSection next = document.AddSection();
        RtfParagraph after = next.AddParagraph("After");
        Assert.Equal(new[] { before, after }, document.Paragraphs);
        Assert.Equal(new[] { before }, document.Sections[0].Blocks);
        Assert.Equal(new[] { after }, next.Blocks);
        Assert.Equal(new[] { "Before", "After" }, RtfDocument.Read(document.ToRtf()).Document.Paragraphs.Select(p => p.ToPlainText()));
    }

    [Fact]
    public void Editing_An_Earlier_Parsed_Section_Keeps_Document_Order() {
        RtfDocument document = RtfDocument.Read(@"{\rtf1\ansi\sectd First\par\sect\sectd Second\par}").Document;
        RtfParagraph added = document.Sections[0].AddParagraph("Between");
        Assert.Equal(new[] { "First", "Between", "Second" }, document.Paragraphs.Select(p => p.ToPlainText()));
        Assert.Same(added, document.Blocks[1]);
        Assert.Equal(new[] { "First", "Between", "Second" }, RtfDocument.Read(document.ToRtf()).Document.Paragraphs.Select(p => p.ToPlainText()));
    }
}
