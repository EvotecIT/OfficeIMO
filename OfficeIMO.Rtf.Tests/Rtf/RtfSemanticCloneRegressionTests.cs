using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfSemanticCloneRegressionTests {
    [Fact]
    public void Clone_Preserves_Payloads_File_References_And_Independent_Formatting() {
        RtfDocument source = RtfDocument.Create();
        source.AddFileReference("file:///C:/Documents/source.docx");
        RtfObject original = source.AddObject(RtfObjectKind.Embedded, new byte[] { 1, 2, 3 });
        original.Result.AddText("Fallback").SetBold();
        source.Fonts[0].Embedding = new RtfFontEmbedding { Data = new byte[] { 7, 8 } };
        source.Info.Title = "Original";

        RtfDocument clone = source.Clone();
        RtfObject copied = Assert.IsType<RtfObject>(Assert.Single(clone.Blocks));
        Assert.Equal(original.Data, copied.Data);
        Assert.Equal(source.FileReferences[0].Path, Assert.Single(clone.FileReferences).Path);
        Assert.NotSame(original.Result, copied.Result);
        copied.Data[0] = 9;
        copied.Result.Runs[0].Bold = false;
        clone.Fonts[0].Embedding!.Data[0] = 10;
        clone.Info.Title = "Copy";

        Assert.Equal(1, original.Data[0]);
        Assert.True(original.Result.Runs[0].Bold);
        Assert.Equal(7, source.Fonts[0].Embedding!.Data[0]);
        Assert.Equal("Original", source.Info.Title);
    }

    [Fact]
    public void Clone_Preserves_Shared_Note_And_Section_References_Inside_Independent_Graph() {
        RtfDocument source = RtfDocument.Create();
        RtfSection section = source.AddSection();
        RtfParagraph paragraph = section.AddParagraph("Body");
        RtfNote note = source.AddNote(RtfNoteKind.Footnote);
        note.AddParagraph("Note");
        paragraph.AddNoteReference(note, "A");
        paragraph.AddNoteReference(note, "B");

        RtfDocument clone = source.Clone();
        RtfParagraph copied = Assert.Single(clone.Paragraphs);
        Assert.Same(copied, Assert.Single(clone.Blocks));
        Assert.Same(copied, Assert.Single(clone.Sections[0].Blocks));
        Assert.Same(copied.Runs[0], copied.Inlines[0]);
        foreach (RtfGeneratedText reference in copied.Inlines.OfType<RtfGeneratedText>()) {
            Assert.Same(Assert.Single(clone.Notes), reference.Note);
            Assert.NotSame(note, reference.Note);
        }
        clone.Sections[0].AddParagraph("New");
        Assert.Equal(2, clone.Blocks.Count);
        Assert.Single(source.Blocks);
    }

    [Fact]
    public void Append_Does_Not_Apply_Ingestion_Filters_To_Already_Owned_Content() {
        RtfDocument source = RtfDocument.Create();
        source.AddParagraph("Body");
        RtfObject original = source.AddObject(RtfObjectKind.Embedded, new byte[] { 1, 2 });
        original.Result.AddText("Object");
        RtfDocument destination = RtfDocument.Create();

        RtfDocumentMergeResult result = destination.AppendDocument(source);

        Assert.Equal(2, result.AppendedBlockCount);
        RtfObject copied = Assert.IsType<RtfObject>(destination.Blocks[1]);
        Assert.Equal(original.Data, copied.Data);
        copied.Data[0] = 3;
        Assert.Equal(1, original.Data[0]);
    }
}
