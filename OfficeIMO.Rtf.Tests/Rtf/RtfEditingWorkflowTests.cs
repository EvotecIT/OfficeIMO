using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Syntax;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfEditingWorkflowTests {
    [Fact]
    public void Semantic_Block_Editing_Synchronizes_Document_And_Section_Order() {
        RtfDocument document = RtfDocument.Create();
        RtfSection section = document.AddSection();
        RtfParagraph first = section.AddParagraph("First");
        RtfParagraph last = section.AddParagraph("Last");

        RtfParagraph middle = document.InsertParagraph(1, "Middle");
        document.MoveBlock(0, 2);
        IRtfBlock removed = document.RemoveBlockAt(1);

        Assert.Same(last, removed);
        Assert.Equal(new[] { middle, first }, document.Blocks.Cast<RtfParagraph>());
        Assert.Equal(new[] { middle, first }, section.Blocks.Cast<RtfParagraph>());
        Assert.Equal(new[] { "Middle", "First" }, document.Paragraphs.Select(paragraph => paragraph.ToPlainText()));
        Assert.Equal(new[] { "Middle", "First" }, RtfDocument.Read(document.ToRtf()).Document.Paragraphs.Select(paragraph => paragraph.ToPlainText()));
    }

    [Fact]
    public void Semantic_Clone_Is_Independent() {
        RtfDocument source = RtfDocument.Create();
        source.AddParagraph("Original").AddText(" styled").SetBold();
        source.AddTable(1, 1).Rows[0].Cells[0].AddParagraph("Cell");

        RtfDocument clone = source.Clone();
        clone.ReplaceText("Original", "Clone");
        clone.InsertParagraph(1, "Inserted");

        Assert.Equal("Original styled", source.Paragraphs[0].ToPlainText());
        Assert.Equal("Clone styled", clone.Paragraphs[0].ToPlainText());
        Assert.Equal(2, source.Blocks.Count);
        Assert.Equal(3, clone.Blocks.Count);
    }

    [Fact]
    public void Semantic_Document_Append_Remaps_Resources_And_Reports_Flattened_Bindings() {
        RtfDocument destination = RtfDocument.Create();
        destination.AddColor(255, 0, 0);
        destination.AddParagraph("Destination");
        RtfDocument source = RtfDocument.Create();
        int blue = source.AddColor(0, 0, 255);
        int consolas = source.AddFont("Consolas");
        source.AddStyle(5, "Imported style");
        RtfParagraph sourceParagraph = source.AddParagraph();
        sourceParagraph.StyleId = 5;
        sourceParagraph.SetList(7, 0, RtfListKind.Decimal);
        RtfRun sourceRun = sourceParagraph.AddText("Imported");
        sourceRun.FontId = consolas;
        sourceRun.ForegroundColorIndex = blue;
        sourceRun.StyleId = 5;
        source.AddHeader().AddParagraph("Source header");

        RtfDocumentMergeResult result = destination.AppendDocument(source);
        RtfParagraph imported = Assert.IsType<RtfParagraph>(destination.Blocks[1]);
        RtfRun importedRun = Assert.Single(imported.Runs);

        Assert.Equal(1, result.AppendedBlockCount);
        Assert.Equal("Imported", imported.ToPlainText());
        Assert.Null(imported.StyleId);
        Assert.Null(imported.ListId);
        Assert.Null(imported.ListDefinitionId);
        Assert.Equal(2, importedRun.ForegroundColorIndex);
        Assert.Equal("Consolas", destination.Fonts.Single(font => font.Id == importedRun.FontId).Name);
        Assert.Null(importedRun.StyleId);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "RtfMergeStylesFlattened");
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "RtfMergeListsFlattened");
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "RtfMergeHeaderFootersOmitted");
        Assert.Throws<RtfConversionLossException>(() => result.Report.RequireNoLoss());
        Assert.Equal(1, sourceRun.ForegroundColorIndex);
        Assert.Equal(5, sourceRun.StyleId);
    }

    [Fact]
    public void Semantic_Document_Append_Of_Plain_Content_Is_NoLoss() {
        RtfDocument destination = RtfDocument.Create();
        destination.AddParagraph("A");
        RtfDocument source = RtfDocument.Create();
        source.AddParagraph("B");

        RtfDocumentMergeResult result = destination.AppendDocument(source);

        Assert.Equal(new[] { "A", "B" }, destination.Paragraphs.Select(paragraph => paragraph.ToPlainText()));
        Assert.False(result.Report.HasLoss);
        result.Report.RequireNoLoss();
    }

    [Fact]
    public void Semantic_Document_Append_Remaps_Attached_Notes_Exactly_Once() {
        RtfDocument destination = RtfDocument.Create();
        destination.AddColor(255, 0, 0);
        RtfDocument source = RtfDocument.Create();
        int blue = source.AddColor(0, 0, 255);
        source.AddColor(0, 255, 0);
        RtfNote note = source.AddNote(RtfNoteKind.Footnote);
        RtfRun noteRun = note.AddParagraph().AddText("Blue note");
        noteRun.ForegroundColorIndex = blue;
        source.AddParagraph("Reference").AddNoteReference(note, "1");

        destination.AppendDocument(source);

        RtfNote imported = Assert.Single(destination.Notes);
        Assert.Equal(2, Assert.Single(Assert.Single(imported.Paragraphs).Runs).ForegroundColorIndex);
        RtfGeneratedText reference = Assert.IsType<RtfGeneratedText>(Assert.Single(destination.Paragraphs[0].Inlines.Skip(1)));
        Assert.Same(imported, reference.Note);
    }

    [Fact]
    public void Rich_Replacement_Crosses_Run_Boundaries_And_Preserves_Unchanged_Suffix_Formatting() {
        RtfParagraph paragraph = new RtfParagraph();
        RtfRun first = paragraph.AddText("Hel").SetBold();
        RtfRun second = paragraph.AddText("lo ").SetItalic();
        RtfRun third = paragraph.AddText("World").SetUnderline();

        int replacements = paragraph.ReplaceText("Hello", "Hi");

        Assert.Equal(1, replacements);
        Assert.Equal("Hi World", paragraph.ToPlainText());
        Assert.Equal("Hi", first.Text);
        Assert.Equal(" ", second.Text);
        Assert.True(second.Italic);
        Assert.Equal("World", third.Text);
        Assert.True(third.Underline);
    }

    [Fact]
    public void Bookmark_Replacement_Works_Across_Paragraphs() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph first = document.AddParagraph("Before ");
        first.AddBookmarkStart("Patient");
        first.AddText("Old first").SetBold();
        RtfParagraph second = document.AddParagraph("Old second");
        second.AddBookmarkEnd("Patient");
        second.AddText(" after");

        bool replaced = document.ReplaceBookmarkText("Patient", "New value");

        Assert.True(replaced);
        Assert.Equal("Before New value", first.ToPlainText());
        Assert.Equal(" after", second.ToPlainText());
        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        Assert.Contains(@"{\*\bkmkstart Patient}New value\par", rtf, StringComparison.Ordinal);
        Assert.Contains(@"{\*\bkmkend Patient} after", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void Lossless_Root_Structure_Editing_Preserves_Trailing_Bytes() {
        const string source = "{\\rtf1\\ansi\\pard A\\par\\pard B\\par}\r\n";
        RtfLosslessEditor editor = RtfDocument.Read(source).EditLossless();
        int firstParagraphNode = editor.SyntaxTree.Root.Children
            .Select((node, index) => new { node, index })
            .First(item => item.node is RtfControlWord control && control.Name == "pard").index;
        int originalNodeCount = editor.RootNodeCount;

        editor.InsertRootParagraph(originalNodeCount, "C");
        editor.MoveRootNodes(originalNodeCount, 3, firstParagraphNode);
        editor.RemoveRootNodes(firstParagraphNode + 3, 3);

        Assert.EndsWith("}\r\n", editor.ToRtf(), StringComparison.Ordinal);
        Assert.Equal(new[] { "C", "B" }, editor.ToReadResult().Document.Paragraphs.Select(paragraph => paragraph.ToPlainText()));
    }

    [Fact]
    public void Lossless_Editor_Replaces_Image_And_Header_Content_Without_Touching_Unknown_Syntax() {
        const string source = "{\\rtf1\\ansi{\\header Old header}{\\pict\\pngblip 010203}{\\*\\unknown Keep}\\pard Body\\par}\r\n";
        RtfLosslessEditor editor = RtfDocument.Read(source).EditLossless();
        var image = new RtfImage(RtfImageFormat.Jpeg, new byte[] { 0xFF, 0xD8, 0xFF, 0xD9 }) {
            SourceWidth = 1,
            SourceHeight = 1
        };

        Assert.True(editor.ReplaceImage(0, image));
        Assert.Equal(1, editor.ReplaceDestinationContent("header", @"\pard New header\par"));

        string edited = editor.ToRtf();
        Assert.Contains(@"{\pict\jpegblip\picw1\pich1", edited, StringComparison.Ordinal);
        Assert.Contains(@"{\*\unknown Keep}", edited, StringComparison.Ordinal);
        Assert.EndsWith("}\r\n", edited, StringComparison.Ordinal);
        RtfReadResult result = editor.ToReadResult();
        Assert.Equal("New header", Assert.Single(result.Document.HeaderFooters).ToPlainText());
        Assert.Equal(new byte[] { 0xFF, 0xD8, 0xFF, 0xD9 }, Assert.Single(result.Document.Blocks.OfType<RtfImage>()).Data);
    }
}
