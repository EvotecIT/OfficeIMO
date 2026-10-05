using OfficeIMO.Core.Internal;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("PAGE", 0x21)]
    [InlineData("NUMPAGES", 0x1A)]
    [InlineData("DOCPROPERTY Title", 0x55)]
    [InlineData("REVNUM", 0x18)]
    public void LegacyDoc_SaveSupportedFieldWritesParsedFieldType(string instruction, int type) {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Field prefix Ω 😀 ");
        paragraph._paragraph.Append(new SimpleField(new Run(new Text("ResultMarker"))) { Instruction = instruction });
        byte[] bytes = document.ToBytes(WordFileFormat.Doc);
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out string? error), error);
        byte[] word = compound!.Streams["WordDocument"];
        byte[] table = compound.Streams["1Table"];
        int offset = BitConverter.ToInt32(word, 0x11A);
        Assert.Equal(22, BitConverter.ToInt32(word, 0x11E));
        Assert.Equal("Field prefix Ω 😀 ".Length, BitConverter.ToInt32(table, offset));
        Assert.Equal(type, table[offset + 17]);
        Assert.Equal(0x80, table[offset + 21]);
    }

    [Fact]
    public void LegacyDoc_SaveFieldInsideHyperlinkWritesNestedFieldEndFlags() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Nested prefix ");
        paragraph.AddHyperLink("NestedLinkMarker", new Uri("https://example.test/nested-field"), addStyle: true);
        paragraph.Hyperlink!._hyperlink.Append(new SimpleField(new Run(new Text("1"))) { Instruction = "PAGE" });
        byte[] bytes = document.ToBytes(WordFileFormat.Doc);
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out string? error), error);
        byte[] word = compound!.Streams["WordDocument"];
        byte[] table = compound.Streams["1Table"];
        int offset = BitConverter.ToInt32(word, 0x11A);
        Assert.Equal(40, BitConverter.ToInt32(word, 0x11E)); // Seven CPs and six field records.
        Assert.Equal(new byte[] { 0x13, 0x58, 0x14, 0, 0x13, 0x21, 0x14, 0, 0x15, 0xC0, 0x15, 0x80 },
            table.Skip(offset + 28).Take(12).ToArray());
        using WordDocument reopened = WordDocument.Load(new MemoryStream(bytes));
        Assert.Contains(reopened.HyperLinks, link => link.Uri?.ToString() == "https://example.test/nested-field");
    }

    [Theory]
    [InlineData("body", 0x11A)]
    [InlineData("cell", 0x11A)]
    [InlineData("header", 0x122)]
    [InlineData("footnote", 0x12A)]
    [InlineData("endnote", 0x21A)]
    public void LegacyDoc_SaveHyperlinkWritesValidFieldCharacterTable(string story, int fibOffset) {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph;
        if (story == "cell") paragraph = document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0];
        else if (story == "header") {
            document.AddHeadersAndFooters();
            paragraph = document.Sections[0].Header.Default!.AddParagraph();
        } else if (story == "footnote") {
            WordParagraph reference = document.AddParagraph("Footnote ").AddFootNote("Note prefix ");
            paragraph = reference.FootNote!.Paragraphs![1];
        } else if (story == "endnote") {
            WordParagraph reference = document.AddParagraph("Endnote ").AddEndNote("Note prefix ");
            paragraph = reference.EndNote!.Paragraphs![1];
        } else paragraph = document.AddParagraph("Body prefix ");
        paragraph.AddHyperLink("DOCLinkMarker", new Uri("https://example.test/doc-fields"), addStyle: true);
        byte[] bytes = document.ToBytes(WordFileFormat.Doc);
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out string? error), error);
        byte[] word = compound!.Streams["WordDocument"];
        byte[] table = compound.Streams["1Table"];
        int offset = BitConverter.ToInt32(word, fibOffset);
        int length = BitConverter.ToInt32(word, fibOffset + 4);
        Assert.Equal(22, length); // Four CP values and three two-byte field records.
        Assert.InRange(offset, 0, table.Length - length);
        int[] positions = Enumerable.Range(0, 4).Select(index => BitConverter.ToInt32(table, offset + index * 4)).ToArray();
        Assert.True(positions[0] < positions[1] && positions[1] < positions[2] && positions[2] < positions[3]);
        Assert.Equal(0x13, table[offset + 16]);
        Assert.Equal(88, table[offset + 17]); // MS-DOC fltHyperlink.
        Assert.Equal(0x14, table[offset + 18]);
        Assert.Equal(0x15, table[offset + 20]);
        Assert.Equal(0x80, table[offset + 21]); // The result separator is present.
        using WordDocument reopened = WordDocument.Load(new MemoryStream(bytes));
        IEnumerable<WordHyperLink> links = story switch {
            "header" => reopened.Sections[0].Header.Default!.HyperLinks,
            "footnote" => reopened.FootNotes.SelectMany(note => note.Paragraphs!)
                .Where(run => run.IsHyperLink).Select(run => run.Hyperlink!),
            "endnote" => reopened.EndNotes.SelectMany(note => note.Paragraphs!)
                .Where(run => run.IsHyperLink).Select(run => run.Hyperlink!),
            _ => reopened.HyperLinks
        };
        Assert.Contains(links, link => link.Text == "DOCLinkMarker" && link.Uri?.ToString() == "https://example.test/doc-fields");
    }
}
