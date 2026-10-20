using System.IO.Compression;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void AddFootNote_RunningParagraphRejectsWithoutChangingSource(bool footer, bool tableCell) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Body remains unchanged");
        document.AddHeadersAndFooters();
        WordHeaderFooter story = footer
            ? RequireSectionFooter(document, 0, HeaderFooterValues.Default)
            : RequireSectionHeader(document, 0, HeaderFooterValues.Default);
        WordParagraph paragraph = tableCell
            ? story.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0]
            : story.AddParagraph();
        paragraph.Text = "Running paragraph";
        SortedDictionary<string, byte[]> before = ReadFootnoteGuardPackageParts(document.ToBytes());

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => paragraph.AddFootNote("Rejected note"));

        Assert.Contains("header or footer", exception.Message, StringComparison.Ordinal);
        Assert.Equal("Running paragraph", paragraph.Text);
        Assert.Empty(document.FootNotes);
        SortedDictionary<string, byte[]> after = ReadFootnoteGuardPackageParts(document.ToBytes());
        Assert.Equal(before.Keys, after.Keys);
        foreach (KeyValuePair<string, byte[]> part in before) {
            Assert.Equal(part.Value, after[part.Key]);
        }
        Assert.Empty(document.ValidateDocument());
    }

    [Fact]
    public void AddFootNote_BodyAndBodyTableParagraphsRetainValidNotesAfterReload() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Body reference").AddFootNote("Body note");
        WordParagraph cellParagraph = document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0];
        cellParagraph.Text = "Cell reference";
        cellParagraph.AddFootNote("Cell note");

        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        WordFootNote bodyNote = Assert.IsType<WordFootNote>(Assert.Single(loaded.Paragraphs, paragraph => paragraph.IsFootNote).FootNote);
        WordFootNote cellNote = Assert.IsType<WordFootNote>(Assert.Single(loaded.Tables[0].Rows[0].Cells[0].Paragraphs, paragraph => paragraph.IsFootNote).FootNote);

        Assert.Collection(new[] { bodyNote, cellNote },
            note => {
                Assert.Equal(1L, note.ReferenceId);
                Assert.Equal("Body note", string.Concat(note.Paragraphs!.Select(paragraph => paragraph.Text)));
                Assert.Equal("Body reference", note.ParentParagraph!.Text);
            },
            note => {
                Assert.Equal(2L, note.ReferenceId);
                Assert.Equal("Cell note", string.Concat(note.Paragraphs!.Select(paragraph => paragraph.Text)));
                Assert.Equal("Cell reference", note.ParentParagraph!.Text);
            });
        Assert.Empty(loaded.ValidateDocument());
    }

    private static SortedDictionary<string, byte[]> ReadFootnoteGuardPackageParts(byte[] bytes) {
        using var archive = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        var parts = new SortedDictionary<string, byte[]>(StringComparer.Ordinal);
        foreach (ZipArchiveEntry entry in archive.Entries) {
            using Stream source = entry.Open();
            using var content = new MemoryStream();
            source.CopyTo(content);
            parts.Add(entry.FullName, content.ToArray());
        }
        return parts;
    }
}
