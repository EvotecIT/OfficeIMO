namespace OfficeIMO.OneNote.Tests;

public sealed class SectionReaderTests {
    [Fact]
    public void ReadsPageTitleAndParagraphText() {
        OneNoteSection section = OneNoteSectionReader.Read(FixturePath("testOneNote2016.one"));

        OneNotePage page = Assert.Single(section.Pages);
        Assert.Equal(0x00E4A88AU, section.ColorArgb);
        Assert.Equal("So good", page.Title);
        Assert.Contains("This is one note 2016", ExtractText(page));
        Assert.NotEmpty(page.Revisions);
        Assert.Single(page.Revisions, revision => revision.IsCurrent);
    }

    [Fact]
    public void MaterializesNestedOutlineTextWithoutFontJunk() {
        OneNoteSection section = OneNoteSectionReader.Read(FixturePath("testOneNote.one"));

        OneNotePage page = Assert.Single(section.Pages);
        string text = ExtractText(page);
        Assert.Equal("Note-ssn-test-mmmm", page.Title);
        Assert.Contains("Nicole Knox", text);
        Assert.DoesNotContain("Calibri", text);
    }

    [Fact]
    public void ResolvesEmbeddedWordPayloadFromFileDataStore() {
        OneNoteSection section = OneNoteSectionReader.Read(FixturePath("testOneNoteEmbeddedWordDoc.one"));

        OneNoteEmbeddedFile embedded = Assert.Single(section.Pages.SelectMany(page => page.DirectContent).OfType<OneNoteEmbeddedFile>());
        Assert.Equal("Dude this is a super cool embedded doc.docx", embedded.FileName);
        Assert.NotNull(embedded.Payload);
        byte[] payload = embedded.Payload!.ToArray(OneNoteReaderOptions.DefaultMaxAssetBytes);
        Assert.True(payload.Length > 4);
        Assert.Equal((byte)'P', payload[0]);
        Assert.Equal((byte)'K', payload[1]);
    }

    [Theory]
    [InlineData("testOneNoteFromOffice365.one", "testOneNoteFromOffice365")]
    [InlineData("testOneNoteFromOffice365-2.one", "Section1.one")]
    public void ReadsOffice365PackageStoreIntoSameSemanticModel(string fixture, string expectedSectionName) {
        OneNoteSection section = OneNoteSectionReader.Read(FixturePath(fixture));

        Assert.Equal(expectedSectionName, section.Name);
        Assert.Equal(2, section.Pages.Count);
        Assert.Equal(new[] { "Section1Page1", "Section1Page2" }, section.Pages.Select(page => page.Title).ToArray());
        Assert.Contains("Section1Page1Content", ExtractText(section.Pages[0]));
        Assert.Contains("Section1Page2Content", ExtractText(section.Pages[1]));
        Assert.All(section.Pages, page => Assert.NotEmpty(page.Revisions));
    }

    [Fact]
    public void ReadsVersionHistoryFromItsReferencedRevisionContext() {
        OneNoteSection section = OneNoteSectionReader.Read(FixturePath("testOneNoteFromOffice365.one"));

        OneNotePage version = Assert.Single(section.Pages.SelectMany(page => page.VersionHistory));
        Assert.True(version.IsVersionHistoryPage);
        Assert.NotNull(version.RevisionContextId);
        Assert.NotEmpty(version.Revisions);
        Assert.All(version.Revisions, revision => Assert.True(revision.IsVersionHistory));
        Assert.Equal("Thursday, November 11, 2021 5:03 PM", version.Title);
    }

    private static string ExtractText(OneNotePage page) {
        return string.Join("\n", page.Outlines.Cast<OneNoteElement>()
            .Concat(page.DirectContent)
            .SelectMany(ExtractText));
    }

    private static IEnumerable<string> ExtractText(OneNoteElement element) {
        if (element is OneNoteParagraph paragraph) {
            foreach (OneNoteTextRun run in paragraph.Runs) yield return run.Text;
            foreach (OneNoteElement child in paragraph.Children)
            foreach (string text in ExtractText(child)) yield return text;
        } else if (element is OneNoteOutline outline) {
            foreach (OneNoteElement child in outline.Children)
            foreach (string text in ExtractText(child)) yield return text;
        } else if (element is OneNoteTable table) {
            foreach (OneNoteTableRow row in table.Rows)
            foreach (OneNoteTableCell cell in row.Cells)
            foreach (OneNoteElement child in cell.Content)
            foreach (string text in ExtractText(child)) yield return text;
        }
    }

    private static string FixturePath(string fileName) => Path.Combine(AppContext.BaseDirectory, "Fixtures", fileName);
}
