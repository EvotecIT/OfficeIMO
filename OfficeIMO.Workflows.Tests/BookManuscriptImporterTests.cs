using OfficeIMO.Epub;
using OfficeIMO.Markdown;
using OfficeIMO.Word;
using System.Xml.Linq;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookManuscriptImporterTests {
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";

    [Fact]
    public async Task LocalImportRetainsSourceMetadataAndDoesNotReadAssetsOutsideItsParent() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-book-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Path.Combine(root, "book"));
        try {
            string source = Path.Combine(root, "book", "fallback-name.html");
            byte[] image = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=");
            await File.WriteAllBytesAsync(Path.Combine(root, "book", "dot.png"), image);
            await File.WriteAllBytesAsync(Path.Combine(root, "outside.png"), image);
            await File.WriteAllTextAsync(source, "<html lang='pl'><title>Source title</title><meta name='author' content='Writer'><h1>Chapter</h1><img alt='A dot' src='dot.png'></html>");
            var imported = await BookManuscriptImporter.ImportFileAsync(source);
            imported.Report.RequireNoLoss();
            Assert.Equal("Source title", imported.Publication.Title);
            Assert.Equal("pl", imported.Publication.Language);
            Assert.Equal("Writer", imported.Publication.Creator);
            Assert.Single(imported.Publication.Manifest, item => item.MediaType == "image/png");
            await File.WriteAllTextAsync(source, "<title>Blocked</title><h1>Chapter</h1><img alt='Outside' src='../outside.png'>");
            var blocked = await BookManuscriptImporter.ImportFileAsync(source);
            Assert.False(blocked.Succeeded);
            Assert.DoesNotContain(blocked.Publication.Manifest, item => item.MediaType == "image/png");
            await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => BookManuscriptImporter.ImportFileAsync(source, maximumInputBytes: 0));
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public async Task SnapshotImportRejectsExecutableWordPartsBeforeProjection() {
        using var word = WordDocument.Create();
        word.AddParagraph("Safe-looking body");
        using var package = new MemoryStream();
        package.Write(word.ToBytes());
        package.Position = 0;
        using (var archive = new System.IO.Compression.ZipArchive(package, System.IO.Compression.ZipArchiveMode.Update, leaveOpen: true)) {
            using var payload = archive.CreateEntry("word/vbaProject.bin").Open();
            payload.Write([1, 2, 3]);
        }
        await Assert.ThrowsAsync<OfficePackageSecurityException>(() => BookManuscriptImporter.ImportBytesAsync(package.ToArray(), ".docx", new EpubManuscriptOptions { Title = "Book" }));
    }

    [Fact]
    public async Task MarkdownImportPreservesHeadingsListsTablesCodeAndFootnoteLinks() {
        var markdown = MarkdownDoc.Parse("---\ntitle: A manuscript\nlanguage: pl\nauthor: Writer\n---\n# Opening\n\n[Next](#closing) and a note[^a].\n\n- List item\n\n| Name | Value |\n| --- | --- |\n| A | B |\n\n```cs\nvar answer = 42;\n```\n\n# Closing\n\nFinal text.\n\n[^a]: A preserved footnote.");
        EpubManuscriptResult result = await BookManuscriptImporter.ImportMarkdownAsync(markdown, new EpubManuscriptOptions());
        Assert.True(result.Succeeded, string.Join("; ", result.Report.FidelityDiagnostics.Select(item => item.Message)));
        Assert.Equal("A manuscript", result.Publication.Title);
        Assert.Equal("pl", result.Publication.Language);
        Assert.Equal("Writer", result.Publication.Creator);
        Assert.Equal(2, result.Publication.Spine.Count);
        XDocument first = result.Publication.GetContentXml("chapter-1");
        Assert.Single(first.Descendants(Html + "ul"));
        Assert.Single(first.Descendants(Html + "table"));
        Assert.Contains("answer", first.Descendants(Html + "code").Single().Value);
        Assert.Contains(first.Descendants(Html + "a"), link => ((string?)link.Attribute("href"))?.StartsWith("chapter-0002.xhtml#") == true);
        var reopened = result.RequireValue().Read();
        Assert.Contains("preserved footnote", reopened.Chapters[1].Text);
    }

    [Fact]
    public async Task WordImportReopensRealDocxAndRetainsNotesAcrossChapterBoundaries() {
        using var word = WordDocument.Create();
        word.AddParagraph("Opening").SetStyle(WordParagraphStyles.Heading1);
        word.AddParagraph("The first chapter").AddFootNote("Source note");
        word.AddParagraph("Closing").SetStyle(WordParagraphStyles.Heading1);
        word.AddParagraph("The last chapter");
        using var source = new MemoryStream(word.ToBytes());
        using var reopenedSource = WordDocument.Load(source);
        EpubManuscriptResult result = await BookManuscriptImporter.ImportWordAsync(reopenedSource,
            new EpubManuscriptOptions { Title = "DOCX manuscript", Creator = "Writer" });
        Assert.True(result.Succeeded, string.Join("; ", result.Report.FidelityDiagnostics.Select(item => item.Message)));
        Assert.Equal(2, result.Publication.Spine.Count);
        var reopened = result.RequireValue().Read();
        Assert.Contains("first chapter", reopened.Chapters[0].Text);
        Assert.Contains("Source note", reopened.Chapters[1].Text);
        Assert.Contains(result.Publication.GetContentXml("chapter-1").Descendants(Html + "a"), link =>
            ((string?)link.Attribute("href"))?.StartsWith("chapter-0002.xhtml#") == true);
    }
}
