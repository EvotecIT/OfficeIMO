using System.IO.Compression;
using OfficeIMO.Epub;

namespace OfficeIMO.Comic.Tests;

public sealed class ComicReaderTests {
    [Theory]
    [InlineData("harbour.cbz", "harbour.cbz")]
    [InlineData("harbour.cb7", "harbour.cb7")]
    [InlineData("harbour.cbr", "harbour.cbr")]
    [InlineData("harbour.cbz", "harbour.cbr")]
    [InlineData("harbour.cb7", "harbour.cbr")]
    public void ImportsProducerComicsInNumericOrderWithBookmarksAndCanvases(string fixture, string name) {
        using Stream source = File.OpenRead(Path.Combine(AppContext.BaseDirectory, "Fixtures", fixture));
        var imported = ComicReader.Import(source, name);
        using var artifact = new MemoryStream();
        imported.Value.Save(artifact);
        artifact.Position = 0;
        var book = EpubDocument.Load(artifact, new EpubLoadOptions { IncludeRawHtml = true, IncludeResourceData = true });
        Assert.Equal("The Harbour Comic", book.Title);
        Assert.Equal("Ada Reader", book.Creator);
        Assert.Equal(new[] { "Cover", "The Signal", "The Ending" }, book.TableOfContents.Select(item => item.Label));
        Assert.Equal(3, book.Chapters.Count);
        Assert.All(book.Chapters, chapter => Assert.True(chapter.IsFixedLayout));
        Assert.Contains("width=120", book.Chapters[1].Html);
        Assert.Contains("width=144", book.Chapters[2].Html);
        Assert.Contains(book.Resources, item => item.Path.EndsWith("image-1.png", StringComparison.Ordinal) && item.Properties == "cover-image");
    }

    [Fact]
    public void Rar4NamedCbrRetainsTheUnsupportedVersionDiagnostic() {
        using Stream source = File.OpenRead(Path.Combine(AppContext.BaseDirectory, "Fixtures", "legacy.rar"));
        var error = Assert.Throws<NotSupportedException>(() => ComicReader.Import(source, "legacy.cbr"));
        Assert.Contains("RAR 4 not supported", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RejectsGifHeaderWithoutFrameData() {
        using var source = new MemoryStream();
        using (var zip = new ZipArchive(source, ZipArchiveMode.Create, leaveOpen: true))
        using (Stream entry = zip.CreateEntry("1.gif").Open()) {
            // Valid GIF signature and logical screen alone do not constitute an image.
            byte[] header = { 71, 73, 70, 56, 57, 97, 1, 0, 1, 0, 0, 0, 0 };
            entry.Write(header, 0, header.Length);
        }
        source.Position = 0;
        Assert.Throws<InvalidDataException>(() => ComicReader.Import(source));
    }
}
