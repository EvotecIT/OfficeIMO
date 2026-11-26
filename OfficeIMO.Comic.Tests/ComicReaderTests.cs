using System.IO.Compression;
using OfficeIMO.Epub;

namespace OfficeIMO.Comic.Tests;

public sealed class ComicReaderTests {
    [Fact]
    public void ImportsInfoZipComicInNumericOrderWithBookmarksAndCanvases() {
        using Stream source = File.OpenRead(Path.Combine(AppContext.BaseDirectory, "Fixtures", "harbour.cbz"));
        var imported = ComicReader.Import(source);
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
