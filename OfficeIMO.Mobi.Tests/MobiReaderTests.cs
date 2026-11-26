using System.Buffers.Binary;
using System.Text;
using OfficeIMO.Epub;

namespace OfficeIMO.Mobi.Tests;

public sealed class MobiReaderTests {
    [Theory]
    [InlineData("harbour.mobi", false)]
    [InlineData("harbour.azw3", true)]
    [InlineData("harbour-combo.mobi", true)]
    public void ImportsCalibreBookWithCoverNavigationAndLinks(string name, bool nested) {
        using var source = new MemoryStream(Fixture(name), writable: false);
        EpubDocument book = ReadPublication(source, name);
        Assert.Equal("Harbour Signal", book.Title);
        Assert.Equal("Ada Reader", book.Creator);
        Assert.Contains(book.Metadata, item => item.Name == "publisher" && item.Value == "Signal Press");
        Assert.Equal("Opening Signal", book.TableOfContents[0].Label);
        Assert.Equal(nested ? 1 : 2, book.TableOfContents.Count);
        var ending = nested ? Assert.Single(book.TableOfContents[0].Children) : book.TableOfContents[1];
        Assert.Equal("Closing Signal", ending.Label);
        Assert.NotNull(ending.Fragment);
        var opening = book.Chapters.First(chapter => chapter.Text.Contains("A small independent ebook"));
        Assert.Contains("https://example.org/reference", opening.Html);
        Assert.Contains("href=\"" + Path.GetFileName(ending.Target) + "#" + ending.Fragment + "\">Read the ending", opening.Html);
        var destination = book.Chapters.Single(chapter => chapter.Path == ending.Target);
        Assert.Contains("id=\"" + ending.Fragment + "\"", destination.Html);
        Assert.Contains("Signal received.", destination.Text);
        Assert.Contains(book.Resources, item => item.Properties == "cover-image" && item.Data?.Length > 0);
        Assert.Contains(book.Resources, item => item.Id == "image-3" && item.Data?.Length > 0);
    }

    [Fact]
    public void ReadsCalibrePalmDocAsLiteralText() {
        using var source = new MemoryStream(Fixture("harbour.pdb"), writable: false);
        EpubDocument book = ReadPublication(source, "harbour.pdb");
        Assert.Equal("Harbour Signal", book.Title);
        Assert.Contains("Opening Signal", Assert.Single(book.Chapters).Text);
        Assert.Contains("Signal received.", book.Chapters[0].Text);
    }

    [Theory]
    [InlineData("harbour.azw3", 0)]
    [InlineData("harbour-combo.mobi", 11)]
    public void RefusesDrmInStandaloneAndDualFormatHeaders(string name, int header) {
        byte[] bytes = Fixture(name);
        int offset = BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(78 + 8 * header));
        bytes[offset + 13] = 2;
        using var source = new MemoryStream(bytes, writable: false);
        var error = Assert.Throws<NotSupportedException>(() => MobiReader.Import(source, name));
        Assert.Contains("DRM-protected", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RefusesAnInvalidPalmDocBackReference() {
        byte[] bytes = Fixture("harbour.pdb");
        int offset = BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(86));
        bytes[offset] = 0x80;
        bytes[offset + 1] = 0x08;
        using var source = new MemoryStream(bytes, writable: false);
        var error = Assert.Throws<InvalidDataException>(() => MobiReader.Import(source, "harbour.pdb"));
        Assert.Contains("back-reference", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RejectsHtmlThatExceedsTheParserElementBudget() {
        byte[] existing = Fixture("harbour.mobi");
        int headerStart = BinaryPrimitives.ReadInt32BigEndian(existing.AsSpan(78));
        int headerEnd = BinaryPrimitives.ReadInt32BigEndian(existing.AsSpan(86));
        byte[] header = existing.AsSpan(headerStart, headerEnd - headerStart).ToArray();
        byte[] text = Encoding.ASCII.GetBytes("<html><body>" + string.Concat(Enumerable.Repeat("<br>", 155_000)) + "</body></html>");
        int records = (text.Length + 4095) / 4096;
        int directoryEnd = 78 + (records + 1) * 8;
        var data = new byte[directoryEnd + 2 + header.Length + text.Length];
        existing.AsSpan(0, 78).CopyTo(data);
        BinaryPrimitives.WriteUInt16BigEndian(data.AsSpan(76), (ushort)(records + 1));
        int newHeader = directoryEnd + 2;
        BinaryPrimitives.WriteInt32BigEndian(data.AsSpan(78), newHeader);
        for (int record = 0; record < records; record++)
            BinaryPrimitives.WriteInt32BigEndian(data.AsSpan(86 + 8 * record), newHeader + header.Length + record * 4096);
        BinaryPrimitives.WriteUInt16BigEndian(header, 1);
        BinaryPrimitives.WriteInt32BigEndian(header.AsSpan(4), text.Length);
        BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(8), (ushort)records);
        BinaryPrimitives.WriteUInt32BigEndian(header.AsSpan(108), uint.MaxValue);
        BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(242), 0);
        header.CopyTo(data, newHeader);
        text.CopyTo(data, newHeader + header.Length);

        using var stream = new MemoryStream(data, writable: false);
        var error = Assert.Throws<InvalidDataException>(() => MobiReader.Import(stream));
        Assert.Contains("aggregate token or element limit", error.Message, StringComparison.Ordinal);
    }

    private static EpubDocument ReadPublication(Stream source, string name) {
        using var epub = new MemoryStream();
        MobiReader.Import(source, name).Value.Save(epub);
        epub.Position = 0;
        return EpubDocument.Load(epub, new EpubLoadOptions { IncludeRawHtml = true, IncludeResourceData = true });
    }

    private static byte[] Fixture(string name) => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", name));
}
