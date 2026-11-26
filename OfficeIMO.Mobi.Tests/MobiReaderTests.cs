using System.Buffers.Binary;
using System.Text;
using System.Xml.Linq;
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
    public void ReadsSingleFlowKf8WithAnUnusedFlowPointer() {
        byte[] bytes = Fixture("harbour.azw3");
        int header = RecordOffset(bytes, 0);
        BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(header + 192), 0);
        BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(header + 196), 1);
        using var source = new MemoryStream(bytes, writable: false);

        EpubDocument book = ReadPublication(source, "harbour.azw3");

        Assert.Contains(book.Chapters, chapter => chapter.Text.Contains("Signal received.", StringComparison.Ordinal));
    }

    [Fact]
    public void ReadsDualKf8WithNoDeclaredResources() {
        byte[] bytes = Fixture("harbour-combo.mobi");
        int header = RecordOffset(bytes, 0);
        int count = FindExthItem(bytes, header, 125);
        BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(count + 8), 0);
        BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(header + 108), 12);
        using var source = new MemoryStream(bytes, writable: false);

        EpubDocument book = ReadPublication(source, "harbour-combo.mobi");

        Assert.Contains(book.Chapters, chapter => chapter.Text.Contains("Signal received.", StringComparison.Ordinal));
        Assert.DoesNotContain(book.Resources, item => item.Properties == "cover-image");
    }

    [Theory]
    [InlineData(103)]
    [InlineData(104)]
    public void OmitsOptionalMetadataThatDecodesToWhitespace(int tag) {
        byte[] bytes = Fixture("harbour.azw3");
        int item = FindExthItem(bytes, RecordOffset(bytes, 0), 205);
        Assert.Equal(12, BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(item + 4)));
        BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(item), (uint)tag);
        Encoding.ASCII.GetBytes("&#9;").CopyTo(bytes, item + 8);
        using var source = new MemoryStream(bytes, writable: false);

        EpubDocument book = ReadPublication(source, "harbour.azw3");

        Assert.Equal("Ada Reader", book.Creator);
        Assert.DoesNotContain(book.Metadata, value => value.Name == "description");
        Assert.Equal("c4b5bb5a-5519-44a6-bada-106df8645d5a", book.Identifier);
        Assert.Contains(book.Chapters, chapter => chapter.Text.Contains("Signal received.", StringComparison.Ordinal));
    }

    [Fact]
    public void UsesLegacyMetadataWhenThePreferredKf8ValueDecodesToWhitespace() {
        byte[] bytes = Fixture("harbour-combo.mobi");
        int author = FindExthItem(bytes, RecordOffset(bytes, 11), 100);
        Assert.Equal(18, BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(author + 4)));
        Encoding.ASCII.GetBytes("&#x0009;  ").CopyTo(bytes, author + 8);
        using var source = new MemoryStream(bytes, writable: false);

        EpubDocument book = ReadPublication(source, "harbour-combo.mobi");

        Assert.Equal("Ada Reader", book.Creator);
    }

    [Fact]
    public void ResolvesEscapedLocalHtmlFragments() {
        using var source = new MemoryStream(BuildLegacyMobi("""
            <html><body><a name="café"></a><a href="#caf%C3%A9">Go to the note</a><p id="café">Note text</p></body></html>
            """), writable: false);

        EpubChapter chapter = Assert.Single(ReadPublication(source, "fragment.mobi").Chapters);

        Assert.NotNull(chapter.Html);
        XElement html = XElement.Parse(chapter.Html);
        XNamespace xhtml = "http://www.w3.org/1999/xhtml";
        string id = Assert.IsType<string>((string?)Assert.Single(html.Descendants(xhtml + "p")).Attribute("id"));
        XElement link = Assert.Single(html.Descendants(xhtml + "a"), anchor => anchor.Attribute("href") != null);
        Assert.Equal("#" + id, (string?)link.Attribute("href"));
    }

    [Fact]
    public void RejectsHtmlThatExceedsTheParserElementBudget() {
        string html = "<html><body>" + string.Concat(Enumerable.Repeat("<br>", 155_000)) + "</body></html>";
        using var stream = new MemoryStream(BuildLegacyMobi(html), writable: false);
        var error = Assert.Throws<InvalidDataException>(() => MobiReader.Import(stream));
        Assert.Contains("aggregate token or element limit", error.Message, StringComparison.Ordinal);
    }

    private static byte[] BuildLegacyMobi(string html) {
        byte[] existing = Fixture("harbour.mobi");
        int headerStart = RecordOffset(existing, 0);
        int headerEnd = RecordOffset(existing, 1);
        byte[] header = existing.AsSpan(headerStart, headerEnd - headerStart).ToArray();
        byte[] text = Encoding.UTF8.GetBytes(html);
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
        BinaryPrimitives.WriteUInt32BigEndian(header.AsSpan(28), 65001);
        BinaryPrimitives.WriteInt32BigEndian(header.AsSpan(4), text.Length);
        BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(8), (ushort)records);
        BinaryPrimitives.WriteUInt32BigEndian(header.AsSpan(108), uint.MaxValue);
        BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(242), 0);
        header.CopyTo(data, newHeader);
        text.CopyTo(data, newHeader + header.Length);
        return data;
    }

    private static int RecordOffset(byte[] bytes, int record) =>
        BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(78 + 8 * record));

    private static int FindExthItem(byte[] bytes, int header, int id) {
        int exth = header + 16 + BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(header + 20));
        int count = BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(exth + 8));
        int offset = exth + 12;
        for (int item = 0; item < count; item++) {
            if (BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(offset)) == id) return offset;
            offset += BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(offset + 4));
        }
        throw new InvalidDataException("The fixture is missing the expected EXTH record.");
    }

    private static EpubDocument ReadPublication(Stream source, string name) {
        using var epub = new MemoryStream();
        MobiReader.Import(source, name).Value.Save(epub);
        epub.Position = 0;
        return EpubDocument.Load(epub, new EpubLoadOptions { IncludeRawHtml = true, IncludeResourceData = true });
    }

    private static byte[] Fixture(string name) => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", name));
}
