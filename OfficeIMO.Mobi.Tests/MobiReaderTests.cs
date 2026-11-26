using System.Buffers.Binary;
using System.IO.Compression;
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
    public void EmptyKindleXhtmlElementsKeepTextAndLinksAtTheirOriginalPositions() {
        const string fixture = "harbour-empty-style.azw3";
        using var source = new MemoryStream(Fixture(fixture), writable: false);
        var imported = MobiReader.Import(source, fixture);
        Assert.DoesNotContain(imported.Report.Diagnostics, item =>
            item.Code == "MOBI_STYLE_OMITTED" || item.Code == "MOBI_ELEMENT_OMITTED" && item.Location == "style");
        using var artifact = new MemoryStream();
        imported.Value.Save(artifact);
        artifact.Position = 0;
        EpubDocument book = EpubDocument.Load(artifact, new EpubLoadOptions { IncludeRawHtml = true });
        EpubChapter opening = Assert.Single(book.Chapters, chapter => chapter.Text.Contains("A small independent ebook"));
        var ending = Assert.Single(book.TableOfContents[0].Children);
        Assert.NotNull(ending.Fragment);
        Assert.Contains("href=\"" + Path.GetFileName(ending.Target) + "#" + ending.Fragment + "\">Read the ending", opening.Html);

        XElement html = XElement.Parse(opening.Html!);
        XNamespace xhtml = "http://www.w3.org/1999/xhtml";
        string svg = FindLinkTarget("SVG marker");
        string local = FindLinkTarget("HTML marker");
        Assert.NotEqual(svg, local);
        XElement paragraph = Assert.Single(html.Descendants(xhtml + "p"), node => node.Value.Contains("Prølogue target phrase"));
        Assert.Empty(Assert.Single(paragraph.Elements(xhtml + "em")).Nodes());
        XElement jump = Assert.Single(paragraph.Elements(xhtml + "a"), link => link.Value == "Jump");
        string path = Path.GetFileName(opening.Path) + "#";
        string href = Assert.IsType<string>((string?)jump.Attribute("href"));
        Assert.StartsWith(path, href);
        XElement marker = Assert.Single(paragraph.Elements(xhtml + "span"),
            node => (string?)node.Attribute("id") == href.Substring(path.Length));
        Assert.Equal("Prølogue ", Assert.IsType<XText>(marker.PreviousNode).Value);
        Assert.Equal("target phrase", Assert.IsType<XText>(marker.NextNode).Value);

        string FindLinkTarget(string label) {
            XElement link = Assert.Single(html.Descendants(xhtml + "a"), element => element.Value == label);
            string target = Assert.IsType<string>((string?)link.Attribute("href"));
            Assert.StartsWith("#", target);
            Assert.Single(html.Descendants(), element => (string?)element.Attribute("id") == target.Substring(1));
            return target;
        }
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
    public void RefusesAnInvalidUtf8NameInTheKindleSkeletonIndex() {
        byte[] bytes = Fixture("harbour.azw3");
        int header = RecordOffset(bytes, 0);
        int skeleton = BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(header + 252));
        int index = RecordOffset(bytes, skeleton + 1);
        int directory = BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(index + 20));
        int row = index + BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(index + directory + 4));
        Assert.True(bytes[row] > 0);
        bytes[row + 1] = 0xff;
        using var source = new MemoryStream(bytes, writable: false);

        var error = Assert.Throws<InvalidDataException>(() => MobiReader.Import(source, "harbour.azw3"));
        Assert.Contains("index entry name", error.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(16, true)]
    [InlineData(44, false)]
    public void ValidatesTheDeclaredKindleFlowTableOffset(int offset, bool readable) {
        byte[] bytes = Fixture("harbour.azw3");
        int header = RecordOffset(bytes, 0);
        int index = BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(header + 192));
        int flow = RecordOffset(bytes, index);
        if (readable) {
            int insertion = flow + 12;
            int padding = offset - 12;
            var padded = new byte[bytes.Length + padding];
            bytes.AsSpan(0, insertion).CopyTo(padded);
            bytes.AsSpan(insertion).CopyTo(padded.AsSpan(insertion + padding));
            int records = BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(76));
            for (int record = index + 1; record < records; record++)
                BinaryPrimitives.WriteInt32BigEndian(padded.AsSpan(78 + 8 * record), RecordOffset(bytes, record) + padding);
            bytes = padded;
        }
        BinaryPrimitives.WriteInt32BigEndian(bytes.AsSpan(flow + 4), offset);
        using var source = new MemoryStream(bytes, writable: false);

        if (readable) {
            EpubDocument book = ReadPublication(source, "padded-flow.azw3");
            Assert.Contains(book.Chapters, chapter => chapter.Text.Contains("Signal received."));
        } else {
            var error = Assert.Throws<InvalidDataException>(() => MobiReader.Import(source, "invalid-flow.azw3"));
            Assert.Contains("FDST", error.Message, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void ReadsTextWhenTheResourceRangeStartsAtTheEnd() {
        byte[] bytes = BuildLegacyMobi("<html><body><p>Text without images.</p></body></html>");
        int header = RecordOffset(bytes, 0);
        uint records = BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(76));
        BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(header + 108), records);
        using var source = new MemoryStream(bytes, writable: false);

        EpubDocument book = ReadPublication(source, "text-only.mobi");

        Assert.Contains(book.Chapters, chapter => chapter.Text.Contains("Text without images."));
        Assert.DoesNotContain(book.Resources, item => item.MediaType?.StartsWith("image/", StringComparison.Ordinal) == true);
    }

    [Fact]
    public void InvalidOptionalTitleRangeKeepsTheExthTitle() {
        byte[] bytes = Fixture("harbour.mobi");
        int header = RecordOffset(bytes, 0);
        int next = RecordOffset(bytes, 1);
        BinaryPrimitives.WriteInt32BigEndian(bytes.AsSpan(header + 84), next - header - 2);
        using var source = new MemoryStream(bytes, writable: false);

        var imported = MobiReader.Import(source, "invalid-title-range.mobi");

        Assert.Equal("Harbour Signal", imported.Value.Title);
        Assert.Contains(imported.Report.Diagnostics, item => item.Code == "MOBI_HEADER_TITLE_RANGE_OMITTED");
    }

    [Fact]
    public void ReadsSingleFlowKf8WithAnUnusedFlowPointer() {
        byte[] bytes = Fixture("harbour.azw3");
        int header = RecordOffset(bytes, 0);
        BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(header + 192), 0);
        BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(header + 196), 1);
        using var source = new MemoryStream(bytes, writable: false);

        EpubDocument book = ReadPublication(source, "harbour.azw3");

        Assert.Contains(book.Chapters, chapter => chapter.Text.IndexOf("Signal received.", StringComparison.Ordinal) >= 0);
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

        Assert.Contains(book.Chapters, chapter => chapter.Text.IndexOf("Signal received.", StringComparison.Ordinal) >= 0);
        Assert.DoesNotContain(book.Resources, item => item.Properties == "cover-image");
    }

    [Fact]
    public void RefusesTheKf8BoundaryInsideLegacyText() {
        byte[] bytes = Fixture("harbour-combo.mobi");
        int header = RecordOffset(bytes, 0);
        int legacyText = BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(header + 8));
        int item = FindExthItem(bytes, header, 121);
        BinaryPrimitives.WriteInt32BigEndian(bytes.AsSpan(item + 8), legacyText + 1);
        "BOUNDARY"u8.CopyTo(bytes.AsSpan(RecordOffset(bytes, legacyText)));
        using var source = new MemoryStream(bytes, writable: false);

        var error = Assert.Throws<InvalidDataException>(() => MobiReader.Import(source, "harbour-combo.mobi"));
        Assert.Contains("valid KF8 boundary", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void WritesKindlePageProgressionToTheEpubSpine() {
        byte[] bytes = Fixture("harbour.azw3");
        int item = FindExthItem(bytes, RecordOffset(bytes, 0), 501);
        BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(item), 527);
        "rtl "u8.CopyTo(bytes.AsSpan(item + 8));
        using var source = new MemoryStream(bytes, writable: false);
        using var artifact = new MemoryStream();
        MobiReader.Import(source, "harbour.azw3").Value.Save(artifact);
        artifact.Position = 0;

        using var package = new ZipArchive(artifact);
        using Stream opf = Assert.Single(package.Entries, entry => entry.Name.EndsWith(".opf", StringComparison.Ordinal)).Open();
        XNamespace opfNamespace = "http://www.idpf.org/2007/opf";
        XElement spine = Assert.Single(XDocument.Load(opf).Descendants(opfNamespace + "spine"));
        Assert.Equal("rtl", (string?)spine.Attribute("page-progression-direction"));
    }

    [Theory]
    [InlineData(103)]
    [InlineData(104)]
    public void OmitsOptionalMetadataThatDecodesToWhitespace(int tag) {
        byte[] bytes = Fixture("harbour.azw3");
        int item = FindExthItem(bytes, RecordOffset(bytes, 0), 205);
        Assert.Equal(12, BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(item + 4)));
        BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(item), (uint)tag);
        Encoding.ASCII.GetBytes(tag == 103 ? "&#9;" : "    ").CopyTo(bytes, item + 8);
        using var source = new MemoryStream(bytes, writable: false);

        EpubDocument book = ReadPublication(source, "harbour.azw3");

        Assert.Equal("Ada Reader", book.Creator);
        Assert.DoesNotContain(book.Metadata, value => value.Name == "description");
        Assert.Equal("c4b5bb5a-5519-44a6-bada-106df8645d5a", book.Identifier);
        Assert.Contains(book.Chapters, chapter => chapter.Text.IndexOf("Signal received.", StringComparison.Ordinal) >= 0);
    }

    [Fact]
    public void UsesLegacyMetadataWhenThePreferredKf8ValueDecodesToWhitespace() {
        byte[] bytes = Fixture("harbour-combo.mobi");
        int fallback = FindExthItem(bytes, RecordOffset(bytes, 0), 100);
        int preferred = FindExthItem(bytes, RecordOffset(bytes, 11), 100);
        Assert.Equal(18, BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(preferred + 4)));
        BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(fallback), 103);
        BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(preferred), 103);
        Encoding.ASCII.GetBytes("&#x0009;  ").CopyTo(bytes, preferred + 8);
        using var source = new MemoryStream(bytes, writable: false);

        EpubDocument book = ReadPublication(source, "harbour-combo.mobi");

        Assert.Contains(book.Metadata, item => item.Name == "description" && item.Value == "Ada Reader");
    }

    [Fact]
    public void PreservesLiteralTitleEntitiesInPlainMetadata() {
        byte[] bytes = Fixture("harbour.azw3");
        int title = FindExthItem(bytes, RecordOffset(bytes, 0), 503);
        Assert.Equal(22, BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(title + 4)));
        Encoding.ASCII.GetBytes("How &lt; Works").CopyTo(bytes, title + 8);
        using var source = new MemoryStream(bytes, writable: false);

        Assert.Equal("How &lt; Works", ReadPublication(source, "harbour.azw3").Title);
    }

    [Theory]
    [InlineData("<html><body><p>Legacy story.</p></body></html>", "Legacy story.", null)]
    [InlineData("First paragraph\r Second paragraph", "First paragraph", "Second paragraph")]
    public void ReadsLegacyMobiWithATextDescriptorOnly(string input, string first, string? second) {
        using var source = new MemoryStream(BuildLegacyMobi(input, includeMobiHeader: false), writable: false);

        EpubChapter chapter = Assert.Single(ReadPublication(source, "legacy.mobi").Chapters);
        Assert.NotNull(chapter.Html);
        XNamespace xhtml = "http://www.w3.org/1999/xhtml";
        string[] expected = second == null ? new[] { first } : new[] { first, second };
        Assert.Equal(expected, XElement.Parse(chapter.Html).Descendants(xhtml + "p").Select(paragraph => paragraph.Value.Trim()));
    }

    [Fact]
    public void Utf8BookReadsThePalmNameWhenNativeTitleFieldsAreEmpty() {
        byte[] bytes = Fixture("harbour.mobi");
        int header = RecordOffset(bytes, 0);
        int title = FindExthItem(bytes, header, 503);
        int titleSize = BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(title + 4)) - 8;
        Array.Clear(bytes, title + 8, titleSize);
        BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(header + 88), 0);
        Array.Clear(bytes, 0, 32);
        Encoding.ASCII.GetBytes("Caf").CopyTo(bytes, 0);
        bytes[3] = 0xe9;
        using var source = new MemoryStream(bytes, writable: false);

        EpubDocument book = ReadPublication(source, "untitled.mobi");

        Assert.Equal("Café", book.Title);
        Assert.Contains(book.Chapters, chapter => chapter.Text.Contains("Signal received."));
    }

    [Fact]
    public void InvalidOptionalUtf8MetadataKeepsTheBookAndReportsTheOmittedValues() {
        byte[] bytes = Fixture("harbour.mobi");
        int header = RecordOffset(bytes, 0);
        int fullName = header + BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(header + 84));
        bytes[fullName] = 0xff;
        bytes[FindExthItem(bytes, header, 503) + 8] = 0xff;
        bytes[FindExthItem(bytes, header, 100) + 8] = 0xff;
        Array.Clear(bytes, 0, 32);
        Encoding.ASCII.GetBytes("Recovered title").CopyTo(bytes, 0);
        using var input = new MemoryStream(bytes, writable: false);

        var imported = MobiReader.Import(input, "damaged-metadata.mobi");
        using var artifact = new MemoryStream();
        imported.Value.Save(artifact);
        artifact.Position = 0;
        EpubDocument book = EpubDocument.Load(artifact);

        Assert.Equal("Recovered title", book.Title);
        Assert.Contains(book.Chapters, chapter => chapter.Text.Contains("Signal received."));
        Assert.Contains(book.Metadata, item => item.Name == "publisher" && item.Value == "Signal Press");
        Assert.Contains(imported.Report.Diagnostics, item => item.Code == "MOBI_METADATA_ENCODING_OMITTED" && item.Location == "header title");
        Assert.Contains(imported.Report.Diagnostics, item => item.Code == "MOBI_METADATA_ENCODING_OMITTED" && item.Location == "EXTH 100");
    }

    [Fact]
    public void StandaloneKindleRecoversTheUniqueDeclaredRasterRangeWithoutAPointer() {
        byte[] bytes = Fixture("harbour.azw3");
        int header = RecordOffset(bytes, 0);
        BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(header + 108), uint.MaxValue);
        using var source = new MemoryStream(bytes, writable: false);

        var imported = MobiReader.Import(source, "no-raster-pointer.azw3");
        using var artifact = new MemoryStream();
        imported.Value.Save(artifact);
        artifact.Position = 0;
        EpubDocument book = EpubDocument.Load(artifact, new EpubLoadOptions { IncludeResourceData = true });

        Assert.Contains(book.Resources, item => item.Id == "image-3" && item.Data?.Length > 0);
        Assert.Contains(book.Resources, item => item.Properties == "cover-image" && item.Data?.Length > 0);
        Assert.Contains(imported.Report.Diagnostics, item => item.Code == "MOBI_RESOURCE_RANGE_INFERRED");
        Assert.DoesNotContain(imported.Report.Diagnostics, item => item.Code == "MOBI_IMAGE_UNAVAILABLE");
    }

    [Fact]
    public void ReportsWhenDistinctDiagnosticsExceedTheLimit() {
        string images = string.Concat(Enumerable.Range(1, 1001).Select(index => $"<img recindex=\"{index}\"/>"));
        using var source = new MemoryStream(BuildLegacyMobi("<html><body>" + images + "</body></html>"), writable: false);

        OfficeConversionReport report = MobiReader.Import(source, "missing-images.mobi").Report;

        Assert.Equal(1_000, report.Diagnostics.Count);
        Assert.Equal("MOBI_DIAGNOSTICS_TRUNCATED", report.Diagnostics[report.Diagnostics.Count - 1].Code);
        Assert.True(report.HasLoss);
    }

    [Fact]
    public void ProvidesATitleWhenPalmDocHasNoNameOrFilenameStem() {
        byte[] bytes = Fixture("harbour.pdb");
        Array.Clear(bytes, 0, 32);
        using var source = new MemoryStream(bytes, writable: false);

        EpubDocument book = ReadPublication(source, ".pdb");
        Assert.Equal("Untitled", book.Title);
        Assert.Contains("Signal received.", Assert.Single(book.Chapters).Text);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ReadsNavigationFromTheTagTableAfterPaddedHeaders(bool usePointer) {
        byte[] bytes = Fixture("harbour.mobi");
        int header = RecordOffset(bytes, 0);
        int index = RecordOffset(bytes, BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(header + 244)));
        Assert.Equal(192, BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(index + 180)));
        BinaryPrimitives.WriteInt32BigEndian(bytes.AsSpan(index + 4), 188);
        if (!usePointer) BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(index + 180), uint.MaxValue);
        using var source = new MemoryStream(bytes, writable: false);

        EpubDocument book = ReadPublication(source, "harbour.mobi");
        Assert.Equal("Opening Signal", book.TableOfContents[0].Label);
        Assert.Equal("Closing Signal", book.TableOfContents[1].Label);
    }

    [Fact]
    public void ResolvesEscapedLocalHtmlFragments() {
        using var source = new MemoryStream(BuildLegacyMobi("""
            <html><body><a name="café"></a><a href="#caf%C3%A9">Go to the note</a><p id="café">Note text</p><a href="#">Back to top</a></body></html>
            """), writable: false);

        EpubChapter chapter = Assert.Single(ReadPublication(source, "fragment.mobi").Chapters);

        Assert.NotNull(chapter.Html);
        XElement html = XElement.Parse(chapter.Html);
        XNamespace xhtml = "http://www.w3.org/1999/xhtml";
        string id = Assert.IsType<string>((string?)Assert.Single(html.Descendants(xhtml + "p")).Attribute("id"));
        XElement link = Assert.Single(html.Descendants(xhtml + "a"), anchor => anchor.Value == "Go to the note");
        Assert.Equal("#" + id, (string?)link.Attribute("href"));
        XElement chapterBody = Assert.Single(html.Descendants(xhtml + "div"));
        string topId = Assert.IsType<string>((string?)chapterBody.Attribute("id"));
        XElement top = Assert.Single(html.Descendants(xhtml + "a"), anchor => anchor.Value == "Back to top");
        Assert.Equal("#" + topId, (string?)top.Attribute("href"));
    }

    [Theory]
    [InlineData("", "fr")]
    [InlineData(" lang=\"de\"", "de")]
    public void RetainsRootAndBodyLinksAndInheritedLanguage(string bodyLanguage, string expectedLanguage) {
        using var source = new MemoryStream(BuildLegacyMobi($"""
            <html id="root" xml:lang="fr"><body id="chapter"{bodyLanguage}>
              <p xml:lang="es"><a href="#root">Root</a><a href="#chapter">Chapter</a></p>
            </body></html>
            """), writable: false);

        string? chapter = Assert.Single(ReadPublication(source, "root.mobi").Chapters).Html;
        Assert.NotNull(chapter);
        XElement html = XElement.Parse(chapter);
        XNamespace xhtml = "http://www.w3.org/1999/xhtml";
        XElement content = Assert.Single(html.Descendants(xhtml + "div"));
        XElement paragraph = Assert.Single(content.Elements(xhtml + "p"));
        Assert.Equal(expectedLanguage, (string?)content.Attribute("lang"));
        Assert.Equal(expectedLanguage, (string?)content.Attribute(XNamespace.Xml + "lang"));
        Assert.Equal("es", (string?)paragraph.Attribute("lang"));
        Assert.Equal("es", (string?)paragraph.Attribute(XNamespace.Xml + "lang"));
        string fragment = "#" + Assert.IsType<string>((string?)content.Attribute("id"));
        Assert.Collection(paragraph.Elements(xhtml + "a"),
            root => Assert.Equal(fragment, (string?)root.Attribute("href")),
            body => Assert.Equal(fragment, (string?)body.Attribute("href")));
    }

    [Fact]
    public void HtmlTopFallbackHonorsMatchingIdsFirst() {
        using var source = new MemoryStream(BuildLegacyMobi("""
            <html><body><a href="#%54OP">Back to document top</a><a href="#top">Named target</a><span id="top">Explicit target</span></body></html>
            """), writable: false);

        string? chapter = Assert.Single(ReadPublication(source, "top-link.mobi").Chapters).Html;
        Assert.NotNull(chapter);
        XElement html = XElement.Parse(chapter);
        XNamespace xhtml = "http://www.w3.org/1999/xhtml";
        XElement body = Assert.Single(html.Descendants(xhtml + "div"));
        XElement target = Assert.Single(html.Descendants(xhtml + "span"));
        XElement[] links = html.Descendants(xhtml + "a").ToArray();
        Assert.Equal("#" + (string?)body.Attribute("id"), (string?)links[0].Attribute("href"));
        Assert.Equal("#" + (string?)target.Attribute("id"), (string?)links[1].Attribute("href"));
    }

    [Fact]
    public void ZeroRowspanIsRetainedAndZeroColspanIsIgnored() {
        using var source = new MemoryStream(BuildLegacyMobi("""
            <html><body><table><tbody><tr><td rowspan="0" colspan="0">Shared label</td><td>First row</td></tr><tr><td>Second row</td></tr></tbody></table></body></html>
            """), writable: false);

        string? chapter = Assert.Single(ReadPublication(source, "rowspan.mobi").Chapters).Html;
        Assert.NotNull(chapter);
        XNamespace xhtml = "http://www.w3.org/1999/xhtml";
        XElement cell = Assert.Single(XElement.Parse(chapter).Descendants(xhtml + "td"), element => element.Value == "Shared label");
        Assert.Equal("0", (string?)cell.Attribute("rowspan"));
        Assert.Null(cell.Attribute("colspan"));
    }

    [Fact]
    public void PreservesContinuedAndReversedOrderedListNumbering() {
        using var source = new MemoryStream(BuildLegacyMobi("""
            <html><body><ol type="A" start="5" reversed><li>Step five</li><li value="10">Step ten</li></ol></body></html>
            """), writable: false);

        EpubChapter chapter = Assert.Single(ReadPublication(source, "list.mobi").Chapters);

        Assert.NotNull(chapter.Html);
        XNamespace xhtml = "http://www.w3.org/1999/xhtml";
        XElement list = Assert.Single(XElement.Parse(chapter.Html).Descendants(xhtml + "ol"));
        Assert.Equal("A", (string?)list.Attribute("type"));
        Assert.Equal("5", (string?)list.Attribute("start"));
        Assert.NotNull(list.Attribute("reversed"));
        Assert.Equal("10", (string?)Assert.Single(list.Elements(xhtml + "li"), element => element.Value == "Step ten").Attribute("value"));
    }

    [Fact]
    public void PreservesInheritedAndExplicitTextDirection() {
        using var source = new MemoryStream(BuildLegacyMobi("""
            <html dir="RTL"><body><p dir="ltr">ABC 123</p><p dir="auto">שלום 123 ABC</p></body></html>
            """), writable: false);

        EpubChapter chapter = Assert.Single(ReadPublication(source, "direction.mobi").Chapters);

        Assert.NotNull(chapter.Html);
        XNamespace xhtml = "http://www.w3.org/1999/xhtml";
        XElement content = Assert.Single(XElement.Parse(chapter.Html).Descendants(xhtml + "div"));
        Assert.Equal("rtl", (string?)content.Attribute("dir"));
        Assert.Equal(new[] { "ltr", "auto" }, content.Elements(xhtml + "p").Select(p => (string?)p.Attribute("dir")));
    }

    [Fact]
    public void LegacyPageBreaksPreserveFilePositionsInsideText() {
        string sourceHtml = "<html><body>" + string.Concat(Enumerable.Repeat("<p>Section</p><mbp:pagebreak/>", 80)) +
            """<p><a filepos="TARGET">Jump</a></p><p>Pré <em>aside</em> after this</p></body></html>""";
        int textOffset = Encoding.UTF8.GetByteCount(sourceHtml.Substring(0, sourceHtml.IndexOf("after this", StringComparison.Ordinal)));
        sourceHtml = sourceHtml.Replace("TARGET", textOffset.ToString("D6", System.Globalization.CultureInfo.InvariantCulture));
        using var source = new MemoryStream(BuildLegacyMobi(sourceHtml), writable: false);

        EpubChapter chapter = Assert.Single(ReadPublication(source, "positions.mobi").Chapters);

        Assert.NotNull(chapter.Html);
        XElement html = XElement.Parse(chapter.Html);
        XNamespace xhtml = "http://www.w3.org/1999/xhtml";
        Assert.Equal(80, html.Descendants(xhtml + "hr").Count());
        string? href = (string?)Assert.Single(html.Descendants(xhtml + "a")).Attribute("href");
        Assert.NotNull(href);
        string fragment = href.Substring(href.IndexOf('#') + 1);
        Assert.NotEqual(href, fragment);
        XElement anchor = Assert.Single(html.Descendants(), item => (string?)item.Attribute("id") == fragment);
        Assert.Equal(xhtml + "p", anchor.Parent!.Name);
        Assert.Equal("after this", Assert.IsType<XText>(anchor.NextNode).Value);
    }

    [Fact]
    public void AbsoluteKindleNavigationKeepsTheCorrectChapterWithAnApproximation() {
        byte[] bytes = Fixture("harbour.azw3");
        int header = RecordOffset(bytes, 0);
        int navigationRecord = BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(header + 244));
        int rows = RecordOffset(bytes, navigationRecord + 1);
        int directory = rows + BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(rows + 20));
        // The fixture stores tag 6 in the high bit of its single control byte.
        for (int index = 0; index < 2; index++) {
            int row = rows + BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(directory + 4 + index * 2));
            int control = row + 1 + bytes[row];
            Assert.NotEqual(0, bytes[control] & 0x80);
            bytes[control] &= 0x7f;
        }
        using var source = new MemoryStream(bytes, writable: false);

        var imported = MobiReader.Import(source, "absolute.azw3");
        using var artifact = new MemoryStream();
        imported.Value.Save(artifact);
        artifact.Position = 0;
        EpubDocument book = EpubDocument.Load(artifact, new EpubLoadOptions { IncludeRawHtml = true });

        var closing = Assert.Single(Assert.Single(book.TableOfContents).Children);
        Assert.Equal("Closing Signal", closing.Label);
        Assert.Contains(book.Chapters, chapter => chapter.Path == closing.Target && chapter.Text.Contains("Signal received."));
        Assert.Contains(imported.Report.Diagnostics, detail => detail.Code == "MOBI_TOC_TARGET_APPROXIMATED");
    }

    [Theory]
    [InlineData("Notes\0Old", "Notes")]
    [InlineData("\0Old", "memo")]
    public void PalmDocNameStopsAtTheFirstNul(string databaseName, string expected) {
        byte[] bytes = Fixture("harbour.pdb");
        Array.Clear(bytes, 0, 32);
        Encoding.ASCII.GetBytes(databaseName).CopyTo(bytes, 0);
        using var source = new MemoryStream(bytes, writable: false);

        Assert.Equal(expected, ReadPublication(source, "memo.pdb").Title);
    }

    [Fact]
    public void RejectsHtmlThatExceedsTheParserElementBudget() {
        string html = "<html><body>" + string.Concat(Enumerable.Repeat("<br>", 155_000)) + "</body></html>";
        using var stream = new MemoryStream(BuildLegacyMobi(html), writable: false);
        var error = Assert.Throws<InvalidDataException>(() => MobiReader.Import(stream));
        Assert.Contains("aggregate token or element limit", error.Message, StringComparison.Ordinal);
    }

    private static byte[] BuildLegacyMobi(string html, bool includeMobiHeader = true) {
        byte[] existing = Fixture("harbour.mobi");
        int headerStart = RecordOffset(existing, 0);
        int headerEnd = RecordOffset(existing, 1);
        byte[] header = existing.AsSpan(headerStart, includeMobiHeader ? headerEnd - headerStart : 16).ToArray();
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
        BinaryPrimitives.WriteInt32BigEndian(header.AsSpan(4), text.Length);
        BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(8), (ushort)records);
        if (includeMobiHeader) {
            BinaryPrimitives.WriteUInt32BigEndian(header.AsSpan(28), 65001);
            BinaryPrimitives.WriteUInt32BigEndian(header.AsSpan(108), uint.MaxValue);
            BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(242), 0);
        }
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
