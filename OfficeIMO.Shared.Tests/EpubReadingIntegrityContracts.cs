using OfficeIMO.Epub;
using System.IO.Compression;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Xunit;
using static OfficeIMO.Shared.Tests.EpubIntegrityFixtures;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubReadingIntegrityContractTests {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Load_PreservesRepeatedAndEmptyReadingPositions(bool spineOrder) {
        byte[] package = Package(new[] { ("c", "chapter.xhtml", "application/xhtml+xml", "") },
            "<itemref idref='c'/><itemref idref='c'/>", new[] { ("chapter.xhtml", Xhtml("")) });
        EpubDocument book = Read(package, new EpubReadOptions { PreferSpineOrder = spineOrder });
        Assert.Equal(new int?[] { 1, 2 }, book.Chapters.Select(chapter => chapter.SpineIndex));
        Assert.All(book.Chapters, chapter => Assert.Equal("", chapter.Text));
        Assert.Equal(2, book.ReadSummary.RequestedChapterCount);
        Assert.True(book.ReadSummary.IsComplete);
        Assert.Single(book.Resources, resource => resource.Id == "c");
    }

    [Theory]
    [InlineData("<p>Hel<em>lo</em>, world<strong>!</strong></p>", "Hello, world!")]
    [InlineData("<p><em>Hello</em> <strong>world</strong></p>", "Hello world")]
    [InlineData("<p>One</p><div>Two<br/>Three</div><table><tr><td>Four</td><td>Five</td></tr></table>", "One Two Three Four Five")]
    [InlineData("<p>A<script><nested>hidden</nested></script><style>hidden</style>B</p>", "AB")]
    public void Load_PreservesInlineContinuityAndBlockBoundaries(string markup, string expected) {
        Assert.Equal(expected, Assert.Single(Read(OneChapter(markup)).Chapters).Text);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Load_SelectsSpinePolicyBeforeOrdering(bool spineOrder) {
        byte[] package = Package(new[] {
            ("z", "z.xhtml", "application/xhtml+xml", ""), ("a", "a.xhtml", "application/xhtml+xml", ""),
            ("n", "notes.xhtml", "application/xhtml+xml", "") },
            "<itemref idref='z'/><itemref idref='a'/><itemref idref='n' linear='no'/>",
            new[] { ("z.xhtml", Xhtml("<p>Z</p>")), ("a.xhtml", Xhtml("<p>A</p>")), ("notes.xhtml", Xhtml("<p>Notes</p>")) });
        EpubDocument book = Read(package, new EpubReadOptions {
            PreferSpineOrder = spineOrder, IncludeNonLinearSpineItems = false
        });
        Assert.Equal(spineOrder ? new[] { "EPUB/z.xhtml", "EPUB/a.xhtml" } : new[] { "EPUB/a.xhtml", "EPUB/z.xhtml" },
            book.Chapters.Select(chapter => chapter.Path));
        Assert.Equal(2, book.ReadSummary.RequestedChapterCount);
        Assert.True(book.ReadSummary.IsComplete);
        Assert.DoesNotContain(book.Diagnostics, diagnostic => diagnostic.Code == "epub.chapter.fallback-scan");
    }

    [Fact]
    public void Load_DoesNotRecoverExcludedNonLinearSpineItems() {
        byte[] package = Package(new[] { ("c", "chapter.xhtml", "application/xhtml+xml", "") },
            "<itemref idref='c' linear='no'/>", new[] { ("chapter.xhtml", Xhtml("<p>Notes</p>")) });
        EpubDocument book = Read(package, new EpubReadOptions { IncludeNonLinearSpineItems = false });
        Assert.Empty(book.Chapters);
        Assert.Equal(0, book.ReadSummary.RequestedChapterCount);
        Assert.True(book.ReadSummary.IsComplete);
    }

    [Theory]
    [InlineData("unknown", "epub.spine.manifest-id-missing")]
    [InlineData("binary", "epub.spine.unsupported-media-type")]
    [InlineData("missing", "epub.spine.resource-missing")]
    public void Load_ReportsUnextractableSpineWithoutReplacingItWithNavigation(string id, string code) {
        byte[] package = Package(new[] {
            ("binary", "data.bin", "application/octet-stream", ""), ("missing", "missing.xhtml", "application/xhtml+xml", "") },
            "<itemref idref='" + id + "'/>", new[] { ("data.bin", "binary") });
        EpubDocument book = Read(package);
        Assert.Empty(book.Chapters);
        Assert.Contains(book.Diagnostics, diagnostic => diagnostic.Code == code);
        Assert.Equal(1, book.ReadSummary.SkippedChapterCount);
        Assert.False(book.ReadSummary.UsedFallbackScan);
        Assert.False(book.ReadSummary.IsComplete);
    }

    [Fact]
    public void Load_RecoveryScanReportsUnknownCompleteness() {
        byte[] package = Package(new[] { ("c", "chapter.xhtml", "application/xhtml+xml", "") },
            "", new[] { ("chapter.xhtml", Xhtml("<p>Recovered</p>")) });
        EpubDocument book = Read(package);
        Assert.Equal("Recovered", Assert.Single(book.Chapters).Text);
        Assert.True(book.ReadSummary.UsedFallbackScan);
        Assert.False(book.ReadSummary.IsComplete);
        Assert.Contains(book.Diagnostics, diagnostic => diagnostic.Code == "epub.chapter.fallback-scan");
    }

    [Theory]
    [InlineData("count")]
    [InlineData("text")]
    [InlineData("size")]
    public async Task LoadAndLoadAsync_ReportPartialChapterCounts(string limit) {
        byte[] package = Package(new[] { ("a", "a.xhtml", "application/xhtml+xml", ""), ("b", "b.xhtml", "application/xhtml+xml", "") },
            "<itemref idref='a'/><itemref idref='b'/>", new[] { ("a.xhtml", Xhtml("<p>One</p>")), ("b.xhtml", Xhtml("<p>Long second</p>")) });
        var options = new EpubReadOptions {
            MaxChapters = limit == "count" ? 1 : 500,
            MaxTotalTextCharacters = limit == "text" ? 3 : 1000,
            MaxChapterBytes = limit == "size" ? Encoding.UTF8.GetByteCount(Xhtml("<p>One</p>")) : 4096
        };
        EpubDocument[] books = { Read(package, options), await EpubDocument.LoadAsync(new MemoryStream(package), options) };
        foreach (EpubDocument book in books) {
            Assert.Single(book.Chapters);
            Assert.Equal(2, book.ReadSummary.RequestedChapterCount);
            Assert.Equal(1, book.ReadSummary.ExtractedChapterCount);
            Assert.Equal(1, book.ReadSummary.SkippedChapterCount);
            Assert.False(book.ReadSummary.IsComplete);
            Assert.Contains(book.Diagnostics, diagnostic => diagnostic.Code == "epub.chapter." +
                (limit == "count" ? "count-limit" : limit == "text" ? "text-total-limit" : "size-limit"));
        }
    }

    [Fact]
    public void Load_RejectsInvalidUtf8WithoutReplacementText() {
        byte[] markup = Encoding.UTF8.GetBytes(Xhtml("<p>Before?After</p>"));
        markup[Array.IndexOf(markup, (byte)'?')] = 0xff;
        EpubDocument book = Read(ReplaceEntry(OneChapter(""), "EPUB/chapter.xhtml", markup));
        Assert.Empty(book.Chapters);
        Assert.Contains(book.Diagnostics, diagnostic => diagnostic.Code == "epub.chapter.invalid-encoding");
        Assert.False(book.ReadSummary.IsComplete);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void Load_AcceptsUtf16ChapterBytes(bool bigEndian, bool bom) {
        var encoding = new UnicodeEncoding(bigEndian, bom, true);
        string content = "<?xml version='1.0' encoding='UTF-16'?>" + Xhtml("<p>Zażółć 世界</p>");
        byte[] bytes = encoding.GetPreamble().Concat(encoding.GetBytes(content)).ToArray();
        EpubDocument book = Read(ReplaceEntry(OneChapter(""), "EPUB/chapter.xhtml", bytes));
        Assert.Equal("Zażółć 世界", Assert.Single(book.Chapters).Text);
        Assert.True(book.ReadSummary.IsComplete);
    }

    [Theory]
    [InlineData("<x:base href='https://example.invalid/'/>")]
    [InlineData("<base x:href='https://example.invalid/'/>")]
    public void Load_IgnoresForeignNamespaceBaseDeclarations(string declaration) {
        string markup = "<html xmlns='http://www.w3.org/1999/xhtml' xmlns:x='urn:extension'><head>" +
            declaration + "</head><body><p>Body</p></body></html>";
        EpubDocument book = Read(ReplaceEntry(OneChapter(""), "EPUB/chapter.xhtml", Encoding.UTF8.GetBytes(markup)));
        Assert.Null(Assert.Single(book.Chapters).BaseHref);
    }

    [Theory]
    [InlineData("r: http://www.idpf.org/vocab/rendition/#", "r:layout", "r:layout-reflowable", true)]
    [InlineData("r: urn:extension", "r:layout", "", false)]
    public void Load_ResolvesRenditionVocabularyForPackageAndSpine(string prefix, string property, string spineProperty, bool fixedLayout) {
        EpubDocument book = Read(OneChapter("<p>Body</p>",
            "<meta property='" + property + "'>pre-paginated</meta>",
            "prefix='" + prefix + "'", spineProperty));
        Assert.Equal(fixedLayout, book.IsFixedLayout);
        if (fixedLayout) Assert.Equal(EpubRenditionLayout.Reflowable, Assert.Single(book.Chapters).RenditionLayout);
    }

    [Fact]
    public void Load_PrefersTheDeclaredSpineNcx() {
        string Ncx(string label) => "<ncx xmlns='http://www.daisy.org/z3986/2005/ncx/'><navMap><navPoint id='p'>" +
            "<navLabel><text>" + label + "</text></navLabel><content src='chapter.xhtml'/></navPoint></navMap></ncx>";
        EpubDocument book = Read(Package(new[] {
            ("c", "chapter.xhtml", "application/xhtml+xml", ""), ("other", "other.ncx", "application/x-dtbncx+xml", ""),
            ("chosen", "chosen.ncx", "application/x-dtbncx+xml", "") },
            "<itemref idref='c'/>", new[] { ("chapter.xhtml", Xhtml("<p>Body</p>")), ("other.ncx", Ncx("Wrong")), ("chosen.ncx", Ncx("Selected")) },
            version: "2.0", spineAttributes: "toc='chosen'"));
        Assert.Equal("Selected", Assert.Single(book.TableOfContents).Label);
    }

    [Fact]
    public void Load_InvalidSpineTocDoesNotChooseAnUnrelatedNcx() {
        EpubDocument book = Read(Package(new[] {
            ("c", "chapter.xhtml", "application/xhtml+xml", ""), ("other", "other.ncx", "application/x-dtbncx+xml", "") },
            "<itemref idref='c'/>", new[] { ("chapter.xhtml", Xhtml("<p>Body</p>")), ("other.ncx", "<ncx xmlns='http://www.daisy.org/z3986/2005/ncx/'/>") },
            version: "2.0", spineAttributes: "toc='missing'"));
        Assert.Empty(book.TableOfContents);
        Assert.Contains(book.Diagnostics, diagnostic => diagnostic.Code == "epub.ncx.spine-toc-invalid");
    }

    [Fact]
    public void Load_NavigationRecognizesElementAndAttributeNamespaces() {
        const string nav = "<html xmlns='http://www.w3.org/1999/xhtml' xmlns:x='urn:extension' xmlns:e='http://www.idpf.org/2007/ops'>" +
            "<head><x:base href='https://example.invalid/'/></head><body>" +
            "<x:nav e:type='toc'><ol><li><a href='wrong.xhtml'>Wrong element</a></li></ol></x:nav>" +
            "<nav x:type='toc'><ol><li><a href='wrong.xhtml'>Wrong attribute</a></li></ol></nav>" +
            "<nav e:type='toc'><ol><li><a x:href='wrong.xhtml' href='chapter.xhtml'>Correct</a></li></ol></nav></body></html>";
        EpubDocument book = Read(Package(new[] { ("c", "chapter.xhtml", "application/xhtml+xml", "") },
            "<itemref idref='c'/>", new[] { ("chapter.xhtml", Xhtml("<p>Body</p>")) }, nav: nav));
        Assert.Equal("EPUB/chapter.xhtml", Assert.Single(book.TableOfContents).Target);
        Assert.Equal("Correct", Assert.Single(book.Chapters).Title);
    }

    [Fact]
    public void Load_LayoutMetadataIgnoresForeignElementsAndAttributes() {
        EpubDocument book = Read(OneChapter("<p>Body</p>",
            "<x:meta xmlns:x='urn:extension' property='rendition:layout'>pre-paginated</x:meta>" +
            "<meta xmlns:x='urn:extension' x:property='rendition:layout'>pre-paginated</meta>"));
        Assert.False(book.IsFixedLayout);
        Assert.False(Assert.Single(book.Chapters).IsFixedLayout);
    }

    [Fact]
    public void Load_HtmlExtensionDoesNotOverrideUnsupportedDeclaredMediaType() {
        EpubDocument book = Read(Package(new[] { ("c", "chapter.xhtml", "application/octet-stream", "") },
            "<itemref idref='c'/>", new[] { ("chapter.xhtml", Xhtml("<p>Not a declared content document</p>")) }));
        Assert.Empty(book.Chapters);
        Assert.Contains(book.Diagnostics, diagnostic => diagnostic.Code == "epub.spine.unsupported-media-type");
        Assert.False(book.ReadSummary.IsComplete);
    }

    [Theory]
    [InlineData("pkg-spine-order-svg")]
    [InlineData("pkg-spine-duplicate-item-rendering")]
    public void Load_IndependentW3cPublicationsRetainFourSpinePositions(string name) {
        string source = Path.Combine(AppContext.BaseDirectory, "w3c", name);
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, true)) {
            foreach (string file in Directory.GetFiles(source, "*", SearchOption.AllDirectories)
                .OrderBy(file => Path.GetFileName(file) == "mimetype" ? 0 : 1)) {
                string path = file.Substring(source.Length + 1).Replace('\\', '/');
                using var destination = archive.CreateEntry(path,
                    path == "mimetype" ? CompressionLevel.NoCompression : CompressionLevel.Optimal).Open();
                using var input = File.OpenRead(file);
                input.CopyTo(destination);
            }
        }
        EpubDocument book = Read(stream.ToArray(), new EpubReadOptions { IncludeRawHtml = true });
        Assert.Equal(new int?[] { 1, 2, 3, 4 }, book.Chapters.Select(chapter => chapter.SpineIndex));
        Assert.True(book.ReadSummary.IsComplete);
        Assert.All(book.Chapters, chapter => Assert.NotNull(chapter.Html));
        Assert.DoesNotContain(book.Chapters, chapter => chapter.Path.EndsWith("nav.xhtml", StringComparison.Ordinal));
    }

    [Fact]
    public void Load_SvgPreservesTextStructureAndRawMarkup() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg'><title>Page</title><rect width='100' height='50'/>" +
            "<text x='10' y='10'>Hel<tspan>lo</tspan>!</text><text x='10' y='30'>Second</text></svg>";
        EpubDocument book = Read(Package(new[] { ("s", "page.svg", "image/svg+xml", "") },
            "<itemref idref='s'/>", new[] { ("page.svg", svg) }), new EpubReadOptions { IncludeRawHtml = true });
        EpubChapter chapter = Assert.Single(book.Chapters);
        Assert.Equal("Hello! Second", chapter.Text);
        Assert.True(chapter.HasStructuredContent);
        Assert.Equal(svg, chapter.Html);
        Assert.True(book.ReadSummary.IsComplete);
    }

    [Fact]
    public void Load_SvgEmbeddedXhtmlBodyDoesNotResetEarlierText() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg'><text>First</text>" +
            "<foreignObject><body xmlns='http://www.w3.org/1999/xhtml'/></foreignObject><text>Second</text></svg>";
        EpubDocument book = Read(Package(new[] { ("s", "page.svg", "image/svg+xml", "") },
            "<itemref idref='s'/>", new[] { ("page.svg", svg) }));
        Assert.Equal("First Second", Assert.Single(book.Chapters).Text);
    }

    [Fact]
    public void Load_CancellationDoesNotConsumeOrDisposeCallerStream() {
        using var stream = new MemoryStream(OneChapter("<p>Body</p>"));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => EpubDocument.Load(stream, null, cancellation.Token));
        Assert.Equal(0, stream.Position);
        Assert.True(stream.CanRead);
    }

    private static EpubDocument Read(byte[] bytes, EpubReadOptions? options = null) =>
        EpubDocument.Load(new MemoryStream(bytes), options);
}
