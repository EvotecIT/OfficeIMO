using System.Security.Cryptography;
using System.Text;
using System.Threading;
using OfficeIMO.Epub;

namespace OfficeIMO.Chm.Tests;

public sealed class ArchiveTests {
    [Fact]
    public void IndependentCompilerArchiveMatchesEveryNativeDecoderEntry() {
        ChmDocument book = ChmDocument.Load(ChmFixture.NativePath);
        string[] manifest = File.ReadAllLines(Path.Combine(AppContext.BaseDirectory, "Fixtures", "svnbook-ms-htmlhelp.sha256.tsv"));
        Assert.Equal(manifest.Length, book.Entries.Count);
        using var sha = SHA256.Create();
        foreach (string row in manifest) {
            string[] columns = row.Split('\t'); ChmEntry entry = Assert.Single(book.Entries, item => item.Name == columns[2]);
            Assert.Equal(long.Parse(columns[0]), entry.Length);
            Assert.Equal(columns[1], BitConverter.ToString(sha.ComputeHash(entry.GetBytes())).Replace("-", "").ToLowerInvariant());
        }
        Assert.Equal("Version Control with Subversion", book.Title);
        Assert.Equal(181, book.Topics.Count); Assert.Equal(16, book.TableOfContents.Count); Assert.Equal(15, book.Index.Count);
        Assert.Equal("Table of Contents", book.TableOfContents[0].Name);
        Assert.Equal("/index.html", book.Topics[0].Path);
        Assert.Contains(book.TableOfContents, item => item.Name == "Preface" && item.Children.Count == 7);
        Assert.Contains(book.Index, item => item.Name.Contains("svn"));
    }

    [Theory]
    [InlineData(2U)] [InlineData(3U)]
    public void HtmlSitemapsRetainHierarchyMultipleTargetsAndSeeAlso(uint version) {
        ChmDocument book = ChmDocument.Load(ChmFixture.Book(version));
        Assert.Equal(version, book.Version); Assert.Equal("OfficeIMO help fixture", book.Title);
        ChmNavigationItem group = Assert.Single(book.TableOfContents); Assert.Equal("Guide", group.Name);
        Assert.Equal(2, group.Children.Count); Assert.Equal("/guide/details.html#part", group.Children[1].Links[0].Target);
        Assert.Equal(new[] { "/welcome.html", "/guide/details.html" }, book.Topics.Select(topic => topic.Path));
        Assert.Equal(2, book.Index[0].Links.Count); Assert.Equal("Details", book.Index[0].Links[1].Title);
        Assert.Equal("Example", Assert.Single(book.Index[1].SeeAlso));
    }

    [Fact]
    public async Task InputAndEntrySnapshotsPreserveCallerOwnership() {
        byte[] input = ChmFixture.Book(); using var stream = new MemoryStream(input); stream.Position = 7;
        ChmDocument book = await ChmDocument.LoadAsync(stream); Assert.Equal(7, stream.Position); Assert.True(stream.CanRead);
        byte[] bytes = book.Topics[0].Entry.GetBytes(); bytes[0] = 0;
        Assert.NotEqual(0, book.Topics[0].Entry.GetBytes()[0]);
        using Stream entry = book.Topics[0].Entry.OpenRead(); Assert.False(entry.CanWrite);
        input[0] = 0; Assert.Contains("Café", book.Topics[0].ReadHtml());
        var forward = new ForwardOnlyStream(ChmFixture.Book()); Assert.Equal(2, ChmDocument.Load(forward).Topics.Count); Assert.True(forward.CanRead);
    }

    [Fact]
    public void EncodingDeclarationsPrecedeLocaleFallback() {
        byte[] latin = Encoding.GetEncoding("iso-8859-1").GetBytes("<p>Café</p>");
        ChmDocument book = ChmDocument.Load(ChmFixture.Archive(new Dictionary<string, byte[]> {
            ["/#SYSTEM"] = ChmFixture.SystemMetadata(), ["/latin.html"] = latin,
            ["/utf8.html"] = ChmFixture.Html("<meta charset='utf-8'><p>Łódź</p>"),
            ["/bom.html"] = Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes("<p>日本</p>")).ToArray()
        }));
        Assert.Contains("Café", book.Topics.Single(item => item.Path == "/latin.html").ReadHtml());
        Assert.Contains("Łódź", book.Topics.Single(item => item.Path == "/utf8.html").ReadHtml());
        Assert.Contains("日本", book.Topics.Single(item => item.Path == "/bom.html").ReadHtml());
    }

    [Fact]
    public void ReferencesAreCaseInsensitiveDecodedAndConfinedToArchive() {
        ChmDocument book = ChmDocument.Load(ChmFixture.Archive(new Dictionary<string, byte[]> {
            ["/dir/a b.html"] = ChmFixture.Html("<p>Topic</p>"), ["/literal%20.html"] = ChmFixture.Html("<p>Literal</p>")
        }));
        Assert.Equal("/dir/a b.html", book.FindEntry("a%20b.html#part", "/dir/source.html")!.Path);
        Assert.NotNull(book.FindEntry("/DIR/A B.HTML")); Assert.NotNull(book.FindEntry("/literal%20.html"));
        foreach (string path in new[] { "../../secret", "../..%2fsecret", "https://example.com/topic", "file:///secret", "other.chm::/topic", "%00.html" })
            Assert.Null(book.FindEntry(path, "/dir/topic.html"));
    }

    [Fact]
    public async Task EscapedVirtualUrisPreserveDistinctSpaceAndLiteralPercentNames() {
        ChmDocument book = ChmDocument.Load(ChmFixture.Archive(new Dictionary<string, byte[]> {
            ["/a b.css"] = Encoding.UTF8.GetBytes("p {color:red}"), ["/a%20b.css"] = Encoding.UTF8.GetBytes("p {color:blue}"),
            ["/a b.html"] = ChmFixture.Html("<p>Space topic</p>"), ["/a%20b.html"] = ChmFixture.Html("<p>Percent topic</p>"),
            ["/start.html"] = ChmFixture.Html("<link rel='stylesheet' href='a%20b.css'><link rel='stylesheet' href='a%2520b.css'><a href='a%20b.html'>Space</a><a href='a%2520b.html'>Percent</a>")
        }));
        var epub = await book.ToEpubPublicationResultAsync(); Assert.True(epub.Succeeded);
        var styles = epub.Publication.Manifest.Where(item => item.MediaType == "text/css")
            .Select(item => Encoding.UTF8.GetString(epub.Publication.GetResourceBytes(item.Id))).ToArray();
        Assert.Contains("p {color:red}", styles); Assert.Contains("p {color:blue}", styles);
        var html = book.ToHtmlDocumentResult();
        var sections = html.Value.Document.QuerySelectorAll("section[data-chm-topic]");
        string space = sections.Single(section => section.GetAttribute("data-chm-topic") == "/a b.html").GetAttribute("id")!;
        string percent = sections.Single(section => section.GetAttribute("data-chm-topic") == "/a%20b.html").GetAttribute("id")!;
        var links = sections.Single(section => section.GetAttribute("data-chm-topic") == "/start.html").QuerySelectorAll("a");
        Assert.Equal("#" + space, links[0].GetAttribute("href")); Assert.Equal("#" + percent, links[1].GetAttribute("href"));
    }

    [Fact]
    public void BudgetsAndCancellationRejectWholeOperations() {
        Assert.Equal("CHM_INPUT_LIMIT", Assert.Throws<ChmReadException>(() => ChmDocument.Load(ChmFixture.Book(), new ChmReadOptions { MaxInputBytes = 10 })).Code);
        Assert.Equal("CHM_ENTRY_LIMIT", Assert.Throws<ChmReadException>(() => ChmDocument.Load(ChmFixture.Book(), new ChmReadOptions { MaxEntries = 2 })).Code);
        Assert.Equal("CHM_EXPANDED_LIMIT", Assert.Throws<ChmReadException>(() => ChmDocument.Load(ChmFixture.NativePath, new ChmReadOptions { MaxExpandedBytes = 1024 })).Code);
        Assert.Equal("CHM_NAVIGATION_LIMIT", Assert.Throws<ChmReadException>(() => ChmDocument.Load(ChmFixture.Book(), new ChmReadOptions { MaxNavigationItems = 1 })).Code);
        Assert.Throws<OperationCanceledException>(() => ChmDocument.Load(ChmFixture.Book(), cancellationToken: new CancellationToken(true)));
    }

    [Theory]
    [InlineData("/../escape.html")] [InlineData("/dir/./topic.html")] [InlineData("/dir//topic.html")] [InlineData("/bad:name.html")]
    public void UnsafeDirectoryIdentityIsRejected(string path) => Assert.Throws<ChmReadException>(() =>
        ChmDocument.Load(ChmFixture.Archive(new Dictionary<string, byte[]> { [path] = new byte[1] })));

    [Fact]
    public void TruncatedContainersAmbiguousPathsAndUnsupportedSectionsFailExplicitly() {
        byte[] valid = ChmFixture.Book();
        foreach (int length in new[] { 0, 40, 87, 95, 110, valid.Length - 1 })
            Assert.Throws<ChmReadException>(() => ChmDocument.Load(valid.Take(length).ToArray()));
        Assert.Equal("CHM_DUPLICATE_PATH", Assert.Throws<ChmReadException>(() => ChmDocument.Load(ChmFixture.Archive(new[] {
            new KeyValuePair<string, byte[]>("/a.html", new byte[1]), new KeyValuePair<string, byte[]>("/A.HTML", new byte[1])
        }))).Code);
        Assert.Equal("CHM_SECTION_UNSUPPORTED", Assert.Throws<ChmReadException>(() => ChmDocument.Load(ChmFixture.Archive(
            new Dictionary<string, byte[]> { ["/a.html"] = new byte[1] }, section: 2))).Code);
    }

    private sealed class ForwardOnlyStream : MemoryStream {
        internal ForwardOnlyStream(byte[] bytes) : base(bytes) { }
        public override bool CanSeek => false;
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    }
}
