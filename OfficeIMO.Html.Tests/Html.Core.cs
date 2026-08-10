using OfficeIMO.Html;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlCoreTests {
    [Fact]
    public void HtmlConversionDocument_LoadDetectsMetaCharsetFromByteInput() {
        byte[] prefix = Encoding.ASCII.GetBytes("<meta charset='windows-1252'><p>caf");
        byte[] suffix = Encoding.ASCII.GetBytes("</p>");
        using var stream = new MemoryStream(prefix.Concat(new byte[] { 0xE9 }).Concat(suffix).ToArray());

        HtmlSemanticBlock paragraph = Assert.Single(HtmlConversionDocument.Load(stream)
            .SemanticDocument.Sections.SelectMany(section => section.Blocks));

        Assert.Equal("café", paragraph.Text);
    }

    [Fact]
    public void HtmlConversionDocument_LoadDetectsMetaCharsetFromNonSeekableInput() {
        byte[] html = BuildWindows1252Html();
        using var stream = new NonSeekableInputStream(html);

        HtmlSemanticBlock paragraph = Assert.Single(HtmlConversionDocument.Load(stream)
            .SemanticDocument.Sections.SelectMany(section => section.Blocks));

        Assert.Equal("café", paragraph.Text);
    }

    [Fact]
    public async Task HtmlConversionDocument_LoadAsyncDetectsMetaCharsetFromNonSeekableInput() {
        byte[] html = BuildWindows1252Html();
        using var stream = new NonSeekableInputStream(html);

        HtmlSemanticBlock paragraph = Assert.Single((await HtmlConversionDocument.LoadAsync(stream))
            .SemanticDocument.Sections.SelectMany(section => section.Blocks));

        Assert.Equal("café", paragraph.Text);
    }

    [Fact]
    public void HtmlConversionDocument_LoadIgnoresUnsupportedMetaCharset() {
        byte[] html = Encoding.UTF8.GetBytes("<meta charset='not-a-real-encoding'><p>café</p>");
        using var stream = new MemoryStream(html);

        HtmlSemanticBlock paragraph = Assert.Single(HtmlConversionDocument.Load(stream)
            .SemanticDocument.Sections.SelectMany(section => section.Blocks));

        Assert.Equal("café", paragraph.Text);
    }

    [Fact]
    public void HtmlConversionDocument_PrescanIgnoresCommentsAndUsesFirstRealDeclaration() {
        byte[] html = Encoding.UTF8.GetBytes(
            "<!-- <meta charset='windows-1252'> --><meta charset='utf-8'><p>€</p>");
        using var stream = new MemoryStream(html);

        HtmlSemanticBlock paragraph = Assert.Single(HtmlConversionDocument.Load(stream)
            .SemanticDocument.Sections.SelectMany(section => section.Blocks));

        Assert.Equal("€", paragraph.Text);
    }

    [Fact]
    public void HtmlConversionDocument_PrescanUsesHtmlEncodingLabelAliases() {
        byte[] prefix = Encoding.ASCII.GetBytes("<meta charset='iso-8859-1'><p>");
        byte[] html = prefix.Concat(new byte[] { 0x80 }).Concat(Encoding.ASCII.GetBytes("</p>")).ToArray();
        using var stream = new MemoryStream(html);

        HtmlSemanticBlock paragraph = Assert.Single(HtmlConversionDocument.Load(stream)
            .SemanticDocument.Sections.SelectMany(section => section.Blocks));

        Assert.Equal("€", paragraph.Text);
    }

    [Fact]
    public void HtmlConversionDocument_MetaDeclaredUtf16FallsBackToUtf8WithoutBom() {
        byte[] html = Encoding.UTF8.GetBytes("<meta charset='utf-16'><p>€</p>");
        using var stream = new MemoryStream(html);

        HtmlSemanticBlock paragraph = Assert.Single(HtmlConversionDocument.Load(stream)
            .SemanticDocument.Sections.SelectMany(section => section.Blocks));

        Assert.Equal("€", paragraph.Text);
    }

    [Fact]
    public void HtmlConversionDocument_PrescanDetectsBomlessUtf16XmlSignature() {
        byte[] html = new UnicodeEncoding(false, false, true)
            .GetBytes("<?xml version='1.0'?><p>café</p>");
        using var stream = new MemoryStream(html);

        HtmlSemanticBlock paragraph = Assert.Single(HtmlConversionDocument.Load(stream)
            .SemanticDocument.Sections.SelectMany(section => section.Blocks));

        Assert.Equal("café", paragraph.Text);
    }

    [Fact]
    public async Task HtmlConversionDocument_IgnoresXmlDeclarationEncoding() {
        byte[] html = Encoding.UTF8.GetBytes(
            "<?xml version='1.0' encoding='windows-1252'?><p>€</p>");

        using var syncStream = new MemoryStream(html);
        HtmlSemanticBlock syncParagraph = Assert.Single(HtmlConversionDocument.Load(syncStream)
            .SemanticDocument.Sections.SelectMany(section => section.Blocks));
        Assert.Equal("€", syncParagraph.Text);

        using var asyncStream = new MemoryStream(html);
        HtmlSemanticBlock asyncParagraph = Assert.Single((await HtmlConversionDocument.LoadAsync(asyncStream))
            .SemanticDocument.Sections.SelectMany(section => section.Blocks));
        Assert.Equal("€", asyncParagraph.Text);
    }

    [Fact]
    public async Task HtmlConversionDocument_MalformedBomDeclaredHtmlUsesReplacementCharacters() {
        byte[] utf8 = new byte[] { 0xEF, 0xBB, 0xBF }
            .Concat(Encoding.ASCII.GetBytes("<p>A"))
            .Concat(new byte[] { 0xFF })
            .Concat(Encoding.ASCII.GetBytes("B</p>"))
            .ToArray();
        byte[] utf16 = new byte[] { 0xFF, 0xFE }
            .Concat(new UnicodeEncoding(false, false).GetBytes("<p>A"))
            .Concat(new byte[] { 0x00, 0xD8 })
            .Concat(new UnicodeEncoding(false, false).GetBytes("B</p>"))
            .ToArray();

        foreach (byte[] html in new[] { utf8, utf16 }) {
            using var syncStream = new MemoryStream(html);
            HtmlSemanticBlock syncParagraph = Assert.Single(HtmlConversionDocument.Load(syncStream)
                .SemanticDocument.Sections.SelectMany(section => section.Blocks));
            Assert.Equal("A\uFFFDB", syncParagraph.Text);

            using var asyncStream = new MemoryStream(html);
            HtmlSemanticBlock asyncParagraph = Assert.Single((await HtmlConversionDocument.LoadAsync(asyncStream))
                .SemanticDocument.Sections.SelectMany(section => section.Blocks));
            Assert.Equal("A\uFFFDB", asyncParagraph.Text);
        }
    }

    [Fact]
    public void HtmlConversionDocument_InvalidCharsetAttributeRejectsLegacyContentFallbackOnSameMeta() {
        byte[] prefix = Encoding.ASCII.GetBytes(
            "<meta charset='invalid' http-equiv='content-type' content='text/html;charset=windows-1251'>"
            + "<meta charset='windows-1252'><p>caf");
        byte[] html = prefix.Concat(new byte[] { 0xE9 }).Concat(Encoding.ASCII.GetBytes("</p>")).ToArray();
        using var stream = new MemoryStream(html);

        HtmlSemanticBlock paragraph = Assert.Single(HtmlConversionDocument.Load(stream)
            .SemanticDocument.Sections.SelectMany(section => section.Blocks));

        Assert.Equal("café", paragraph.Text);
    }

    [Fact]
    public async Task HtmlConversionDocument_LoadAsyncUsesCancellableReadsForSeekablePrescan() {
        using var stream = new AsyncOnlySeekableStream(Encoding.UTF8.GetBytes("<p>Body</p>"));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            HtmlConversionDocument.LoadAsync(stream, cancellationToken: cancellation.Token));
    }

    [Fact]
    public void HtmlConversionDocument_ExplicitEncodingOverridesConflictingBom() {
        byte[] prefix = new byte[] { 0xEF, 0xBB, 0xBF }
            .Concat(Encoding.ASCII.GetBytes("<p>caf"))
            .ToArray();
        byte[] html = prefix.Concat(new byte[] { 0xE9 }).Concat(Encoding.ASCII.GetBytes("</p>")).ToArray();
        using var stream = new MemoryStream(html);

        HtmlConversionDocument conversion = HtmlConversionDocument.Load(
            stream,
            encoding: Encoding.GetEncoding("iso-8859-1"));

        Assert.Contains(conversion.SemanticDocument.Sections.SelectMany(section => section.Blocks),
            block => block.Text == "café");
    }

    [Fact]
    public void HtmlDataUri_DecodeTextHonorsDeclaredCharset() {
        Assert.True(HtmlDataUri.TryParse(
            "data:text/plain;charset=iso-8859-1;base64,6Q==",
            out HtmlDataUri dataUri));

        Assert.Equal("é", dataUri.DecodeText());
    }

    [Theory]
    [InlineData("iso-8859-1")]
    [InlineData("us-ascii")]
    public void HtmlDataUri_DecodeTextUsesWebEncodingAliases(string charset) {
        Assert.True(HtmlDataUri.TryParse(
            $"data:text/plain;charset={charset};base64,gA==",
            out HtmlDataUri dataUri));

        Assert.Equal("€", dataUri.DecodeText());
    }

    [Theory]
    [InlineData("data:text/plain;charset=bogus,hello")]
    [InlineData("data:text/plain;charset=utf-8,%FF")]
    public void HtmlDataUri_TryDecodeTextReturnsFalseForCharsetFailures(string source) {
        Assert.True(HtmlDataUri.TryParse(source, out HtmlDataUri dataUri));

        Assert.False(dataUri.TryDecodeText(out string text));
        Assert.Equal(string.Empty, text);
    }

    [Fact]
    public void HtmlStylesheetDecoderHonorsContentTypeAndCssCharset() {
        byte[] body = Encoding.ASCII.GetBytes(".label::before{content:'");
        byte[] suffix = Encoding.ASCII.GetBytes("';}");
        byte[] contentTypeBytes = body.Concat(new byte[] { 0xE9 }).Concat(suffix).ToArray();
        Assert.True(HtmlRenderStylesheetText.TryDecode(
            contentTypeBytes,
            "text/css; charset=windows-1252",
            out string contentTypeCss));

        byte[] declaration = Encoding.ASCII.GetBytes("@charset \"windows-1252\";.label::before{content:'");
        byte[] declarationBytes = declaration.Concat(new byte[] { 0xE9 }).Concat(suffix).ToArray();
        Assert.True(HtmlRenderStylesheetText.TryDecode(declarationBytes, "text/css", out string declaredCss));

        Assert.Contains("é", contentTypeCss, StringComparison.Ordinal);
        Assert.Contains("é", declaredCss, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("text/css; charset=us-ascii", null)]
    [InlineData("text/css", "@charset \"iso-8859-1\";")]
    public void HtmlStylesheetDecoderUsesWebEncodingAliases(string contentType, string? declaration) {
        byte[] prefix = Encoding.ASCII.GetBytes((declaration ?? string.Empty) + ".label::before{content:'");
        byte[] suffix = Encoding.ASCII.GetBytes("';}");
        byte[] stylesheet = prefix.Concat(new byte[] { 0x80 }).Concat(suffix).ToArray();

        Assert.True(HtmlRenderStylesheetText.TryDecode(stylesheet, contentType, out string css));
        Assert.Contains("€", css, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("utf-16")]
    [InlineData("utf-16le")]
    [InlineData("utf-16be")]
    public void HtmlStylesheetDecoderTreatsDeclaredUtf16LabelsAsUtf8(string charset) {
        byte[] stylesheet = Encoding.UTF8.GetBytes(".café{color:red}");

        Assert.True(HtmlRenderStylesheetText.TryDecode(
            stylesheet,
            $"text/css; charset={charset}",
            out string contentTypeCss));
        Assert.Equal(".café{color:red}", contentTypeCss);

        byte[] declared = Encoding.UTF8.GetBytes($"@charset \"{charset}\";.café{{color:red}}");
        Assert.True(HtmlRenderStylesheetText.TryDecode(declared, "text/css", out string declaredCss));
        Assert.Equal($"@charset \"{charset}\";.café{{color:red}}", declaredCss);
    }

    [Fact]
    public void HtmlStylesheetDecoderHonorsUtf16BomOverDeclaredEncoding() {
        var utf16 = new UnicodeEncoding(false, true, true);
        byte[] stylesheet = utf16.GetPreamble()
            .Concat(utf16.GetBytes(".café{color:red}"))
            .ToArray();

        Assert.True(HtmlRenderStylesheetText.TryDecode(
            stylesheet,
            "text/css; charset=utf-16be",
            out string css));
        Assert.Equal(".café{color:red}", css);
    }

    [Fact]
    public void HtmlStylesheetDecoderIgnoresNonCanonicalCharsetDeclaration() {
        byte[] stylesheet = Encoding.UTF8.GetBytes("@charset 'windows-1252';.café{color:red}");

        Assert.True(HtmlRenderStylesheetText.TryDecode(stylesheet, "text/css", out string css));
        Assert.Contains("café", css, StringComparison.Ordinal);
        Assert.DoesNotContain("cafÃ©", css, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlStylesheetDecoderParsesContentTypeParametersOutsideQuotedValues() {
        byte[] utf8Stylesheet = Encoding.UTF8.GetBytes(".café{color:red}");
        Assert.True(HtmlRenderStylesheetText.TryDecode(
            utf8Stylesheet,
            "text/css; title=\"x;charset=windows-1252\"",
            out string utf8Css));
        Assert.Contains("café", utf8Css, StringComparison.Ordinal);
        Assert.DoesNotContain("cafÃ©", utf8Css, StringComparison.Ordinal);

        Assert.True(HtmlRenderStylesheetText.TryDecode(
            utf8Stylesheet,
            "text/css; title=\"x\\\";charset=windows-1252\"",
            out string escapedQuoteCss));
        Assert.Contains("café", escapedQuoteCss, StringComparison.Ordinal);

        byte[] prefix = Encoding.ASCII.GetBytes(".label::before{content:'");
        byte[] suffix = Encoding.ASCII.GetBytes("';}");
        byte[] windows1252Stylesheet = prefix.Concat(new byte[] { 0xE9 }).Concat(suffix).ToArray();
        Assert.True(HtmlRenderStylesheetText.TryDecode(
            windows1252Stylesheet,
            "text/css; title=\"x;charset=utf-8\"; charset=\"windows-1252\"",
            out string windows1252Css));
        Assert.Contains("é", windows1252Css, StringComparison.Ordinal);
    }

    private static byte[] BuildWindows1252Html() {
        byte[] prefix = Encoding.ASCII.GetBytes("<meta charset='windows-1252'><p>caf");
        byte[] suffix = Encoding.ASCII.GetBytes("</p>");
        return prefix.Concat(new byte[] { 0xE9 }).Concat(suffix).ToArray();
    }

    private sealed class NonSeekableInputStream : Stream {
        private readonly MemoryStream _inner;

        internal NonSeekableInputStream(byte[] bytes) {
            _inner = new MemoryStream(bytes, writable: false);
        }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) =>
            _inner.ReadAsync(buffer, offset, count, cancellationToken);
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing) {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }
    }

    private sealed class AsyncOnlySeekableStream : MemoryStream {
        internal AsyncOnlySeekableStream(byte[] bytes) : base(bytes, writable: false) { }

        public override int Read(byte[] buffer, int offset, int count) =>
            throw new InvalidOperationException("The async load path must not use synchronous prefix reads.");
    }

    [Fact]
    public void HtmlDocumentParser_ResolvesBaseElementAgainstFallbackBaseUri() {
        var document = HtmlDocumentParser.ParseDocument("""<base href="images/"><p>Body</p>""");

        Uri? baseUri = HtmlDocumentParser.ResolveEffectiveBaseUri(
            document,
            new Uri("https://example.test/articles/2026/"));

        Assert.Equal("https://example.test/articles/2026/images/", baseUri?.AbsoluteUri);
        Assert.Equal(document.Body, HtmlDocumentParser.GetConversionRoot(document, useBodyContentsOnly: true));
    }

    [Fact]
    public void HtmlDocumentParser_ResolvesProtocolRelativeBaseElementAgainstWebSchemeWhenFallbackIsFile() {
        var document = HtmlDocumentParser.ParseDocument("""<base href="//cdn.example.test/assets/"><img src="logo.png">""");

        Uri? baseUri = HtmlDocumentParser.ResolveEffectiveBaseUri(
            document,
            new Uri("file:///C:/content/page.html"));

        Assert.Equal("https://cdn.example.test/assets/", baseUri?.AbsoluteUri);
    }

    [Fact]
    public void HtmlImageSourceResolver_ResolvesPictureSourceSetAgainstBaseUri() {
        var document = HtmlDocumentParser.ParseDocument("""
<picture>
  <source media="(min-width: 800px)" srcset="media/wide.webp 1x, media/wide@2x.webp 2x">
  <img src="media/fallback.png" alt="Storm">
</picture>
""");

        var picture = document.QuerySelector("picture")!;
        string source = HtmlImageSourceResolver.ResolvePictureSource(
            picture,
            new Uri("https://example.test/news/2026/"),
            HtmlUrlPolicy.CreateOfficeIMOProfile());

        Assert.Equal("https://example.test/news/2026/media/wide.webp", source);

        string normalized = HtmlImageSourceResolver.ResolveNormalizedSrcSetAttributes(
            picture.QuerySelector("source")!,
            new Uri("https://example.test/news/2026/"),
            HtmlUrlPolicy.CreateOfficeIMOProfile(),
            "srcset");

        Assert.Equal("https://example.test/news/2026/media/wide.webp 1x, https://example.test/news/2026/media/wide@2x.webp 2x", normalized);
    }

    [Fact]
    public void HtmlImageSourceResolver_OrdersParentPictureSourcesBeforeImageFallback() {
        var document = HtmlDocumentParser.ParseDocument("""
<picture>
  <source srcset="media/hero.webp 1x">
  <img src="media/fallback.png" alt="Hero">
</picture>
""");

        var image = document.QuerySelector("img")!;
        IReadOnlyList<string> candidates = HtmlImageSourceResolver.ResolveImageSourceCandidates(
            image,
            new Uri("https://example.test/news/2026/"),
            HtmlUrlPolicy.CreateOfficeIMOProfile());

        Assert.Collection(
            candidates,
            source => Assert.Equal("https://example.test/news/2026/media/hero.webp", source),
            source => Assert.Equal("https://example.test/news/2026/media/fallback.png", source));

        string resolved = HtmlImageSourceResolver.ResolveImageSource(
            image,
            new Uri("https://example.test/news/2026/"),
            HtmlUrlPolicy.CreateOfficeIMOProfile());

        Assert.Equal("https://example.test/news/2026/media/hero.webp", resolved);
    }

    [Fact]
    public void HtmlImageSourceResolver_UsesLazyAttributesBeforePlaceholderSource() {
        var document = HtmlDocumentParser.ParseDocument("""<img src="data:image/gif;base64,AAAA" data-lazy-src="media/photo.png">""");
        string source = HtmlImageSourceResolver.ResolveImageSource(
            document.QuerySelector("img")!,
            new Uri("https://example.test/gallery/"),
            HtmlUrlPolicy.CreateOfficeIMOProfile());

        Assert.Equal("https://example.test/gallery/media/photo.png", source);
    }

    [Fact]
    public void HtmlImageSourceResolver_UsesSourceSetBeforeImageSourceFallback() {
        var document = HtmlDocumentParser.ParseDocument("""<img src="media/fallback.png" srcset="media/hero.webp 1x" alt="Hero">""");
        var image = document.QuerySelector("img")!;

        IReadOnlyList<string> candidates = HtmlImageSourceResolver.ResolveImageSourceCandidates(
            image,
            new Uri("https://example.test/gallery/"),
            HtmlUrlPolicy.CreateOfficeIMOProfile());

        Assert.Collection(
            candidates,
            source => Assert.Equal("https://example.test/gallery/media/hero.webp", source),
            source => Assert.Equal("https://example.test/gallery/media/fallback.png", source));
        Assert.Equal(
            "https://example.test/gallery/media/hero.webp",
            HtmlImageSourceResolver.ResolveImageSource(image, new Uri("https://example.test/gallery/"), HtmlUrlPolicy.CreateOfficeIMOProfile()));
    }

    [Fact]
    public void HtmlImageSourceResolver_LimitsResponsiveCandidatesAndKeepsImageFallback() {
        var document = HtmlDocumentParser.ParseDocument("""
<picture>
  <source srcset="media/one.webp 1x, media/one.webp 2x, media/two.webp 3x, media/three.webp 4x">
  <img src="media/fallback.png" srcset="media/four.webp 4x" alt="Hero">
</picture>
""");
        var image = document.QuerySelector("img")!;

        IReadOnlyList<string> candidates = HtmlImageSourceResolver.ResolveImageSourceCandidates(
            image,
            new Uri("https://example.test/gallery/"),
            HtmlUrlPolicy.CreateOfficeIMOProfile(),
            allowParentPictureFallback: true,
            maxResponsiveCandidates: 2);

        Assert.Collection(
            candidates,
            source => Assert.Equal("https://example.test/gallery/media/one.webp", source),
            source => Assert.Equal("https://example.test/gallery/media/fallback.png", source));
    }

    [Fact]
    public void HtmlImageSourceResolver_CountsRejectedSrcSetEntriesTowardExpansionLimit() {
        var document = HtmlDocumentParser.ParseDocument("""
<img srcset="javascript:alert(1) 1x, javascript:alert(2) 2x, media/good.webp 3x" src="media/fallback.png" alt="Hero">
""");
        var image = document.QuerySelector("img")!;

        IReadOnlyList<string> candidates = HtmlImageSourceResolver.ResolveImageSourceCandidates(
            image,
            new Uri("https://example.test/gallery/"),
            HtmlUrlPolicy.CreateOfficeIMOProfile(),
            allowParentPictureFallback: true,
            maxResponsiveCandidates: 2);

        var source = Assert.Single(candidates);
        Assert.Equal("https://example.test/gallery/media/fallback.png", source);
    }

    [Fact]
    public void HtmlImageSourceResolver_CountsDuplicateSrcSetEntriesTowardExpansionLimit() {
        var document = HtmlDocumentParser.ParseDocument("""
<img srcset="media/one.webp 1x, media/one.webp 2x, media/two.webp 3x" src="media/fallback.png" alt="Hero">
""");
        var image = document.QuerySelector("img")!;

        IReadOnlyList<string> candidates = HtmlImageSourceResolver.ResolveImageSourceCandidates(
            image,
            new Uri("https://example.test/gallery/"),
            HtmlUrlPolicy.CreateOfficeIMOProfile(),
            allowParentPictureFallback: true,
            maxResponsiveCandidates: 2);

        Assert.Collection(
            candidates,
            source => Assert.Equal("https://example.test/gallery/media/one.webp", source),
            source => Assert.Equal("https://example.test/gallery/media/fallback.png", source));
    }

    [Fact]
    public void HtmlImageSourceResolver_CountsRejectedPictureUrlAttributesTowardExpansionLimit() {
        var document = HtmlDocumentParser.ParseDocument("""
<picture>
  <source src="javascript:alert(1)" data-lazy-src="media/good.webp">
  <img src="media/fallback.png" alt="Hero">
</picture>
""");
        var image = document.QuerySelector("img")!;

        IReadOnlyList<string> candidates = HtmlImageSourceResolver.ResolveImageSourceCandidates(
            image,
            new Uri("https://example.test/gallery/"),
            HtmlUrlPolicy.CreateOfficeIMOProfile(),
            allowParentPictureFallback: true,
            maxResponsiveCandidates: 1);

        var source = Assert.Single(candidates);
        Assert.Equal("https://example.test/gallery/media/fallback.png", source);
    }

    [Fact]
    public void HtmlSrcSetParser_CanLimitCandidateExpansion() {
        IReadOnlyList<HtmlSrcSetCandidate> candidates = HtmlSrcSetParser.Parse(
            "one.png 1x, two.png 2x, three.png 3x",
            maxCandidates: 2);

        Assert.Collection(
            candidates,
            candidate => Assert.Equal("one.png", candidate.Url),
            candidate => Assert.Equal("two.png", candidate.Url));
    }

    [Fact]
    public void HtmlSrcSetParser_SplitsCommaSeparatedCandidatesWithoutWhitespace() {
        IReadOnlyList<HtmlSrcSetCandidate> candidates = HtmlSrcSetParser.Parse("small.png,large.png 2x");

        Assert.Collection(
            candidates,
            candidate => {
                Assert.Equal("small.png", candidate.Url);
                Assert.Equal(string.Empty, candidate.Descriptor);
            },
            candidate => {
                Assert.Equal("large.png", candidate.Url);
                Assert.Equal("2x", candidate.Descriptor);
            });

        string normalized = HtmlImageSourceResolver.ResolveNormalizedSrcSet(
            "small.png,large.png 2x",
            new Uri("https://example.test/images/"),
            HtmlUrlPolicy.CreateOfficeIMOProfile());

        Assert.Equal("https://example.test/images/small.png, https://example.test/images/large.png 2x", normalized);
    }

    [Fact]
    public void HtmlSrcSetParser_SplitsQueryCandidatesAtCommaSeparators() {
        IReadOnlyList<HtmlSrcSetCandidate> candidates = HtmlSrcSetParser.Parse("small.png?v=1,large.png?v=1 2x");

        Assert.Collection(
            candidates,
            candidate => {
                Assert.Equal("small.png?v=1", candidate.Url);
                Assert.Equal(string.Empty, candidate.Descriptor);
            },
            candidate => {
                Assert.Equal("large.png?v=1", candidate.Url);
                Assert.Equal("2x", candidate.Descriptor);
            });
    }

    [Fact]
    public void HtmlSrcSetParser_SplitsBareExtensionlessCandidatesAtCommaSeparators() {
        IReadOnlyList<HtmlSrcSetCandidate> candidates = HtmlSrcSetParser.Parse("small,large 2x");

        Assert.Collection(
            candidates,
            candidate => {
                Assert.Equal("small", candidate.Url);
                Assert.Equal(string.Empty, candidate.Descriptor);
            },
            candidate => {
                Assert.Equal("large", candidate.Url);
                Assert.Equal("2x", candidate.Descriptor);
            });
    }

    [Fact]
    public void HtmlSrcSetParser_SplitsExtensionlessQueryCandidatesAtCommaSeparators() {
        IReadOnlyList<HtmlSrcSetCandidate> candidates = HtmlSrcSetParser.Parse("image?w=200,image?w=400 2x");

        Assert.Collection(
            candidates,
            candidate => {
                Assert.Equal("image?w=200", candidate.Url);
                Assert.Equal(string.Empty, candidate.Descriptor);
            },
            candidate => {
                Assert.Equal("image?w=400", candidate.Url);
                Assert.Equal("2x", candidate.Descriptor);
            });
    }

    [Fact]
    public void HtmlSrcSetParser_PreservesCommaValuedQueryStringsBeforeFollowingCandidate() {
        IReadOnlyList<HtmlSrcSetCandidate> candidates = HtmlSrcSetParser.Parse("photo.jpg?tags=red,blue 1x, photo@2x.jpg 2x");

        Assert.Collection(
            candidates,
            candidate => {
                Assert.Equal("photo.jpg?tags=red,blue", candidate.Url);
                Assert.Equal("1x", candidate.Descriptor);
            },
            candidate => {
                Assert.Equal("photo@2x.jpg", candidate.Url);
                Assert.Equal("2x", candidate.Descriptor);
            });
    }

    [Fact]
    public void HtmlSrcSetParser_SplitsDataUriCandidateBeforeFollowingUrl() {
        IReadOnlyList<HtmlSrcSetCandidate> candidates = HtmlSrcSetParser.Parse("data:image/png;base64,AAAA, https://cdn.test/fallback.png 2x");

        Assert.Collection(
            candidates,
            candidate => {
                Assert.Equal("data:image/png;base64,AAAA", candidate.Url);
                Assert.Equal(string.Empty, candidate.Descriptor);
            },
            candidate => {
                Assert.Equal("https://cdn.test/fallback.png", candidate.Url);
                Assert.Equal("2x", candidate.Descriptor);
            });
    }

    [Fact]
    public void HtmlImageDataUri_ParsesAndDecodesBase64Images() {
        string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\"/>";
        string dataUri = "data:image/svg+xml;base64," + Convert.ToBase64String(Encoding.UTF8.GetBytes(svg));

        Assert.True(HtmlImageDataUri.TryParse(dataUri, out var image));
        Assert.True(image.IsBase64);
        Assert.Equal("image/svg+xml", image.MediaType);
        Assert.Equal(".svg", image.FileExtension);
        Assert.Equal(svg, image.DecodeText());
        Assert.Equal(Encoding.UTF8.GetByteCount(svg), image.EstimateDecodedByteCount());
    }

    [Fact]
    public void HtmlImageDataUri_TryDecodeBytesReturnsFalseForBadEscapes() {
        Assert.True(HtmlImageDataUri.TryParse("data:image/png;base64,%ZZ", out var image));
        Assert.True(image.IsBase64);
        Assert.False(image.TryDecodeBytes(out byte[] bytes));
        Assert.Empty(bytes);
    }

    [Fact]
    public void HtmlImageDataUri_DecodesNonBase64PercentEscapesAsBytes() {
        Assert.True(HtmlImageDataUri.TryParse("data:image/png,%89PNG%0D%0A%1A%0A", out var image));

        Assert.False(image.IsBase64);
        Assert.Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }, image.DecodeBytes());
        Assert.Equal(8, image.EstimateDecodedByteCount());
    }

    [Fact]
    public void HtmlImageDataUri_MatchesOnlyExactBase64Flag() {
        string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\"/>";
        string dataUri = "data:image/svg+xml;name=base64," + Uri.EscapeDataString(svg);

        Assert.True(HtmlImageDataUri.TryParse(dataUri, out var image));
        Assert.False(image.IsBase64);
        Assert.Equal(svg, image.DecodeText());
    }

    [Fact]
    public void HtmlImageDataUri_IgnoresBase64WhitespaceWhenEstimatingDecodedBytes() {
        Assert.True(HtmlImageDataUri.TryParse("data:image/png;base64,QUJD%0A", out var image));

        Assert.True(image.IsBase64);
        Assert.Equal(3, image.EstimateDecodedByteCount());
        Assert.Equal(new byte[] { 65, 66, 67 }, image.DecodeBytes());
    }

    [Fact]
    public void HtmlUrlPolicyEvaluator_RejectsScriptUrlsAndCanResolveRelativeUrls() {
        var policy = HtmlUrlPolicy.CreateWebOnlyProfile();

        string rejected = HtmlUrlPolicyEvaluator.ResolveUrl("javascript:alert(1)", new Uri("https://example.test/"), policy);
        string resolved = HtmlUrlPolicyEvaluator.ResolveUrl("../docs/index.html", new Uri("https://example.test/news/2026/"), policy);
        string rootRelative = HtmlUrlPolicyEvaluator.ResolveUrl("/img/demo.png", new Uri("https://example.test/news/2026/"), policy);

        Assert.Equal(string.Empty, rejected);
        Assert.Equal("https://example.test/news/docs/index.html", resolved);
        Assert.Equal("https://example.test/img/demo.png", rootRelative);
    }

    [Fact]
    public void HtmlUrlPolicyEvaluator_TransformsResolvedUrlsAndRevalidatesTransformOutput() {
        var policy = HtmlUrlPolicy.CreateWebOnlyProfile();
        policy.ResolvedUrlTransform = value => value.Replace(
            "https://example.test/book/",
            "/library/book::");

        string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(
            "images/cover.png",
            new Uri("https://example.test/book/"),
            policy);

        Assert.Equal("/library/book::images/cover.png", resolved);
        Assert.NotNull(policy.Clone().ResolvedUrlTransform);

        policy.ResolvedUrlTransform = static _ => "javascript:alert(1)";
        Assert.Equal(
            string.Empty,
            HtmlUrlPolicyEvaluator.ResolveUrl("images/cover.png", new Uri("https://example.test/book/"), policy));
    }

    [Fact]
    public void HtmlUrlPolicyEvaluator_ResolvesProtocolRelativeUrlsAgainstWebSchemeWhenBaseIsFile() {
        var policy = HtmlUrlPolicy.CreateWebOnlyProfile();

        string resolved = HtmlUrlPolicyEvaluator.ResolveUrl("//cdn.example.test/app.png", new Uri("file:///C:/content/page.html"), policy);

        Assert.Equal("https://cdn.example.test/app.png", resolved);
    }

    [Fact]
    public void HtmlUrlPolicyEvaluator_ResolvesProtocolRelativeUrlsWithoutBaseAgainstHttps() {
        var policy = HtmlUrlPolicy.CreateOfficeIMOProfile();

        string resolved = HtmlUrlPolicyEvaluator.ResolveUrl("//cdn.example.test/app.png", null, policy);

        Assert.Equal("https://cdn.example.test/app.png", resolved);
    }

    [Theory]
    [InlineData("java\nscript:alert(1)")]
    [InlineData("vb\rscript:msgbox(1)")]
    [InlineData("java\tscript:alert(1)")]
    public void HtmlUrlPolicyEvaluator_RejectsUrlsWithEmbeddedControlCharacters(string rawUrl) {
        var policy = HtmlUrlPolicy.CreateWebOnlyProfile();

        Assert.False(HtmlUrlPolicyEvaluator.IsAllowed(rawUrl, policy));
        Assert.Equal(string.Empty, HtmlUrlPolicyEvaluator.ResolveUrl(rawUrl, new Uri("https://example.test/"), policy));
    }

    [Theory]
    [InlineData("C:secret.docx")]
    [InlineData("C:\\secret.docx")]
    [InlineData("C:/secret.docx")]
    public void HtmlUrlPolicyEvaluator_RejectsWindowsDrivePathsWhenFileUrlsAreDisallowed(string rawUrl) {
        var policy = HtmlUrlPolicy.CreateHyperlinkProfile();

        Assert.False(HtmlUrlPolicyEvaluator.IsAllowed(rawUrl, policy));
        Assert.Equal(string.Empty, HtmlUrlPolicyEvaluator.ResolveUrl(rawUrl, new Uri("https://example.test/"), policy));
    }
}
