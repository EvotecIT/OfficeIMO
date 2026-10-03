using System.IO.Compression;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using OfficeIMO.Epub;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubWriterResourceBoundaryContracts {
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";
    private static readonly XNamespace Opf = "http://www.idpf.org/2007/opf";
    private static readonly XNamespace Dc = "http://purl.org/dc/elements/1.1/";

    [Theory]
    [InlineData("<img src='extra.png' alt='Extra'/>")]
    [InlineData("<style>p { background-image: url(extra.png); }</style>")]
    [InlineData("<style media='print'>p { background-image: url(extra.png); }</style>")]
    [InlineData("<style>@media print { p { background-image: url(extra.png); } }</style>")]
    [InlineData("<style media='print'>p { background-image: image-set(\"extra.png\" 1x); }</style>")]
    public async Task RewrittenContent_RequiresManifestDeclarationsForRetainedPayloads(string markup) {
        byte[] original = AddEntry(EpubWritingContracts.CreateBook().Write().Bytes, "EPUB/extra.png", new byte[] { 1 });
        EpubPublication book = EpubPublication.Load(new MemoryStream(original));
        Assert.Equal(original, book.Write().Bytes);
        SetBody(book, markup);
        await AssertRejectedSave<InvalidDataException>(book);
    }

    [Theory]
    [InlineData("<img src='https://example.test/remote.png' alt='Remote'/>", "image/png")]
    [InlineData("<iframe src='https://example.test/remote.xhtml'/>", "application/xhtml+xml")]
    [InlineData("<object data='https://example.test/remote.xhtml'/>", "application/xhtml+xml")]
    [InlineData("<link rel='stylesheet' href='https://example.test/remote.css'/>", "text/css")]
    [InlineData("<style>p { background-image: url(https://example.test/remote.png); }</style>", "image/png")]
    [InlineData("<style media='print'>p { background-image: url(https://example.test/remote.png); }</style>", "image/png")]
    [InlineData("<style>@media print { p { background-image: url(https://example.test/remote.png); } }</style>", "image/png")]
    [InlineData("<style media='print'>p { background-image: image-set(\"https://example.test/remote.png\" 1x); }</style>", "image/png")]
    public async Task RewrittenContent_RejectsProhibitedRemoteResourceKindsEvenWhenManifested(string markup, string mediaType) {
        EpubPublication book = LoadWithRemoteDeclaration("https://example.test/" + (mediaType == "text/css" ? "remote.css" :
            mediaType == "image/png" ? "remote.png" : "remote.xhtml"), mediaType);
        SetBody(book, markup);
        await AssertRejectedSave<NotSupportedException>(book);
    }

    [Theory]
    [InlineData("<iframe src='data:text/html,%3Cscript%3Ealert(1)%3C/script%3E'/>")]
    [InlineData("<object data='data:image/svg+xml,%3Csvg%20onload=%22alert(1)%22/%3E'/>")]
    [InlineData("<embed src='data:text/html;base64,PHNjcmlwdD5hbGVydCgxKTwvc2NyaXB0Pg=='/>")]
    [InlineData("<iframe src='data:image/png;base64,AQ=='/>")]
    [InlineData("<img src='data:image/svg+xml,%3Csvg%20onload=%22alert(1)%22/%3E' alt='Active'/>")]
    [InlineData("<style>p { background: url(data:image/svg+xml;base64,PHN2ZyBvbmxvYWQ9ImFsZXJ0KDEpIi8+); }</style>")]
    [InlineData("<style>@media print { p { background: url(data:image/svg+xml;base64,PHN2ZyBvbmxvYWQ9ImFsZXJ0KDEpIi8+); } }</style>")]
    [InlineData("<style media='print'>p { background: image-set(\"data:image/svg+xml;base64,PHN2ZyBvbmxvYWQ9ImFsZXJ0KDEpIi8+\" 1x); }</style>")]
    [InlineData("<a href='data:text/html,%3Cscript%3Ealert(1)%3C/script%3E'>Open</a>")]
    public async Task RawContent_CannotBypassTheNonScriptedBoundaryWithDataUrls(string markup) {
        EpubPublication book = EpubWritingContracts.CreateBook();
        XDocument content = BodyDocument(book, markup);
        book.UpdateResource("first", Encoding.UTF8.GetBytes(content.ToString()));
        await AssertRejectedSave<NotSupportedException>(book);
    }

    [Theory]
    [InlineData("chapter")]
    [InlineData("xml")]
    [InlineData("raw-add")]
    public async Task TypedAndRawInsertion_RejectEmbeddedDataDocumentsBeforeSaving(string route) {
        const string markup = "<iframe src='data:text/html,%3Cscript%3Ealert(1)%3C/script%3E'/>";
        EpubPublication book = EpubWritingContracts.CreateBook();
        if (route == "chapter") book.AddChapter("embedded", "EPUB/embedded.xhtml", "Embedded", markup);
        else if (route == "xml") SetBody(book, markup);
        else book.AddResource("embedded", "EPUB/embedded.xhtml", "application/xhtml+xml", Encoding.UTF8.GetBytes(EpubIntegrityFixtures.Xhtml(markup)));
        await AssertRejectedSave<NotSupportedException>(book);
    }

    [Fact]
    public void UnchangedImportedDataDocuments_RetainTheirPayloadAcrossMetadataEdits() {
        EpubPublication book = EpubWritingContracts.CreateBook();
        byte[] payload = Encoding.UTF8.GetBytes(EpubIntegrityFixtures.Xhtml("<iframe src='data:text/html,%3Cscript%3Ealert(1)%3C/script%3E'/>"));
        byte[] input = EditPackage(book.Write().Bytes, root => root.Descendants(Opf + "item").Single(item => (string?)item.Attribute("id") == "first")
            .SetAttributeValue("properties", "scripted"));
        input = EpubWritingContracts.ReplaceEntry(input, "EPUB/first.xhtml", payload);
        book = EpubPublication.Load(new MemoryStream(input)); Assert.Equal(input, book.Write().Bytes);
        book.Title = "Metadata edit";
        Assert.Equal(payload, EpubPublication.Load(new MemoryStream(book.Write().Bytes)).GetResourceBytes("first"));
    }

    [Fact]
    public void InertEmbeddedImagesAndManifestedResources_RemainAuthorable() {
        EpubPublication book = EpubWritingContracts.CreateBook();
        book.AddResource("image", "EPUB/image.png", "image/png", new byte[] { 1 });
        SetBody(book, "<img src='image.png' alt='Local'/><img src='data:image/png;base64,AQ==' alt='Embedded'/>");
        Assert.NotEmpty(book.Write().Bytes);
    }

    [Fact]
    public void InlineRemoteFonts_RequireAndRetainTheirManifestDeclaration() {
        EpubPublication book = LoadWithRemoteDeclaration("https://example.test/font.otf", "font/otf");
        SetBody(book, "<style media='print'>@font-face { font-family: test; src: url(https://example.test/font.otf); }</style>");
        Assert.Contains("remote-resources", EpubPublication.Load(new MemoryStream(book.Write().Bytes)).Manifest.Single(item => item.Id == "first").Properties!);
        book = EpubWritingContracts.CreateBook();
        SetBody(book, "<style>@font-face { font-family: test; src: url(https://example.test/font.otf); }</style>");
        Assert.Throws<InvalidDataException>(() => book.Write());
    }

    [Theory]
    [InlineData("images/")]
    [InlineData("./")]
    public void RelativeContentBase_IsAppliedToDirectAndSharedResourceDiscovery(string baseHref) {
        EpubPublication book = EpubWritingContracts.CreateBook();
        string path = baseHref == "images/" ? "EPUB/images/cover.png" : "EPUB/cover.png";
        book.AddResource("image", path, "image/png", new byte[] { 1 });
        XDocument content = BodyDocument(book, "<img src='cover.png' alt='Cover'/><p style='background-image: url(cover.png)'>Image</p>");
        content.Root!.Element(Html + "head")!.Elements(Html + "link").Remove();
        content.Root.Element(Html + "head")!.AddFirst(new XElement(Html + "base", new XAttribute("href", baseHref)));
        book.SetContentXml("first", content);
        Assert.NotEmpty(book.Write().Bytes);
    }

    [Theory]
    [InlineData("@media print")]
    [InlineData("@supports (unknown-property: value)")]
    public void ConditionalRemoteFonts_AreValidatedAndDeclaredAcrossContexts(string condition) {
        string css = condition + " { @font-face { font-family: test; src: url(https://example.test/font.otf); } }";
        EpubPublication book = EpubWritingContracts.CreateBook();
        SetBody(book, "<style>" + css + "</style>");
        Assert.Throws<InvalidDataException>(() => book.Write());
        book = LoadWithRemoteDeclaration("https://example.test/font.otf", "font/otf");
        SetBody(book, "<style>" + css + "</style>");
        Assert.Contains("remote-resources", EpubPublication.Load(new MemoryStream(book.Write().Bytes)).Manifest.Single(item => item.Id == "first").Properties!);
    }

    [Fact]
    public async Task SvgDataDocuments_AndNewEmbedsOfRetainedScriptedContentAreRejected() {
        EpubPublication book = EpubWritingContracts.CreateBook();
        book.AddResource("vector", "EPUB/vector.svg", "image/svg+xml", Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg' xmlns:xlink='http://www.w3.org/1999/xlink'><use xlink:href='data:image/svg+xml;base64,PHN2ZyBvbmxvYWQ9ImFsZXJ0KDEpIi8+' /></svg>"));
        await AssertRejectedSave<NotSupportedException>(book);
        byte[] input = EditPackage(EpubWritingContracts.CreateBook().Write().Bytes, root => root.Descendants(Opf + "item")
            .Single(item => (string?)item.Attribute("id") == "second").SetAttributeValue("properties", "scripted"));
        input = EpubWritingContracts.ReplaceEntry(input, "EPUB/second.xhtml", Encoding.UTF8.GetBytes(EpubIntegrityFixtures.Xhtml("<script>alert(1)</script>")));
        book = EpubPublication.Load(new MemoryStream(input));
        SetBody(book, "<iframe src='second.xhtml'/>");
        await AssertRejectedSave<NotSupportedException>(book);
    }

    [Theory]
    [InlineData(false, "javascript:alert(1)")]
    [InlineData(true, "javascript:alert(1)")]
    [InlineData(false, "file:///private.txt")]
    [InlineData(true, "vbscript:alert(1)")]
    public async Task SvgXlinkHyperlinks_UseTheSharedHyperlinkPolicy(bool standalone, string url) {
        EpubPublication book = EpubWritingContracts.CreateBook();
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' xmlns:xlink='http://www.w3.org/1999/xlink'><a xlink:href='" + url + "'><text>Open</text></a></svg>";
        if (standalone) book.AddResource("vector", "EPUB/vector.svg", "image/svg+xml", Encoding.UTF8.GetBytes(svg));
        else SetBody(book, svg);
        if (url.StartsWith("file:", StringComparison.Ordinal)) await AssertRejectedSave<InvalidDataException>(book);
        else await AssertRejectedSave<NotSupportedException>(book);
    }

    [Theory]
    [InlineData("EPUB/PACKAGE.OPF")]
    [InlineData("epub/package.opf")]
    public void VirtualPackagePath_IsReservedBeforeAnyMutation(string path) {
        EpubPublication book = EpubWritingContracts.CreateBook();
        byte[] before = book.Write().Bytes;
        Assert.Throws<ArgumentException>(() => book.AddResource("collision", path, "application/octet-stream", new byte[] { 1 }));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Theory]
    [InlineData("not a language")]
    [InlineData("en_US")]
    [InlineData("en-")]
    [InlineData("en-a")]
    [InlineData("en-a-foo-a-bar")]
    [InlineData("sl-rozaj-rozaj")]
    public void LanguageAuthoring_RejectsMalformedTagsAtomically(string language) {
        Assert.Throws<ArgumentException>(() => EpubPublication.Create("Book", language));
        EpubPublication book = EpubWritingContracts.CreateBook();
        byte[] before = book.Write().Bytes;
        Assert.Throws<ArgumentException>(() => book.Language = language);
        Assert.Throws<ArgumentException>(() => book.AddDublinCoreMetadata("language", language));
        Assert.Throws<ArgumentException>(() => book.AddDublinCoreMetadata("title", "Localized", language: language));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public async Task Save_ValidatesEveryRetainedDublinCoreLanguage() {
        byte[] input = EditPackage(EpubWritingContracts.CreateBook().Write().Bytes, root =>
            root.Element(Opf + "metadata")!.Add(new XElement(Dc + "language", "not a language")));
        EpubPublication book = EpubPublication.Load(new MemoryStream(input));
        Assert.Equal(input, book.Write().Bytes);
        book.Title = "Edited";
        await AssertRejectedSave<InvalidDataException>(book);
    }

    [Theory]
    [InlineData("pl-PL")]
    [InlineData("zh-Hant-TW")]
    [InlineData("sl-rozaj-biske-1994")]
    [InlineData("en-US-u-ca-gregory-x-book")]
    [InlineData("x-private")]
    [InlineData("i-klingon")]
    [InlineData("sgn-BE-FR")]
    public void LanguageAuthoring_RetainsWellFormedTagsBeyondInstalledCultures(string language) {
        EpubPublication book = EpubWritingContracts.CreateBook(); book.Language = language;
        book.AddDublinCoreMetadata("language", language);
        book.AddDublinCoreMetadata("title", "Localized", language: language);
        Assert.Equal(language, EpubPublication.Load(new MemoryStream(book.Write().Bytes)).Language);
    }

    internal static EpubPublication LoadWithRemoteDeclaration(string href, string mediaType) => EpubPublication.Load(new MemoryStream(
        EditPackage(EpubWritingContracts.CreateBook().Write().Bytes, root => root.Element(Opf + "manifest")!.Add(
            new XElement(Opf + "item", new XAttribute("id", "remote"), new XAttribute("href", href), new XAttribute("media-type", mediaType))))));

    private static void SetBody(EpubPublication book, string markup) => book.SetContentXml("first", BodyDocument(book, markup));
    private static XDocument BodyDocument(EpubPublication book, string markup) {
        XDocument content = book.GetContentXml("first");
        content.Root!.Element(Html + "body")!.Add(XElement.Parse("<div xmlns='" + Html + "'>" + markup + "</div>"));
        return content;
    }
    private static async Task AssertRejectedSave<T>(EpubPublication book) where T : Exception {
        using var output = new MemoryStream(); output.Write(new byte[] { 1, 2, 3 }, 0, 3); output.Position = 0;
        Assert.Throws<T>(() => book.Save(output));
        await Assert.ThrowsAsync<T>(() => book.SaveAsync(output));
        Assert.Equal(new byte[] { 1, 2, 3 }, output.ToArray());
    }
    private static byte[] EditPackage(byte[] input, Action<XElement> edit) {
        using var archive = new ZipArchive(new MemoryStream(input));
        using Stream stream = archive.GetEntry("EPUB/package.opf")!.Open();
        XDocument package = XDocument.Load(stream); edit(package.Root!);
        return EpubWritingContracts.ReplaceEntry(input, "EPUB/package.opf", Encoding.UTF8.GetBytes(package.ToString()));
    }
    private static byte[] AddEntry(byte[] input, string path, byte[] data) {
        using var output = new MemoryStream(); output.Write(input, 0, input.Length);
        using (var archive = new ZipArchive(output, ZipArchiveMode.Update, true)) {
            using Stream stream = archive.CreateEntry(path).Open(); stream.Write(data, 0, data.Length);
        }
        return output.ToArray();
    }
}
