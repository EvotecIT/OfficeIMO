using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using OfficeIMO.Epub;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubWriterContentClassificationContracts {
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";
    private static readonly XNamespace Opf = "http://www.idpf.org/2007/opf";

    [Theory]
    [InlineData("chapter")]
    [InlineData("xml")]
    [InlineData("raw-add")]
    [InlineData("raw-update")]
    public async Task Forms_AreRejectedThroughEveryContentAuthoringRoute(string route) {
        EpubPublication book = EpubWritingContracts.CreateBook();
        const string form = "<form><label>Name <input name='name'/></label></form>";
        byte[] before = book.Write().Bytes;
        if (route == "chapter" || route == "xml") {
            Assert.Throws<NotSupportedException>(() => {
                if (route == "chapter") book.AddChapter("form", "EPUB/form.xhtml", "Form", form);
                else {
                    XDocument content = book.GetContentXml("first");
                    content.Root!.Element(Html + "body")!.Add(XElement.Parse("<form xmlns='" + Html + "'><input name='name'/></form>"));
                    book.SetContentXml("first", content);
                }
            });
            Assert.Equal(before, book.Write().Bytes);
            return;
        }
        byte[] bytes = Encoding.UTF8.GetBytes(EpubIntegrityFixtures.Xhtml(form));
        if (route == "raw-add") book.AddResource("form", "EPUB/form.xhtml", "application/xhtml+xml", bytes);
        else book.UpdateResource("first", bytes);
        using var destination = new MemoryStream(new byte[] { 1, 2, 3 }, true);
        Assert.Throws<NotSupportedException>(() => book.Save(destination));
        await Assert.ThrowsAsync<NotSupportedException>(() => book.SaveAsync(destination));
        Assert.Equal(new byte[] { 1, 2, 3 }, destination.ToArray());
    }

    [Fact]
    public void Forms_RejectEmbeddedXhtmlFormsInSvgButRetainUnchangedImportedScriptedContent() {
        EpubPublication book = EpubWritingContracts.CreateBook();
        byte[] svg = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg'><foreignObject><form xmlns='http://www.w3.org/1999/xhtml'><input name='name'/></form></foreignObject></svg>");
        book.AddResource("vector", "EPUB/vector.svg", "image/svg+xml", svg);
        Assert.Throws<NotSupportedException>(() => book.Write());

        book = EpubWritingContracts.CreateBook();
        byte[] source = book.Write().Bytes;
        XDocument package = book.GetPackageXml();
        package.Descendants(Opf + "item").Single(item => (string?)item.Attribute("id") == "first").SetAttributeValue("properties", "scripted");
        source = EpubWritingContracts.ReplaceEntry(source, book.PackagePath, Encoding.UTF8.GetBytes(package.ToString()));
        byte[] form = Encoding.UTF8.GetBytes(EpubIntegrityFixtures.Xhtml("<form><input name='name'/></form>"));
        source = EpubWritingContracts.ReplaceEntry(source, "EPUB/first.xhtml", form);
        book = EpubPublication.Load(new MemoryStream(source));
        Assert.Equal(source, book.Write().Bytes);
        book.Title = "Metadata edit";
        EpubPublication saved = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        Assert.Equal(form, saved.GetResourceBytes("first"));
        Assert.Contains("scripted", saved.Manifest.Single(item => item.Id == "first").Properties!);
    }

    [Theory]
    [InlineData("audio")]
    [InlineData("video")]
    [InlineData("picture")]
    [InlineData("print-style")]
    [InlineData("alternate-style")]
    [InlineData("base")]
    public void RemoteDeclarations_IncludeUnselectedResourceAlternatives(string route) {
        EpubPublication book = EpubWritingContracts.CreateBook();
        XDocument content = book.GetContentXml("first");
        XElement head = content.Root!.Element(Html + "head")!;
        XElement body = content.Root.Element(Html + "body")!;
        if (route == "audio" || route == "video") {
            book.AddResource("media", "EPUB/local.mp3", "audio/mpeg", new byte[] { 1 });
            body.Add(new XElement(Html + route,
                new XElement(Html + "source", new XAttribute("src", "local.mp3"), new XAttribute("type", "audio/mpeg")),
                new XElement(Html + "source", new XAttribute("src", "https://example.test/remote.mp3"), new XAttribute("type", "audio/mpeg"))));
        } else if (route == "picture") {
            body.Add(new XElement(Html + "picture", new XElement(Html + "source", new XAttribute("srcset", "image.bin")),
                new XElement(Html + "img", new XAttribute("src", "https://example.test/remote.png"), new XAttribute("alt", "Fallback"))));
            book.AddResource("image", "EPUB/image.bin", "image/png", new byte[] { 1 });
        } else if (route == "base") {
            head.Add(new XElement(Html + "base", new XAttribute("href", "https://example.test/")));
            body.Add(new XElement(Html + "audio", new XAttribute("src", "remote.mp3")));
        } else head.Add(new XElement(Html + "link", new XAttribute("rel", route == "print-style" ? "stylesheet" : "alternate stylesheet"),
            new XAttribute("title", "Alternative"), new XAttribute("media", route == "print-style" ? "print" : "all"),
            new XAttribute("href", "https://example.test/remote.css")));
        book.SetContentXml("first", content);
        EpubPublication saved = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        Assert.Contains("remote-resources", saved.Manifest.Single(item => item.Id == "first").Properties!);
    }

    [Theory]
    [InlineData("javascript:alert(1)")]
    [InlineData("vbscript:alert(1)")]
    [InlineData("mailto:someone@example.test")]
    public async Task UnselectedSources_UseTheSharedResourceUrlPolicyBeforeSaving(string url) {
        EpubPublication book = EpubWritingContracts.CreateBook();
        book.AddResource("media", "EPUB/local.mp3", "audio/mpeg", new byte[] { 1 });
        XDocument content = book.GetContentXml("first");
        content.Root!.Element(Html + "body")!.Add(new XElement(Html + "audio",
            new XElement(Html + "source", new XAttribute("src", "local.mp3"), new XAttribute("type", "audio/mpeg")),
            new XElement(Html + "source", new XAttribute("src", url), new XAttribute("type", "audio/mpeg"))));
        book.SetContentXml("first", content);
        using var destination = new MemoryStream(new byte[] { 1, 2, 3 }, true);
        Assert.Throws<NotSupportedException>(() => book.Save(destination));
        await Assert.ThrowsAsync<NotSupportedException>(() => book.SaveAsync(destination));
        Assert.Equal(new byte[] { 1, 2, 3 }, destination.ToArray());
    }

    [Fact]
    public void Hyperlinks_DoNotAcquireRemoteResourceDeclarations() {
        EpubPublication book = EpubWritingContracts.CreateBook();
        XDocument content = book.GetContentXml("first");
        content.Root!.Element(Html + "head")!.Add(new XElement(Html + "link", new XAttribute("rel", "canonical"), new XAttribute("href", "https://example.test/book")));
        content.Root.Element(Html + "body")!.Add(new XElement(Html + "a", new XAttribute("href", "mailto:author@example.test"), "Contact"));
        book.SetContentXml("first", content);
        Assert.Null(EpubPublication.Load(new MemoryStream(book.Write().Bytes)).Manifest.Single(item => item.Id == "first").Properties);
    }
}
