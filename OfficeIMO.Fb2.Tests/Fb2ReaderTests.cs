using System.Text;
using System.Xml;
using OfficeIMO.Epub;

namespace OfficeIMO.Fb2.Tests;

public sealed class Fb2ReaderTests {
    [Theory]
    [InlineData("harbour.fb2")]
    [InlineData("harbour.fb2.zip")]
    public void ImportsCalibreBookAndItsZipWithImagesAndLinks(string name) {
        using Stream source = File.OpenRead(Path.Combine(AppContext.BaseDirectory, "Fixtures", name));
        var imported = Fb2Reader.Import(source, name);
        using var artifact = new MemoryStream();
        imported.Value.Save(artifact);
        artifact.Position = 0;
        var book = EpubDocument.Load(artifact, new EpubLoadOptions { IncludeRawHtml = true, IncludeResourceData = true });
        Assert.Equal("Harbour Signal", book.Title);
        Assert.Equal("Ada Reader", book.Creator);
        Assert.Contains(book.Metadata, item => item.Name == "publisher" && item.Value == "Signal Press");
        Assert.Contains(book.Resources, item => item.Properties == "cover-image" && item.Data?.Length > 0);
        var body = book.Chapters.Single(chapter => chapter.Text.Contains("Signal received."));
        Assert.Contains("https://example.org/reference", body.Html);
        Assert.Contains("image-2.png", body.Html);
    }

    [Fact]
    public void RejectsDoctypeBeforeResolvingEntities() {
        using var source = new MemoryStream(Encoding.UTF8.GetBytes(
            "<!DOCTYPE FictionBook [<!ENTITY sample 'injected'>]><FictionBook xmlns='http://www.gribuser.ru/xml/fictionbook/2.0'><body>&sample;</body></FictionBook>"));
        Assert.Throws<XmlException>(() => Fb2Reader.Import(source));
    }

    [Theory]
    [InlineData("#foot")]
    [InlineData("#%66oot")]
    public void PreservesNestedContentsAndLinksIntoTheNotesBody(string target) {
        using var source = new MemoryStream(Encoding.UTF8.GetBytes($$"""
            <FictionBook xmlns="http://www.gribuser.ru/xml/fictionbook/2.0" xmlns:l="http://www.w3.org/1999/xlink">
              <description><title-info><book-title>Notes</book-title><lang>en</lang></title-info></description>
              <body><section id="start"><title><p>Opening</p><p>Signal</p></title><section><title><p>Nested</p></title>
              <p><a l:href="{{target}}">Note</a> <a l:href="file:///private/note">Local file</a></p></section></section></body>
              <body name="notes" xml:lang="fr"><section id="foot"><title><p>Footnote</p></title><p>Note body.</p></section></body>
            </FictionBook>
            """));
        var imported = Fb2Reader.Import(source);
        Assert.Contains(imported.Report.Diagnostics, item => item.Code == "FB2_LINK_OMITTED" && item.LossKind == OfficeConversionLossKind.Omission);
        using var output = new MemoryStream();
        imported.Value.Save(output);
        output.Position = 0;
        var book = EpubDocument.Load(output, new EpubLoadOptions { IncludeRawHtml = true });
        Assert.Equal("Opening Signal", book.TableOfContents[0].Children[0].Label);
        Assert.Equal("Nested", book.TableOfContents[0].Children[0].Children[0].Label);
        Assert.Contains("body-2.xhtml#fb2-3", book.Chapters[0].Html);
        Assert.DoesNotContain("file:", book.Chapters[0].Html);
        Assert.Contains("Note body.", book.Chapters[1].Text);
        Assert.Contains("xml:lang=\"fr\"", book.Chapters[1].Html);
    }
}
