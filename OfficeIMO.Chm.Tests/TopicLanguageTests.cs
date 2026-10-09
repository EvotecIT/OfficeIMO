using System.Xml.Linq;
using OfficeIMO.Epub;

namespace OfficeIMO.Chm.Tests;

public sealed class TopicLanguageTests {
    [Fact]
    public void TopicLanguageDirectionAndContainerAnchorsSurviveHtmlAndEpub() {
        ChmDocument book = ChmDocument.Load(ChmFixture.Archive(new Dictionary<string, byte[]> {
            ["/a.html"] = ChmFixture.Html("<html lang='ar' dir='rtl' class='arabic' id='root'><head><meta charset='utf-8'></head>" +
                "<body id='body' style='color: navy'><p>مرحبا</p><a href='#body'>Body</a></body></html>"),
            ["/b.html"] = ChmFixture.Html("<html lang='en' dir='ltr'><head><meta charset='utf-8'></head><body lang='fr'><p>Bonjour</p></body></html>")
        }));
        var projected = book.ToHtmlDocumentResult();
        Assert.Equal("ar", projected.Value.Document.DocumentElement!.GetAttribute("lang"));
        var sections = projected.Value.Document.QuerySelectorAll("section[data-chm-topic]");
        Assert.Equal("ar", sections[0].GetAttribute("lang")); Assert.Equal("rtl", sections[0].GetAttribute("dir"));
        Assert.Equal("fr", sections[1].GetAttribute("lang")); Assert.Equal("ltr", sections[1].GetAttribute("dir"));
        Assert.Equal("arabic", projected.Value.Document.QuerySelector("#chm-topic-1-root")!.GetAttribute("class"));
        Assert.Contains("navy", projected.Value.Document.QuerySelector("#chm-topic-1-body")!.GetAttribute("style"));
        Assert.Equal("#chm-topic-1-body", projected.Value.Document.QuerySelector("a")!.GetAttribute("href"));
        var imported = book.ToEpubPublicationResult(); Assert.True(imported.Succeeded);
        Assert.Equal("ar", imported.Publication.Language);
        XNamespace xhtml = "http://www.w3.org/1999/xhtml";
        XElement section = imported.Publication.GetContentXml("chapter-1").Descendants(xhtml + "section").First();
        Assert.Equal("ar", (string?)section.Attribute("lang")); Assert.Equal("rtl", (string?)section.Attribute("dir"));
        Assert.Equal("de", book.ToEpubPublicationResult(epubOptions: new EpubManuscriptOptions { Language = "de" }).Publication.Language);
    }

    [Fact]
    public void UndeclaredTopicLanguageUsesHelpBookLocaleAndUnsupportedContainerAttributesAreReported() {
        ChmDocument book = ChmDocument.Load(ChmFixture.Archive(new Dictionary<string, byte[]> {
            ["/#SYSTEM"] = ChmFixture.SystemMetadata(1045), ["/topic.html"] = ChmFixture.Html("<meta charset='utf-8'><body bgcolor='white'><p>Łódź</p></body>")
        }));
        var projection = book.ToHtmlDocumentResult();
        Assert.Equal("pl-PL", projection.Value.Document.QuerySelector("section")!.GetAttribute("lang"));
        Assert.Contains(projection.Report.FidelityDiagnostics, item => item.Code == "CHM_CONTAINER_ATTRIBUTE_OMITTED");
        Assert.Equal("pl-PL", book.ToEpubPublicationResult().Publication.Language);
    }
}
