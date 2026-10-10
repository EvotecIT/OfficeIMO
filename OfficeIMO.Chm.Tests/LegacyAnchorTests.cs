using OfficeIMO.Epub;
using System.Xml.Linq;

namespace OfficeIMO.Chm.Tests;

public sealed class LegacyAnchorTests {
    [Fact]
    public void Html4NamedAnchorsRemainAddressableAcrossHelpTopicsAndInEpub() {
        ChmDocument book = ChmDocument.Load(ChmFixture.Archive(new Dictionary<string, byte[]> {
            ["/a.html"] = ChmFixture.Html("<p><a href='b.html#legacy'>Legacy destination</a><a href='b.html#modern'>Modern destination</a></p>"),
            ["/b.html"] = ChmFixture.Html("<a name='legacy' id='modern'>Destination</a><a name='named-only'></a><p>Target content.</p>")
        }));
        var converted = book.ToEpubPublicationResult(); Assert.True(converted.Succeeded);
        XNamespace xhtml = "http://www.w3.org/1999/xhtml";
        var roots = converted.Publication.Manifest.Where(item => item.MediaType == "application/xhtml+xml")
            .Select(item => converted.Publication.GetContentXml(item.Id)).ToArray();
        var ids = roots.SelectMany(root => root.Descendants().Attributes("id")).Select(attribute => attribute.Value).ToArray();
        Assert.Contains("chm-topic-2-legacy", ids); Assert.Contains("chm-topic-2-modern", ids); Assert.Contains("chm-topic-2-named-only", ids);
        Assert.DoesNotContain(roots.SelectMany(root => root.Descendants(xhtml + "a")), anchor => anchor.Attribute("name") != null);
        Assert.Contains(roots.SelectMany(root => root.Descendants(xhtml + "a").Attributes("href")), href => href.Value.EndsWith("#chm-topic-2-legacy"));
        Assert.NotEmpty(converted.Publication.Write().RequireValue());
    }
}
