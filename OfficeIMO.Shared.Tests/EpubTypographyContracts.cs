using OfficeIMO.Epub;
using OfficeIMO.Html;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubTypographyContracts {
    [Theory]
    [InlineData(EpubTypographyProfile.Basic, "line-height:1.5")]
    [InlineData(EpubTypographyProfile.Prose, "text-indent:1.2em")]
    [InlineData(EpubTypographyProfile.Technical, "border:.06em solid currentColor")]
    public void ImportedProfileSurvivesReopenAndPrecedesPublisherStyles(EpubTypographyProfile profile, string expectedDeclaration) {
        var input = HtmlConversionDocument.Parse("<html dir='rtl'><head><title>Profile</title><style>p{line-height:2}</style></head><body><h1>Title</h1><p>First paragraph.</p><p>Second paragraph.</p></body></html>");
        var options = new EpubManuscriptOptions { TypographyProfile = profile };
        var copy = options.Clone();
        options.TypographyProfile = EpubTypographyProfile.Basic;
        var result = EpubManuscript.ImportHtml(input, copy);
        result.Report.RequireNoLoss();
        var book = EpubPublication.Load(new MemoryStream(result.Publication.Write().Bytes));
        string css = Encoding.UTF8.GetString(book.GetResourceBytes("manuscript-defaults"));
        Assert.Contains(expectedDeclaration, css);
        Assert.Contains("white-space:pre-wrap", css);
        Assert.DoesNotContain("!important", css);
        XNamespace html = "http://www.w3.org/1999/xhtml";
        XDocument chapter = book.GetContentXml("chapter-1");
        Assert.Equal("rtl", (string?)chapter.Root!.Attribute("dir"));
        var links = chapter.Root.Element(html + "head")!.Elements(html + "link").Select(element => (string?)element.Attribute("href")).ToArray();
        Assert.Equal(2, links.Length);
        Assert.EndsWith("defaults.css", links[0]);
        Assert.Contains("line-height:2", Encoding.UTF8.GetString(book.GetResourceBytes("manuscript-style-1")));
    }

    [Fact]
    public void DisabledDefaultStylesRetainPublisherCssAndRejectUndefinedProfiles() {
        var input = HtmlConversionDocument.Parse("<title>Profile</title><style>p{line-height:2}</style><h1>Title</h1><p>Text.</p>");
        var result = EpubManuscript.ImportHtml(input, new EpubManuscriptOptions { IncludeDefaultStyles = false, TypographyProfile = EpubTypographyProfile.Technical });
        Assert.DoesNotContain(result.Publication.Manifest, item => item.Id == "manuscript-defaults");
        Assert.Contains("line-height:2", Encoding.UTF8.GetString(result.Publication.GetResourceBytes("manuscript-style-1")));
        Assert.Throws<ArgumentOutOfRangeException>(() => EpubManuscript.ImportHtml(input,
            new EpubManuscriptOptions { IncludeDefaultStyles = false, TypographyProfile = (EpubTypographyProfile)99 }));
    }
}
