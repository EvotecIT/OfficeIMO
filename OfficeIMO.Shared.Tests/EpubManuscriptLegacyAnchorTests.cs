using System.Xml.Linq;
using OfficeIMO.Epub;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubManuscriptLegacyAnchorTests {
    [Theory]
    [InlineData("<p id='legacy'>Existing</p><a name='legacy'>Alias</a>")]
    [InlineData("<a name='legacy' id='modern'>Alias</a><p id='legacy'>Existing</p>")]
    [InlineData("<a name='legacy'>First</a><a name='legacy' id='modern'>Second</a>")]
    public void LegacyAnchorCollisionsAreReportedWithoutCreatingDuplicateIds(string body) {
        var result = EpubManuscript.ImportHtml(HtmlConversionDocument.Parse("<title>Manual</title><h1>Manual</h1><a href='#legacy'>Go</a>" + body));
        Assert.False(result.Succeeded);
        Assert.Contains(result.Report.FidelityDiagnostics, diagnostic => diagnostic.Code == "EPUB_IMPORT_ID_REFERENCE_INVALID" && diagnostic.LossKind == OfficeConversionLossKind.Failure);
        XDocument chapter = result.Publication.GetContentXml("chapter-1");
        string[] ids = chapter.Descendants().Attributes("id").Select(attribute => attribute.Value).ToArray();
        Assert.Equal(ids.Length, ids.Distinct(StringComparer.Ordinal).Count());
        Assert.Single(ids, id => id == "legacy");
        Assert.NotEmpty(result.Publication.Write().RequireValue());
    }

    [Fact]
    public void AnAnchorMayRetainTheSameNameAndIdOnItsOwnElement() {
        var result = EpubManuscript.ImportHtml(HtmlConversionDocument.Parse("<title>Manual</title><h1>Manual</h1><a href='#legacy'>Go</a><a name='legacy' id='legacy'>Target</a>"));
        Assert.True(result.Succeeded);
        Assert.Single(result.Publication.GetContentXml("chapter-1").Descendants().Attributes("id"), attribute => attribute.Value == "legacy");
        Assert.NotEmpty(result.Publication.Write().RequireValue());
    }
}
