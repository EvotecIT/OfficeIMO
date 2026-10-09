using System.Xml.Linq;
using OfficeIMO.Epub;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Shared.Tests {

    public sealed class EpubManuscriptLegacyAnchorTests {
        [Theory]
        [InlineData("<p xml:id='legacy'>Existing</p><a name='legacy'>Alias</a>")]
        [InlineData("<a name='legacy'>Alias</a><p xml:id='legacy'>Existing</p>")]
        public void RetainedXmlIdsReserveLegacyDestinations(string body) {
            var result = EpubManuscript.ImportHtml(HtmlConversionDocument.Parse("<title>Manual</title><h1>Manual</h1>" + body));
            Assert.False(result.Succeeded);
            Assert.Contains(result.Report.FidelityDiagnostics, diagnostic => diagnostic.Code == "EPUB_IMPORT_ID_REFERENCE_INVALID");
            Assert.Single(result.Publication.GetContentXml("chapter-1").Descendants().Attributes(XNamespace.Xml + "id"), attribute => attribute.Value == "legacy");
            Assert.NotEmpty(result.Publication.Write().RequireValue());
        }

        [Theory]
        [InlineData("<a name='legacy' xml:id='legacy'>Target</a>")]
        [InlineData("<a name='legacy' id='legacy' xml:id='legacy'>Target</a>")]
        public void AnAnchorMayRetainItsOwnXmlId(string anchor) {
            var result = EpubManuscript.ImportHtml(HtmlConversionDocument.Parse("<title>Manual</title><h1>Manual</h1><a href='#legacy'>Go</a>" + anchor));
            Assert.True(result.Succeeded);
            Assert.NotEmpty(result.Publication.Write().RequireValue());
        }

        [Fact]
        public void DistinctLegacyAndXmlDestinationsBothRemainReachable() {
            var result = EpubManuscript.ImportHtml(HtmlConversionDocument.Parse("<title>Manual</title><h1>Manual</h1>" +
                "<a href='#legacy'>Old</a><a href='#modern'>New</a><a name='legacy' xml:id='modern'>Target</a>"));
            Assert.True(result.Succeeded);
            XDocument chapter = result.Publication.GetContentXml("chapter-1");
            Assert.Single(chapter.Descendants().Attributes("id"), attribute => attribute.Value == "legacy");
            Assert.Single(chapter.Descendants().Attributes(XNamespace.Xml + "id"), attribute => attribute.Value == "modern");
            Assert.NotEmpty(result.Publication.Write().RequireValue());
        }

        [Theory]
        [InlineData("", "#target")]
        [InlineData("<h1>Two</h1>", "chapter-0002.xhtml#target")]
        public void LinksResolveRetainedXmlIdsAcrossChapterBoundaries(string heading, string expected) {
            var result = EpubManuscript.ImportHtml(HtmlConversionDocument.Parse("<title>Manual</title><h1>One</h1><a href='#target'>Go</a>" + heading + "<p xml:id='target'>Target</p>"));
            Assert.True(result.Succeeded);
            Assert.Equal(expected, (string?)Assert.Single(result.Publication.GetContentXml("chapter-1").Descendants(XName.Get("a", "http://www.w3.org/1999/xhtml"))).Attribute("href"));
            Assert.NotEmpty(result.Publication.Write().RequireValue());
        }

        [Fact]
        public void SvgDataResourcesRetainTheirViewFragment() {
            string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 20 10'><view id='closeup' viewBox='10 0 10 10'/><rect width='20' height='10'/></svg>";
            string source = "data:image/svg+xml;base64," + Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(svg)) + "#closeup";
            var result = EpubManuscript.ImportHtml(HtmlConversionDocument.Parse("<title>Manual</title><h1>Manual</h1><svg xmlns='http://www.w3.org/2000/svg'><image href='" + source + "'/></svg>"));
            Assert.True(result.Succeeded);
            Assert.EndsWith("#closeup", (string?)Assert.Single(result.Publication.GetContentXml("chapter-1").Descendants(XName.Get("image", "http://www.w3.org/2000/svg"))).Attribute("href"));
            Assert.NotEmpty(result.Publication.Write().RequireValue());
        }

        [Theory]
        [InlineData("<meta id='legacy'>", "")]
        [InlineData("", "<script id='legacy'>ignored()</script>")]
        [InlineData("", "<style id='legacy'>p { color: red }</style>")]
        [InlineData("", "<form><p id='legacy'>Discarded subtree</p></form>")]
        [InlineData("", "<input disabled type='checkbox' id='legacy'>")]
        [InlineData("", "<button id='legacy'>Discarded control</button>")]
        public void OmittedElementsCannotReserveLegacyAnchorDestinations(string head, string omittedBody) {
            var result = EpubManuscript.ImportHtml(HtmlConversionDocument.Parse("<html><head><title>Manual</title>" + head +
                "</head><body><h1>Manual</h1><a href='#legacy'>Go</a>" + omittedBody + "<a name='legacy'>Target</a></body></html>"));
            Assert.True(result.Succeeded);
            Assert.DoesNotContain(result.Report.FidelityDiagnostics, diagnostic => diagnostic.Code == "EPUB_IMPORT_ID_REFERENCE_INVALID");
            Assert.Single(result.Publication.GetContentXml("chapter-1").Descendants().Attributes("id"), attribute => attribute.Value == "legacy");
            Assert.NotEmpty(result.Publication.Write().RequireValue());
        }

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
}
