using OfficeIMO.Html;
using OfficeIMO.Html.Dom;
using Xunit;

namespace OfficeIMO.Tests {

    public sealed class HtmlQualifiedAttributesTests {
        [Fact]
        public void SvgXlinkAndXmlAttributesRetainTheirQualifiedNamesAcrossNormalizationAndOwnedImport() {
            const string source = "<svg><g xml:lang='ar'><use href='#modern' xlink:href='#legacy'/></g></svg>";
            var document = HtmlConversionDocument.Parse(source);
            HtmlElement original = document.Document.QuerySelector("use")!;
            Assert.Equal("#legacy", original.GetAttribute("http://www.w3.org/1999/xlink", "href"));
            Assert.Equal("#legacy", original.GetAttribute("xlink:href"));
            var normalized = document.CreateDocumentForConversion();
            HtmlElement converted = normalized.QuerySelector("use")!;
            Assert.Equal("#modern", converted.GetAttribute("href"));
            Assert.Equal("#legacy", converted.GetAttribute("http://www.w3.org/1999/xlink", "href"));
            Assert.Equal("#legacy", converted.GetAttribute("xlink:href"));
            Assert.Equal("ar", normalized.QuerySelector("g")!.GetAttribute("http://www.w3.org/XML/1998/namespace", "lang"));
            Assert.Contains("xlink:href=", normalized.DocumentElement!.OuterHtml);
            Assert.Contains("xml:lang=", normalized.DocumentElement.OuterHtml);
        }

        private const string XlinkNamespace = "http://www.w3.org/1999/xlink";
        private const string SvgStart = "<svg xmlns='http://www.w3.org/2000/svg' xmlns:xlink='http://www.w3.org/1999/xlink'>";

        [Theory]
        [InlineData("<image data-note=\" href='decoy' \" xlink:href='real'/>", " href='decoy' ")]
        [InlineData("<image data-note=' href=\"decoy\" ' xlink:href='real'/>", " href=\"decoy\" ")]
        [InlineData("<image data-note=\" xlink:href='decoy' href='decoy' \" xlink:href=real />", " xlink:href='decoy' href='decoy' ")]
        public void SvgHrefNormalizationIgnoresAttributeTextInsideQuotedValues(string image, string note) {
            HtmlElement element = HtmlConversionDocument.Parse(SvgStart + image + "</svg>").Document.QuerySelector("image")!;

            Assert.Equal(note, element.GetAttribute("data-note"));
            Assert.Null(element.GetAttribute("href"));
            Assert.Equal("real", element.GetAttribute(XlinkNamespace, "href"));
        }

        [Theory]
        [InlineData("<image href='new' xlink:href='old'/>")]
        [InlineData("<image xlink:href='old' href='new'/>")]
        [InlineData("<image href=new data-note='x' xlink:href=old />")]
        [InlineData("<image data-note=it's href='new' xlink:href='old'/>")]
        [InlineData("<image href='new' href='duplicate' xlink:href='old'/>")]
        public void SvgHrefNormalizationKeepsBothRealHrefAttributes(string image) {
            HtmlElement element = HtmlConversionDocument.Parse(SvgStart + image + "</svg>").Document.QuerySelector("image")!;

            Assert.Equal("new", element.GetAttribute("href"));
            Assert.Equal("old", element.GetAttribute(XlinkNamespace, "href"));
        }

        [Theory]
        [InlineData("script")]
        [InlineData("style")]
        [InlineData("textarea")]
        [InlineData("title")]
        public void SvgHrefNormalizationLeavesRawTextUnchanged(string name) {
            const string text = "<svg><image href='new' xlink:href='old'/></svg>";

            HtmlConversionDocument document = HtmlConversionDocument.Parse("<" + name + ">" + text + "</" + name + ">");

            Assert.Equal(text, document.Document.QuerySelector(name)!.TextContent);
        }

        [Fact]
        public void SvgHrefNormalizationLeavesCommentsAndCdataUnchanged() {
            const string comment = " <svg><image href='new' xlink:href='old'/></svg> ";
            const string cdata = " a > b <image href='new' xlink:href='old'/> ";

            HtmlConversionDocument document = HtmlConversionDocument.Parse(
                "<div><!--" + comment + "--></div><svg><![CDATA[" + cdata + "]]></svg>");

            HtmlNode commentNode = Assert.Single(document.Document.QuerySelector("div")!.ChildNodes);
            Assert.Equal(HtmlNodeKind.Comment, commentNode.Kind);
            Assert.Equal(comment, commentNode.TextContent);
            Assert.Equal(cdata, document.Document.QuerySelector("svg")!.TextContent);
        }

        [Theory]
        [InlineData("<!-->")]
        [InlineData("<!--->")]
        [InlineData("<!-- x --!>")]
        public void SvgHrefNormalizationResumesAfterShortCommentClosers(string comment) {
            HtmlElement element = HtmlConversionDocument.Parse(
                comment + SvgStart + "<image href='new' xlink:href='old'/></svg><!-- -->").Document.QuerySelector("image")!;

            Assert.Equal("new", element.GetAttribute("href"));
            Assert.Equal("old", element.GetAttribute(XlinkNamespace, "href"));
        }

        [Fact]
        public void SvgHrefNormalizationReadsOnlyTheRealAnnotationXmlEncoding() {
            HtmlElement element = HtmlConversionDocument.Parse(
                "<math><annotation-xml data-note=\" encoding=x \" encoding='text/html'><section><svg>" +
                "<image href='new' xlink:href='old'/></svg></section></annotation-xml></math>").Document.QuerySelector("image")!;

            Assert.Equal("new", element.GetAttribute("href"));
            Assert.Equal("old", element.GetAttribute(XlinkNamespace, "href"));
        }
    }
}
