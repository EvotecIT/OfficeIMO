using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using OfficeIMO.Html;
using System;
using System.IO;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void Html_FigureWithCaption_Converts() {
            string assetPath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png");
            byte[] imageBytes = File.ReadAllBytes(assetPath);
            string base64 = Convert.ToBase64String(imageBytes);
            string html = $"<figure><img src=\"data:image/png;base64,{base64}\" alt=\"Logo\"/><figcaption>Logo caption</figcaption></figure>";

            using var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());

            Assert.Single(doc.Images);
            Assert.Equal("Logo caption", doc.Paragraphs[1].Text);
            Assert.Equal("Caption", doc.Paragraphs[1].StyleId);

            string roundTrip = doc.ToHtml();
            Assert.Contains("<figure>", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<figcaption>Logo caption</figcaption>", roundTrip, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Html_FigureWithLeadingCaption_RoundTripsAsFigure() {
            string assetPath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png");
            byte[] imageBytes = File.ReadAllBytes(assetPath);
            string base64 = Convert.ToBase64String(imageBytes);
            string html = $"<figure><figcaption>Logo caption</figcaption><img src=\"data:image/png;base64,{base64}\" alt=\"Logo\"/></figure>";

            using var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());

            Assert.Single(doc.Images);
            Assert.Equal("Logo caption", doc.Paragraphs[0].Text);
            Assert.Equal("Caption", doc.Paragraphs[0].StyleId);

            string roundTrip = doc.ToHtml();
            Assert.Contains("<figure>", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<figcaption>Logo caption</figcaption>", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.True(
                roundTrip.IndexOf("<figcaption", StringComparison.OrdinalIgnoreCase) <
                roundTrip.IndexOf("<img", StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void WordToHtml_FigureWithCaption_RendersFigure() {
            string assetPath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png");
            using var doc = WordDocument.Create();
            doc.AddParagraph().AddImage(assetPath);
            doc.AddParagraph("Logo caption").SetStyleId("Caption");

            string html = doc.ToHtml();
            Assert.Contains("<figure>", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<figcaption>Logo caption</figcaption>", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Html_FigureWithMultipleContentBlocksReportsFlattening() {
            const string html = "<figure><p>First</p><p>Second</p><figcaption>Grouped content</figcaption></figure>";

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(new HtmlToWordOptions());
            using var document = conversion.Value;

            var diagnostic = Assert.Single(conversion.Report.Diagnostics, item => item.Code == "HtmlFigureStructureFlattened");
            Assert.Equal(HtmlConversionLossKind.Approximation, diagnostic.LossKind);
            Assert.Contains(document.Paragraphs, paragraph => paragraph.Text == "First");
            Assert.Contains(document.Paragraphs, paragraph => paragraph.Text == "Second");
            Assert.Contains(document.Paragraphs, paragraph => paragraph.Text == "Grouped content");
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void Html_NonImageFigureCaption_RoundTripsWithExactParentageAndCardinality(bool captionFirst) {
            string html = captionFirst
                ? "<figure><figcaption>Figure caption</figcaption><p>Figure body</p></figure>"
                : "<figure><p>Figure body</p><figcaption>Figure caption</figcaption></figure>";

            using var document = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());
            string roundTrip = document.ToHtml();
            var parsed = HtmlDocumentParser.ParseDocument(roundTrip);
            var figure = Assert.Single(parsed.QuerySelectorAll("figure"));

            Assert.Equal(2, figure.Children.Length);
            Assert.Equal(captionFirst ? "figcaption" : "p", figure.Children[0].LocalName);
            Assert.Equal(captionFirst ? "p" : "figcaption", figure.Children[1].LocalName);
            Assert.Equal("Figure caption", Assert.Single(figure.QuerySelectorAll(":scope > figcaption")).TextContent);
            Assert.Equal("Figure body", Assert.Single(figure.QuerySelectorAll(":scope > p")).TextContent);
            Assert.DoesNotContain(parsed.Body!.Children, element => element.LocalName == "p");
        }

        [Fact]
        public void Html_TableFigureWithLeadingCaption_RoundTripsAsOneFigure() {
            const string html = "<figure><figcaption>Table caption</figcaption><table><tr><td>Cell</td></tr></table></figure>";

            using var document = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());
            string roundTrip = document.ToHtml();
            var parsed = HtmlDocumentParser.ParseDocument(roundTrip);
            var figure = Assert.Single(parsed.QuerySelectorAll("figure"));

            Assert.Equal(2, figure.Children.Length);
            Assert.Equal("figcaption", figure.Children[0].LocalName);
            Assert.Equal("table", figure.Children[1].LocalName);
            Assert.Equal("Table caption", Assert.Single(figure.QuerySelectorAll(":scope > figcaption")).TextContent);
            Assert.Equal("Cell", Assert.Single(figure.QuerySelectorAll(":scope > table td")).TextContent);
        }
    }
}
