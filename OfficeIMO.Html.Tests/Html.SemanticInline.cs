using OfficeIMO.Word.Html;
using System;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlDelAndInsRoundTripAsSemanticTags() {
            const string html = "<p>Changed <del>old</del> to <ins>new</ins>.</p>";
            using var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());

            string roundTrip = doc.ToHtml();

            Assert.Contains("<del>old</del>", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<ins>new</ins>", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("<s>old</s>", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("<u>new</u>", roundTrip, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void HtmlMarkRoundTripsAsSemanticTag() {
            const string html = "<p>Please <mark>review</mark> this.</p>";
            using var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());

            string roundTrip = doc.ToHtml();

            Assert.Contains("<mark>review</mark>", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("background-color", roundTrip, StringComparison.OrdinalIgnoreCase);
        }
    }
}
