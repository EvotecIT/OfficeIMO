using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public class HtmlQuotes {
        [Fact]
        public void QuotesRoundTrip() {
            const string html = "<p>Before <q>quoted</q> after</p>";
            using var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument();
            var runs = doc.Paragraphs[0].GetRuns().ToList();
            Assert.Equal("HtmlQuote", runs[1].CharacterStyleId);
            Assert.Equal("HtmlQuote", runs[3].CharacterStyleId);
            Assert.Equal("quoted", runs[2].Text);

            string roundTrip = doc.ToHtml();
            Assert.Contains("<q>quoted</q>", roundTrip, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void EquationInsideQuoteRoundTripsInsideQuote() {
            const string html = "<p>Before <q>quoted <math aria-label=\"x=1\"><mi>x</mi><mo>=</mo><mn>1</mn></math> after</q> tail</p>";
            using var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument();

            string roundTrip = doc.ToHtml();

            int quoteStart = roundTrip.IndexOf("<q>", StringComparison.OrdinalIgnoreCase);
            int math = roundTrip.IndexOf("<math", StringComparison.OrdinalIgnoreCase);
            int quoteEnd = roundTrip.IndexOf("</q>", StringComparison.OrdinalIgnoreCase);
            Assert.True(quoteStart >= 0 && quoteStart < math && math < quoteEnd, roundTrip);
        }
    }
}

