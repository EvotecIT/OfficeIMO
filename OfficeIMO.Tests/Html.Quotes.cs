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
            using var doc = html.LoadFromHtml();
            var runs = doc.Paragraphs[0].GetRuns().ToList();
            Assert.Equal("HtmlQuote", runs[1].CharacterStyleId);
            Assert.Equal("HtmlQuote", runs[3].CharacterStyleId);
            Assert.Equal("quoted", runs[2].Text);

            string roundTrip = doc.ToHtml();
            Assert.Contains("<q>quoted</q>", roundTrip, StringComparison.OrdinalIgnoreCase);
        }
    }
}

