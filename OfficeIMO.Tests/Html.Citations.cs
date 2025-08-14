using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public class HtmlCitations {
        [Fact]
        public void CiteIsItalicAndRoundsTrip() {
            const string html = "<p>This is a <cite>citation</cite>.</p>";
            using var doc = html.LoadFromHtml();
            var runs = doc.Paragraphs[0].GetRuns().ToList();
            Assert.Equal("HtmlCite", runs[1].CharacterStyleId);
            Assert.True(runs[1].Italic);
            Assert.Equal("citation", runs[1].Text);

            string roundTrip = doc.ToHtml();
            Assert.Contains("<cite", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("</cite>", roundTrip, StringComparison.OrdinalIgnoreCase);
        }
    }
}

