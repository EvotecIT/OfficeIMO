using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public class HtmlDefinitions {
        [Fact]
        public void DfnIsItalicAndRoundsTrip() {
            const string html = "<p>A <dfn>term</dfn> appears.</p>";
            using var doc = html.LoadFromHtml();
            var runs = doc.Paragraphs[0].GetRuns().ToList();
            Assert.Equal("HtmlDfn", runs[1].CharacterStyleId);
            Assert.True(runs[1].Italic);
            Assert.Equal("term", runs[1].Text);

            string roundTrip = doc.ToHtml();
            Assert.Contains("<dfn", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("</dfn>", roundTrip, StringComparison.OrdinalIgnoreCase);
        }
    }
}

