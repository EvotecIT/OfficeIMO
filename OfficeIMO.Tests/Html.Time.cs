using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public class HtmlTimeTag {
        [Fact]
        public void TimeRoundsTripWithDateTime() {
            const string html = "<p>On <time datetime=\"2023-01-01\">2023-01-01</time> we met.</p>";
            using var doc = html.LoadFromHtml();
            var runs = doc.Paragraphs[0].GetRuns().ToList();
            Assert.Equal("HtmlTime", runs[1].CharacterStyleId);
            Assert.Equal("2023-01-01", runs[1].Text);

            string roundTrip = doc.ToHtml();
            Assert.Contains("<time", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("datetime=\"2023-01-01", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("2023-01-01</time>", roundTrip, StringComparison.OrdinalIgnoreCase);
        }
    }
}

