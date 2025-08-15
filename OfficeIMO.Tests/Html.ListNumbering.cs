        public void HtmlToWord_ListNumbering_SeparatedLists() {
            string html = "<ol><li>One</li></ol><p>Break</p><ol><li>Two</li></ol>";
            var options = new HtmlToWordOptions { ContinueNumbering = true };
            var doc = html.LoadFromHtml(options);
            Assert.True(doc.Paragraphs.Count(p => p.IsListItem) >= 2);
            Assert.Single(doc.Paragraphs.Where(p => p.IsListItem).Select(p => p._listNumberId).Distinct());
            Assert.Contains(doc.Paragraphs, p => !p.IsListItem);
        }

        [Fact]
        public void HtmlToWord_ListNumbering_NestedStartAndType_RoundTrip() {
            string html = "<ol start=\"5\" type=\"A\"><li>Outer<ol start=\"3\" type=\"a\"><li>Inner</li></ol></li></ol>";
            var doc = html.LoadFromHtml();
            string roundTrip = doc.ToHtml();
            Assert.Contains("<ol start=\"5\" type=\"A\">", roundTrip);
            Assert.Contains("<ol start=\"3\" type=\"a\">", roundTrip);
        }
    }
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_ListNumbering_ContiguousLists() {
            string html = "<ol><li>One</li></ol><ol><li>Two</li></ol>";
            var options = new HtmlToWordOptions { ContinueNumbering = true };
            var doc = html.LoadFromHtml(options);
            Assert.True(doc.Paragraphs.Count(p => p.IsListItem) >= 2);
            Assert.Single(doc.Paragraphs.Where(p => p.IsListItem).Select(p => p._listNumberId).Distinct());
        }

        [Fact]
        public void HtmlToWord_ListNumbering_SeparatedLists() {
            string html = "<ol><li>One</li></ol><p>Break</p><ol><li>Two</li></ol>";
            var options = new HtmlToWordOptions { ContinueNumbering = true };
            var doc = html.LoadFromHtml(options);
            Assert.True(doc.Paragraphs.Count(p => p.IsListItem) >= 2);
            Assert.Single(doc.Paragraphs.Where(p => p.IsListItem).Select(p => p._listNumberId).Distinct());
            Assert.Contains(doc.Paragraphs, p => !p.IsListItem);
        }
    }
}