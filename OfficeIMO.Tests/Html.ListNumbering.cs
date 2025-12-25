using System.Linq;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
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

        [Fact]
        public void HtmlToWord_ListNumbering_RtlDirection() {
            string html = "<ol dir=\"rtl\"><li>One</li><li>Two</li></ol>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            var listItems = doc.Paragraphs.Where(p => p.IsListItem).ToList();
            Assert.NotEmpty(listItems);
            Assert.All(listItems, p => Assert.True(p.BiDi));
        }
    }
}
