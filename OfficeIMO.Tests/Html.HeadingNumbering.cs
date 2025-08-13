using System.Linq;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_HeadingNumbering_NestedHeadings() {
            string html = "<h1>One</h1><h2>Two</h2><h3>Three</h3>";
            var options = new HtmlToWordOptions { SupportsHeadingNumbering = true };
            var doc = html.LoadFromHtml(options);

            var headings = doc.Paragraphs.Where(p => p.IsListItem && (p.Text == "One" || p.Text == "Two" || p.Text == "Three")).ToArray();
            Assert.Equal(3, headings.Length);
            Assert.All(headings, p => Assert.Equal(WordListStyle.Headings111, p.ListStyle));
            Assert.Equal(new int?[] { 0, 1, 2 }, headings.Select(p => p.ListItemLevel).ToArray());
        }
    }
}
