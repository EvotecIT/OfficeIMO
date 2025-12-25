using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
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

        [Fact]
        public void HtmlToWord_ListStyleType_FromCssOrdered() {
            string html = "<ol style=\"list-style-type: upper-roman\"><li>One</li><li>Two</li></ol>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            var first = doc.Paragraphs.First(p => p.IsListItem);
            var info = DocumentTraversal.GetListInfo(first);
            Assert.True(info.HasValue);
            Assert.Equal(NumberFormatValues.UpperRoman, info.Value.NumberFormat);
        }

        [Fact]
        public void HtmlToWord_ListStyleType_FromCssUnordered() {
            string html = "<ul style=\"list-style-type: square\"><li>One</li></ul>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            var first = doc.Paragraphs.First(p => p.IsListItem);
            var info = DocumentTraversal.GetListInfo(first);
            Assert.True(info.HasValue);
            Assert.Equal("■", info.Value.LevelText);
        }

        [Fact]
        public void HtmlToWord_ListItemValue_ResetsNumbering() {
            string html = "<ol><li>One</li><li value=\"5\">Five</li><li>Six</li></ol>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            var items = doc.Paragraphs
                .Where(p => p.IsListItem)
                .GroupBy(p => p._paragraph)
                .Select(g => g.First())
                .ToList();
            Assert.True(items.Count >= 3);

            var firstInfo = DocumentTraversal.GetListInfo(items[0]);
            var secondInfo = DocumentTraversal.GetListInfo(items[1]);

            Assert.True(firstInfo.HasValue);
            Assert.True(secondInfo.HasValue);
            Assert.Equal(1, firstInfo.Value.Start);
            Assert.Equal(5, secondInfo.Value.Start);
        }

        [Fact]
        public void HtmlToWord_ListReversed_DefaultStart() {
            string html = "<ol reversed><li>One</li><li>Two</li><li>Three</li></ol>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            var first = doc.Paragraphs.First(p => p.IsListItem);
            var info = DocumentTraversal.GetListInfo(first);
            Assert.True(info.HasValue);
            Assert.Equal(3, info.Value.Start);
        }
    }
}
