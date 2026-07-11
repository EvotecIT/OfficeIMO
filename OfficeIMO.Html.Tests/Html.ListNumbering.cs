using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_ListNumbering_ContiguousLists() {
            string html = "<ol><li>One</li></ol><ol><li>Two</li></ol>";
            var options = new HtmlToWordOptions { ContinueNumbering = true };
            var doc = html.ToWordDocument(options);
            Assert.True(doc.Paragraphs.Count(p => p.IsListItem) >= 2);
            Assert.Single(doc.Paragraphs.Where(p => p.IsListItem).Select(p => p._listNumberId).Distinct());
        }

        [Fact]
        public void HtmlToWord_ListNumbering_SeparatedLists() {
            string html = "<ol><li>One</li></ol><p>Break</p><ol><li>Two</li></ol>";
            var options = new HtmlToWordOptions { ContinueNumbering = true };
            var doc = html.ToWordDocument(options);
            Assert.True(doc.Paragraphs.Count(p => p.IsListItem) >= 2);
            Assert.Single(doc.Paragraphs.Where(p => p.IsListItem).Select(p => p._listNumberId).Distinct());
            Assert.Contains(doc.Paragraphs, p => !p.IsListItem);
        }

        [Fact]
        public void HtmlToWord_ListNumbering_RtlDirection() {
            string html = "<ol dir=\"rtl\"><li>One</li><li>Two</li></ol>";      

            var doc = html.ToWordDocument(new HtmlToWordOptions());

            var listItems = doc.Paragraphs.Where(p => p.IsListItem).ToList();   
            Assert.NotEmpty(listItems);
            Assert.All(listItems, p => Assert.True(p.BiDi));
        }

        [Fact]
        public void HtmlToWord_ListStyleType_FromCssOrdered() {
            string html = "<ol style=\"list-style-type: upper-roman\"><li>One</li><li>Two</li></ol>";

            var doc = html.ToWordDocument(new HtmlToWordOptions());

            var first = doc.Paragraphs.First(p => p.IsListItem);
            var info = DocumentTraversal.GetListInfo(first);
            Assert.True(info.HasValue);
            Assert.Equal(NumberFormatValues.UpperRoman, info.Value.NumberFormat);
        }

        [Fact]
        public void HtmlToWord_ListStyleType_FromCssInternationalOrdered() {
            string html = "<ol style=\"list-style-type: lower-russian !important\"><li>One</li><li>Two</li></ol>";

            var doc = html.ToWordDocument(new HtmlToWordOptions());

            var first = doc.Paragraphs.First(p => p.IsListItem);
            var info = DocumentTraversal.GetListInfo(first);
            Assert.True(info.HasValue);
            Assert.Equal(NumberFormatValues.RussianLower, info.Value.NumberFormat);
        }

        [Fact]
        public void HtmlToWord_ListStyleType_FromCssUnordered() {
            string html = "<ul style=\"list-style-type: square\"><li>One</li></ul>";

            var doc = html.ToWordDocument(new HtmlToWordOptions());

            var first = doc.Paragraphs.First(p => p.IsListItem);
            var info = DocumentTraversal.GetListInfo(first);
            Assert.True(info.HasValue);
            Assert.Equal("■", info.Value.LevelText);
        }

        [Fact]
        public void HtmlToWord_ListStyleType_FromQuotedDashBullet() {
            string html = "<ul style=\"list-style: '- ' outside\"><li>One</li></ul>";

            var doc = html.ToWordDocument(new HtmlToWordOptions());

            var first = doc.Paragraphs.First(p => p.IsListItem);
            var info = DocumentTraversal.GetListInfo(first);
            Assert.True(info.HasValue);
            Assert.Equal("-", info.Value.LevelText);
        }

        [Fact]
        public void HtmlToWord_ListStyleType_DecodesExportedDashMarkerEscapes() {
            string html = "<ul style=\"list-style-type:'\\2013'\"><li>One</li></ul><ul style=\"list-style-type:'\\2014'\"><li>Two</li></ul>";

            var doc = html.ToWordDocument(new HtmlToWordOptions());

            var listItems = doc.Paragraphs
                .Where(p => p.IsListItem)
                .GroupBy(p => p._paragraph)
                .Select(group => group.First())
                .ToList();
            Assert.Equal(2, listItems.Count);
            var first = DocumentTraversal.GetListInfo(listItems[0]);
            var second = DocumentTraversal.GetListInfo(listItems[1]);
            Assert.True(first.HasValue);
            Assert.True(second.HasValue);
            Assert.Equal("\u2013", first.Value.LevelText);
            Assert.Equal("\u2014", second.Value.LevelText);
        }

        [Fact]
        public void HtmlToWord_ListStyleType_ImportsQuotedEditorMarkers() {
            string html = "<ul style=\"list-style-type:'*'\"><li>Star</li></ul><ul style=\"list-style-type:'+'\"><li>Plus</li></ul>";

            var doc = html.ToWordDocument(new HtmlToWordOptions());

            var listItems = doc.Paragraphs
                .Where(p => p.IsListItem)
                .GroupBy(p => p._paragraph)
                .Select(group => group.First())
                .ToList();
            Assert.Equal(2, listItems.Count);
            var first = DocumentTraversal.GetListInfo(listItems[0]);
            var second = DocumentTraversal.GetListInfo(listItems[1]);
            Assert.True(first.HasValue);
            Assert.True(second.HasValue);
            Assert.Equal("*", first.Value.LevelText);
            Assert.Equal("+", second.Value.LevelText);
        }

        [Fact]
        public void HtmlToWord_ListStyleType_ImportsArbitraryQuotedMarkers() {
            string html = "<ul style=\"list-style-type:'\\2713'\"><li>Done</li></ul><ul style=\"list-style:'◆' outside\"><li>Diamond</li></ul>";

            var options = new HtmlToWordOptions();
            var doc = html.ToWordDocument(options);

            var listItems = doc.Paragraphs
                .Where(p => p.IsListItem)
                .GroupBy(p => p._paragraph)
                .Select(group => group.First())
                .ToList();
            Assert.Equal(2, listItems.Count);
            var first = DocumentTraversal.GetListInfo(listItems[0]);
            var second = DocumentTraversal.GetListInfo(listItems[1]);
            Assert.True(first.HasValue);
            Assert.True(second.HasValue);
            Assert.Equal("✓", first.Value.LevelText);
            Assert.Equal("◆", second.Value.LevelText);
            Assert.DoesNotContain(options.Diagnostics, diagnostic =>
                string.Equals(diagnostic.Code, "UnsupportedCssValue", StringComparison.OrdinalIgnoreCase) &&
                diagnostic.Source?.Contains("list-style", StringComparison.OrdinalIgnoreCase) == true);
        }

        [Fact]
        public void HtmlToWord_ListDefinitions_ApplyExportedIndentMetadata() {
            string html = "<ol data-left-indent-twips=\"1440\" data-hanging-indent-twips=\"360\" style=\"list-style-type:decimal\"><li>One</li></ol>";

            var doc = html.ToWordDocument(new HtmlToWordOptions());

            var first = doc.Paragraphs.First(p => p.IsListItem);
            var info = DocumentTraversal.GetListInfo(first);
            Assert.True(info.HasValue);
            Assert.Equal(1440, info.Value.LeftIndentTwips);
            Assert.Equal(360, info.Value.HangingIndentTwips);
        }

        [Fact]
        public void HtmlToWord_MarkdownTaskList_CheckboxInputsBecomeWordControls() {
            string html = "<ul class=\"contains-task-list\"><li class=\"task-list-item\"><input class=\"task-list-item-checkbox\" type=\"checkbox\" disabled checked>Done</li><li class=\"task-list-item\"><input class=\"task-list-item-checkbox\" type=\"checkbox\" disabled /> Open</li></ul>";

            var doc = html.ToWordDocument(new HtmlToWordOptions());

            var listItems = doc.Paragraphs
                .Where(p => p.IsListItem)
                .GroupBy(p => p._paragraph)
                .Select(g => g.OrderByDescending(p => p.Text.Length).First())
                .ToList();
            Assert.Equal(2, listItems.Count);
            Assert.Contains(listItems, p => p.Text.Contains("Done"));
            Assert.Contains(listItems, p => p.Text.Contains("Open"));

            var checkboxes = doc._wordprocessingDocument!.MainDocumentPart!.Document.Body!
                .Descendants<SdtRun>()
                .Select(run => run.SdtProperties?.Elements<W14.SdtContentCheckBox>().FirstOrDefault())
                .Where(checkBox => checkBox != null)
                .ToList();

            Assert.Equal(2, checkboxes.Count);
            Assert.Equal(W14.OnOffValues.One, checkboxes[0]!.Elements<W14.Checked>().Single().Val!.Value);
            Assert.Equal(W14.OnOffValues.Zero, checkboxes[1]!.Elements<W14.Checked>().Single().Val!.Value);
        }

        [Fact]
        public void HtmlToWord_ListItemBlockChildrenPreserveSourceOrder() {
            string html = "<ul><li>Intro<p>Details</p><table><tr><td>Metric</td></tr></table></li></ul><p>After</p>";

            using var doc = html.ToWordDocument(new HtmlToWordOptions());
            using MemoryStream stream = doc.SaveAsMemoryStream();
            stream.Position = 0;
            using WordprocessingDocument package = WordprocessingDocument.Open(stream, false);

            var body = package.MainDocumentPart!.Document.Body!;
            var sequence = body.ChildElements
                .Where(element => element is Paragraph || element is Table)
                .Select(element => element is Table
                    ? "table:" + string.Join("|", element.Descendants<Text>().Select(text => text.Text))
                    : "p:" + string.Concat(element.Descendants<Text>().Select(text => text.Text)))
                .Where(value => !string.Equals(value, "p:", StringComparison.Ordinal))
                .ToList();

            int introIndex = sequence.FindIndex(value => string.Equals(value, "p:Intro", StringComparison.Ordinal));
            int detailsIndex = sequence.FindIndex(value => string.Equals(value, "p:Details", StringComparison.Ordinal));
            int tableIndex = sequence.FindIndex(value => value.StartsWith("table:", StringComparison.Ordinal) && value.Contains("Metric", StringComparison.Ordinal));
            int afterIndex = sequence.FindIndex(value => string.Equals(value, "p:After", StringComparison.Ordinal));

            Assert.True(introIndex >= 0, string.Join(", ", sequence));
            Assert.True(detailsIndex > introIndex, string.Join(", ", sequence));
            Assert.True(tableIndex > detailsIndex, string.Join(", ", sequence));
            Assert.True(afterIndex > tableIndex, string.Join(", ", sequence));
        }

        [Fact]
        public void HtmlToWord_ListItemValue_ResetsNumbering() {
            string html = "<ol><li>One</li><li value=\"5\">Five</li><li>Six</li></ol>";

            var doc = html.ToWordDocument(new HtmlToWordOptions());

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

            var doc = html.ToWordDocument(new HtmlToWordOptions());

            var first = doc.Paragraphs.First(p => p.IsListItem);
            var info = DocumentTraversal.GetListInfo(first);
            Assert.True(info.HasValue);
            Assert.Equal(3, info.Value.Start);
        }
    }
}
