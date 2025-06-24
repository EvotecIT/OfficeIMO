using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_ListItemsFromBodyTablesHeadersFooters() {
            var filePath = Path.Combine(_directoryWithFiles, "ListItemsEnumerator.docx");

            using (var document = WordDocument.Create(filePath)) {
                var bodyList = document.AddList(WordListStyle.Bulleted);
                bodyList.AddItem("Body1");
                bodyList.AddItem("Body2");

                var table = document.AddTable(1, 1);
                var tableList = table.Rows[0].Cells[0].AddList(WordListStyle.Bulleted);
                tableList.AddItem("Table1");
                tableList.AddItem("Table2");

                document.AddHeadersAndFooters();
                var headerList = document.Header.Default.AddList(WordListStyle.Bulleted);
                headerList.AddItem("Header1");
                headerList.AddItem("Header2");

                var footerList = document.Footer.Default.AddList(WordListStyle.Bulleted);
                footerList.AddItem("Footer1");
                footerList.AddItem("Footer2");

                Assert.Equal(4, document.Lists.Count);
                Assert.Equal(2, bodyList.ListItems.Count);
                Assert.Equal(2, tableList.ListItems.Count);
                Assert.Equal(2, headerList.ListItems.Count);
                Assert.Equal(2, footerList.ListItems.Count);

                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                Assert.Equal(4, document.Lists.Count);

                var bodyList = document.Lists.First(l => l.ListItems.First().Text == "Body1");
                Assert.Equal(new[] { "Body1", "Body2" }, bodyList.ListItems.Select(i => i.Text).ToArray());

                var tableList = document.Lists.First(l => l.ListItems.First().Text == "Table1");
                Assert.Equal(new[] { "Table1", "Table2" }, tableList.ListItems.Select(i => i.Text).ToArray());

                var headerList = document.Lists.First(l => l.ListItems.First().Text == "Header1");
                Assert.Equal(new[] { "Header1", "Header2" }, headerList.ListItems.Select(i => i.Text).ToArray());

                var footerList = document.Lists.First(l => l.ListItems.First().Text == "Footer1");
                Assert.Equal(new[] { "Footer1", "Footer2" }, footerList.ListItems.Select(i => i.Text).ToArray());
            }
        }
    }
}
