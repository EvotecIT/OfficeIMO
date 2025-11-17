using System.IO;
using System.Linq;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_AddItemAfterAnchorOnlyForFirstInsertion() {
            var filePath = Path.Combine(_directoryWithFiles, "ListInsertedAfterAnchor.docx");

            using (var document = WordDocument.Create(filePath)) {
                var anchorParagraph = document.AddParagraph("Anchor paragraph");
                var list = document.AddList(WordListStyle.Bulleted);

                list.AddItem("First item", wordParagraph: anchorParagraph);
                list.AddItem("Second item", wordParagraph: anchorParagraph);
                list.AddItem("Third item", wordParagraph: anchorParagraph);

                Assert.Equal(4, document.Paragraphs.Count);
                Assert.Equal(new[] { "Anchor paragraph", "First item", "Second item", "Third item" }, document.Paragraphs.Select(p => p.Text).ToArray());
                Assert.False(document.Paragraphs[0].IsListItem);
                Assert.True(document.Paragraphs[1].IsListItem);
                Assert.True(document.Paragraphs[2].IsListItem);
                Assert.True(document.Paragraphs[3].IsListItem);
                Assert.Equal(new[] { "First item", "Second item", "Third item" }, list.ListItems.Select(i => i.Text).ToArray());

                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                Assert.Equal(4, document.Paragraphs.Count);
                Assert.Equal(new[] { "Anchor paragraph", "First item", "Second item", "Third item" }, document.Paragraphs.Select(p => p.Text).ToArray());

                var list = Assert.Single(document.Lists);
                Assert.Equal(new[] { "First item", "Second item", "Third item" }, list.ListItems.Select(i => i.Text).ToArray());
                Assert.False(document.Paragraphs[0].IsListItem);
                Assert.True(document.Paragraphs[1].IsListItem);
                Assert.True(document.Paragraphs[2].IsListItem);
                Assert.True(document.Paragraphs[3].IsListItem);
            }
        }
    }
}
