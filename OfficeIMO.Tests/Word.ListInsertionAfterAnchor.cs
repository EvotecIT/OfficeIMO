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

        [Fact]
        public void Test_AddItemAfterAnchorWhenAnchorParagraphHasFollowingContent() {
            var filePath = Path.Combine(_directoryWithFiles, "ListInsertedAfterContentAnchor.docx");

            using (var document = WordDocument.Create(filePath)) {
                var placeholder = document.AddParagraph("place holder here");
                document.AddParagraph("This is some new content replacing the placeholder.");

                var list = placeholder.AddList(WordListStyle.Numbered);

                foreach (var item in new[] { "First item", "Second item", "Third item" }) {
                    var listItem = list.AddItem(null, wordParagraph: placeholder);
                    listItem.Text = item;
                }

                Assert.Equal(
                    new[] {
                        "place holder here",
                        "First item",
                        "Second item",
                        "Third item",
                        "This is some new content replacing the placeholder."
                    },
                    document.Paragraphs.Select(p => p.Text).ToArray());

                Assert.Equal(3, list.ListItems.Count);
                Assert.All(list.ListItems, p => Assert.True(p.IsListItem));

                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                Assert.Equal(
                    new[] {
                        "place holder here",
                        "First item",
                        "Second item",
                        "Third item",
                        "This is some new content replacing the placeholder."
                    },
                    document.Paragraphs.Select(p => p.Text).ToArray());

                var list = Assert.Single(document.Lists);
                Assert.Equal(3, list.ListItems.Count);
            }
        }
    }
}
