using System.Linq;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public void MarkdownToWord_NestedOrderedAndUnorderedLists() {
            string md = "1. First\n   - Second\n     1. Third\n2. Fourth";

            var doc = OfficeIMO.Markdown.MarkdownReader.Parse(md).ToWordDocument(new MarkdownToWordOptions());

            var paragraphs = doc.Paragraphs.Where(p => p.IsListItem && !string.IsNullOrWhiteSpace(p.Text)).ToArray();
            Assert.Equal(new[] {0, 1, 2, 0}, paragraphs.Select(p => p.ListItemLevel.GetValueOrDefault()).ToArray());
            Assert.Equal(new[] {"First", "Second", "Third", "Fourth"}, paragraphs.Select(p => p.Text.Trim()).ToArray());
        }

        [Fact]
        public void MarkdownToWord_NestedOrderedList_PreservesItsOwnStartValue() {
            const string md = "1. Outer\n   4. Nested four\n   5. Nested five\n2. Next";

            using var doc = OfficeIMO.Markdown.MarkdownReader.Parse(md).ToWordDocument(new MarkdownToWordOptions());

            var paragraphs = doc.Paragraphs
                .Where(paragraph => paragraph.IsListItem && !string.IsNullOrWhiteSpace(paragraph.Text))
                .ToArray();
            Assert.Equal(new[] { "Outer", "Nested four", "Nested five", "Next" }, paragraphs.Select(paragraph => paragraph.Text.Trim()).ToArray());

            DocumentTraversal.ListInfo outer = DocumentTraversal.GetListInfo(paragraphs[0])!.Value;
            DocumentTraversal.ListInfo nested = DocumentTraversal.GetListInfo(paragraphs[1])!.Value;
            Assert.Equal(0, outer.Level);
            Assert.Equal(1, outer.Start);
            Assert.Equal(1, nested.Level);
            Assert.Equal(4, nested.Start);
        }

        [Fact]
        public void MarkdownToWord_TaskLists() {
            string md = "- [ ] Task1\n- [x] Task2\n  - [ ] Subtask";

            var doc = OfficeIMO.Markdown.MarkdownReader.Parse(md).ToWordDocument(new MarkdownToWordOptions());

            Assert.Equal(new[] {false, true, false}, doc.CheckBoxes.Select(cb => cb.IsChecked).ToArray());

            var taskParagraphs = doc.Paragraphs.Where(p => p.IsListItem && !string.IsNullOrWhiteSpace(p.Text)).ToArray();
            Assert.Equal(new[] {0, 0, 1}, taskParagraphs.Select(p => p.ListItemLevel.GetValueOrDefault()).ToArray());
            Assert.Equal(new[] {"Task1", "Task2", "Subtask"}, taskParagraphs.Select(p => p.Text.Trim()).ToArray());
        }

        [Fact]
        public void MarkdownToWord_ListItem_WithMultipleParagraphs_Preserves_AllParagraphBlocks() {
            const string md = """
                - first paragraph

                  second paragraph
                - next item
                """;

            using var doc = OfficeIMO.Markdown.MarkdownReader.Parse(md).ToWordDocument(new MarkdownToWordOptions());

            var listParagraphs = doc.Paragraphs
                .Where(p => p.IsListItem && !string.IsNullOrWhiteSpace(p.Text))
                .ToArray();

            Assert.Equal(new[] { 0, 0, 0 }, listParagraphs.Select(p => p.ListItemLevel.GetValueOrDefault()).ToArray());
            Assert.Equal(new[] { "first paragraph", "second paragraph", "next item" }, listParagraphs.Select(p => p.Text.Trim()).ToArray());
        }
    }
}
