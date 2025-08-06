using System.Linq;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public void MarkdownToWord_NestedOrderedAndUnorderedLists() {
            string md = "1. First\n   - Second\n     1. Third\n2. Fourth";

            var doc = md.LoadFromMarkdown(new MarkdownToWordOptions());

            var paragraphs = doc.Paragraphs.Where(p => p.IsListItem && !string.IsNullOrWhiteSpace(p.Text)).ToArray();
            Assert.Equal(new[] {0, 1, 2, 0}, paragraphs.Select(p => p.ListItemLevel.GetValueOrDefault()).ToArray());
            Assert.Equal(new[] {"First", "Second", "Third", "Fourth"}, paragraphs.Select(p => p.Text.Trim()).ToArray());
        }

        [Fact]
        public void MarkdownToWord_TaskLists() {
            string md = "- [ ] Task1\n- [x] Task2\n  - [ ] Subtask";

            var doc = md.LoadFromMarkdown(new MarkdownToWordOptions());

            Assert.Equal(new[] {false, true, false}, doc.CheckBoxes.Select(cb => cb.IsChecked).ToArray());

            var taskParagraphs = doc.Paragraphs.Where(p => p.IsListItem && !string.IsNullOrWhiteSpace(p.Text)).ToArray();
            Assert.Equal(new[] {0, 0, 1}, taskParagraphs.Select(p => p.ListItemLevel.GetValueOrDefault()).ToArray());
            Assert.Equal(new[] {"Task1", "Task2", "Subtask"}, taskParagraphs.Select(p => p.Text.Trim()).ToArray());
        }
    }
}
