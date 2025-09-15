using System.Linq;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite {
    public class Markdown_Reader_Quote_Content_Tests {
        [Fact]
        public void Quote_With_List_And_TaskItems() {
            var md = string.Join("\n", new[] {
                "> Shopping",
                ">",
                "> - Apples",
                "> - Bananas",
                "> - [x] Done"
            });
            var doc = MarkdownReader.Parse(md);
            var qb = Assert.IsType<QuoteBlock>(doc.Blocks[0]);
            Assert.True(qb.Children.Count >= 2);
            Assert.IsType<ParagraphBlock>(qb.Children[0]);
            var ul = Assert.IsType<UnorderedListBlock>(qb.Children[1]);
            Assert.Equal(3, ul.Items.Count);
            Assert.True(ul.Items[2].IsTask);
            Assert.True(ul.Items[2].Checked);
            var round = doc.ToMarkdown();
            Assert.Contains("> - [x] Done", round);
        }

        [Fact]
        public void Quote_With_CodeBlock_And_Caption() {
            var md = string.Join("\n", new[] {
                "> ```csharp",
                "> Console.WriteLine(\"Hi\");",
                "> ```",
                "> _Sample_"
            });
            var doc = MarkdownReader.Parse(md);
            var qb = Assert.IsType<QuoteBlock>(doc.Blocks[0]);
            var code = Assert.IsType<CodeBlock>(qb.Children[0]);
            Assert.Equal("csharp", code.Language);
            Assert.Equal("Sample", code.Caption);
            var outMd = doc.ToMarkdown();
            Assert.Contains("> ```csharp", outMd);
            Assert.Contains("> _Sample_", outMd);
        }

        [Fact]
        public void Quote_Nested_Quote_With_List() {
            var md = string.Join("\n", new[] {
                "> Outer",
                "> > Inner",
                "> > - a",
                "> > - b",
                "> After"
            });
            var doc = MarkdownReader.Parse(md);
            var qb = Assert.IsType<QuoteBlock>(doc.Blocks[0]);
            Assert.Equal(3, qb.Children.Count);
            var inner = Assert.IsType<QuoteBlock>(qb.Children[1]);
            Assert.IsType<UnorderedListBlock>(inner.Children[1]);
        }

        [Fact]
        public void Quote_With_Table() {
            var md = string.Join("\n", new[] {
                "> | A | B |",
                "> | --- | ---: |",
                "> | 1 | 2 |"
            });
            var doc = MarkdownReader.Parse(md);
            var qb = Assert.IsType<QuoteBlock>(doc.Blocks[0]);
            var table = Assert.IsType<TableBlock>(qb.Children.First());
            Assert.Equal(new[] { "A", "B" }, table.Headers);
            Assert.Equal(new[] { ColumnAlignment.None, ColumnAlignment.Right }, table.Alignments);
        }
    }
}
