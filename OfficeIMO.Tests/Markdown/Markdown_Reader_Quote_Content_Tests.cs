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
            Assert.Equal(2, qb.Children.Count);
            var inner = Assert.IsType<QuoteBlock>(qb.Children[1]);
            Assert.Equal(2, inner.Children.Count);
            var list = Assert.IsType<UnorderedListBlock>(inner.Children[1]);
            Assert.Equal(2, list.Items.Count);

            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("<blockquote><p>Outer</p><blockquote><p>Inner</p><ul><li>a</li><li>b After</li></ul></blockquote></blockquote>", html, StringComparison.Ordinal);
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

        [Fact]
        public void Quote_Lazy_Continuation_Extends_Unordered_List_Item() {
            const string md = "> - item\ncontinuation";

            var doc = MarkdownReader.Parse(md);
            var qb = Assert.IsType<QuoteBlock>(doc.Blocks[0]);
            var list = Assert.IsType<UnorderedListBlock>(qb.Children.First());
            Assert.Single(list.Items);

            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("<blockquote><ul><li>item continuation</li></ul></blockquote>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Quote_Lazy_Continuation_Extends_Ordered_List_Item() {
            const string md = "> 1. item\ncontinuation";

            var doc = MarkdownReader.Parse(md);
            var qb = Assert.IsType<QuoteBlock>(doc.Blocks[0]);
            var list = Assert.IsType<OrderedListBlock>(qb.Children.First());
            Assert.Single(list.Items);

            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("<blockquote><ol><li>item continuation</li></ol></blockquote>", html, StringComparison.Ordinal);
        }
    }
}
