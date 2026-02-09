using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite {
    public class Markdown_Reader_TablesAndLists_Tests {
        [Fact]
        public void Table_Parsing_Does_Not_Split_Escaped_Pipes() {
            string md = """
| Col1 | Col2 |
| --- | --- |
| a \| b | c |
""";
            var doc = MarkdownReader.Parse(md);
            var table = Assert.IsType<TableBlock>(doc.Blocks[0]);
            Assert.Equal(2, table.Headers.Count);
            Assert.Single(table.Rows);
            Assert.Equal(2, table.Rows[0].Count);
            Assert.Equal("a | b", table.Rows[0][0]);
        }

        [Fact]
        public void Table_Parsing_Does_Not_Split_Pipes_Inside_Code_Spans() {
            string md = """
| Col1 | Col2 |
| --- | --- |
| `a|b` | c |
""";
            var doc = MarkdownReader.Parse(md);
            var table = Assert.IsType<TableBlock>(doc.Blocks[0]);
            Assert.Equal(2, table.Headers.Count);
            Assert.Single(table.Rows);
            Assert.Equal(2, table.Rows[0].Count);

            var html = doc.ToHtmlFragment();
            Assert.Contains("<code>a|b</code>", html);
        }

        [Fact]
        public void Table_Headers_Render_Inline_Markup() {
            string md = """
| **H** | `C` |
| --- | --- |
| a | b |
""";
            var doc = MarkdownReader.Parse(md);
            var html = doc.ToHtmlFragment();
            Assert.Contains("<th><strong>H</strong></th>", html);
            Assert.Contains("<th><code>C</code></th>", html);
        }

        [Fact]
        public void Table_Cells_Resolve_Reference_Links() {
            string md = """
| Col |
| --- |
| [x][r] |

[r]: https://example.com
""";
            var doc = MarkdownReader.Parse(md);
            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("<a href=\"https://example.com\">x</a>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Table_Cells_Respect_Url_Policy() {
            string md = """
| Col |
| --- |
| [x](file:///c:/test) |
""";
            var doc = MarkdownReader.Parse(md, new MarkdownReaderOptions { DisallowFileUrls = true, HtmlBlocks = false, InlineHtml = false });
            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.DoesNotContain("file:", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(">x<", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Table_Does_Not_Trigger_On_Two_Lines_With_Pipes_Without_Alignment_Row() {
            string md = """
a | b
c | d
""";
            var doc = MarkdownReader.Parse(md);
            Assert.IsType<ParagraphBlock>(doc.Blocks[0]);
            Assert.DoesNotContain(doc.Blocks, b => b is TableBlock);
        }

        [Fact]
        public void Table_Does_Not_Trigger_On_Single_Outer_Pipe_Row() {
            string md = "| a | b |";
            var doc = MarkdownReader.Parse(md);
            Assert.IsType<ParagraphBlock>(doc.Blocks[0]);
            Assert.DoesNotContain(doc.Blocks, b => b is TableBlock);
        }

        [Fact]
        public void Table_Parses_Headerless_Tables_With_Outer_Pipes_And_Two_Rows() {
            string md = """
| a | b |
| c | d |
""";
            var doc = MarkdownReader.Parse(md);
            var table = Assert.IsType<TableBlock>(doc.Blocks[0]);
            Assert.Empty(table.Headers);
            Assert.Equal(2, table.Rows.Count);
            Assert.Equal("a", table.Rows[0][0].Trim());
            Assert.Equal("d", table.Rows[1][1].Trim());
        }

        [Fact]
        public void Builder_Paragraph_AutoSpacing_Remains_Convenient() {
            var doc = MarkdownDoc.Create().P(p => p.Text("Hello").Bold("World"));
            var md = doc.ToMarkdown().Replace("\r", "");
            Assert.Contains("Hello **World**", md);
        }

        [Fact]
        public void List_Item_Allows_Multiple_Paragraphs_When_Indented() {
            string md = """
- first paragraph

  second paragraph
- next item
""";
            var doc = MarkdownReader.Parse(md);
            var list = Assert.IsType<UnorderedListBlock>(doc.Blocks[0]);
            Assert.Equal(2, list.Items.Count);
            Assert.Single(list.Items[0].AdditionalParagraphs);

            var html = doc.ToHtmlFragment();
            Assert.Contains("<li><p>first paragraph</p><p>second paragraph</p></li>", html);
        }

        [Fact]
        public void List_Item_Can_Contain_Nested_Ordered_List() {
            string md = """
- outer
  1. one
  2. two
- next
""";
            var doc = MarkdownReader.Parse(md);
            var list = Assert.IsType<UnorderedListBlock>(doc.Blocks[0]);
            Assert.Equal(2, list.Items.Count);
            Assert.Single(list.Items[0].Children);
            Assert.IsType<OrderedListBlock>(list.Items[0].Children[0]);

            var html = doc.ToHtmlFragment();
            Assert.Contains("<li>outer<ol>", html);
        }

        [Fact]
        public void List_Item_Can_Contain_Nested_Fenced_Code_Block() {
            string md = """
- outer
  ```csharp
  Console.WriteLine(1);
  ```
- next
""";
            var doc = MarkdownReader.Parse(md);
            var list = Assert.IsType<UnorderedListBlock>(doc.Blocks[0]);
            Assert.Equal(2, list.Items.Count);
            Assert.Single(list.Items[0].Children);
            var code = Assert.IsType<CodeBlock>(list.Items[0].Children[0]);
            Assert.Equal("csharp", code.Language);

            var html = doc.ToHtmlFragment();
            Assert.Contains("<li>outer<pre><code class=\"language-csharp\">", html);
            Assert.Contains("Console.WriteLine(1);", html);
        }

        [Fact]
        public void List_Item_Can_Contain_Nested_Indented_Code_Block() {
            // Inside a list item, indented code is typically "continuation indent + 4 spaces".
            string md = """
- outer
      line1
      line2
- next
""";
            var doc = MarkdownReader.Parse(md);
            var list = Assert.IsType<UnorderedListBlock>(doc.Blocks[0]);
            Assert.Equal(2, list.Items.Count);
            Assert.Single(list.Items[0].Children);
            var code = Assert.IsType<CodeBlock>(list.Items[0].Children[0]);
            Assert.Contains("line1", code.Content);

            var html = doc.ToHtmlFragment();
            Assert.Contains("<pre><code>", html);
            Assert.Contains("line1", html);
        }

        [Fact]
        public void List_Item_Can_Contain_Nested_Blockquote() {
            string md = """
- outer
  > quote 1
  > quote 2
- next
""";
            var doc = MarkdownReader.Parse(md);
            var list = Assert.IsType<UnorderedListBlock>(doc.Blocks[0]);
            Assert.Equal(2, list.Items.Count);
            Assert.Single(list.Items[0].Children);
            Assert.IsType<QuoteBlock>(list.Items[0].Children[0]);

            var html = doc.ToHtmlFragment();
            Assert.Contains("<blockquote>", html);
            Assert.Contains("quote 1", html);
        }

        [Fact]
        public void List_Item_Can_Contain_Nested_Table() {
            string md = """
- outer
  | A | B |
  | --- | ---: |
  | 1 | 2 |
- next
""";
            var doc = MarkdownReader.Parse(md);
            var list = Assert.IsType<UnorderedListBlock>(doc.Blocks[0]);
            Assert.Equal(2, list.Items.Count);
            Assert.Single(list.Items[0].Children);
            var table = Assert.IsType<TableBlock>(list.Items[0].Children[0]);
            Assert.Equal(2, table.Headers.Count);

            var html = doc.ToHtmlFragment();
            Assert.Contains("<table>", html);
            Assert.Contains(">A<", html);
        }

        [Fact]
        public void List_Item_Can_Contain_Nested_Details_Block() {
            string md = """
- outer
  <details>
  <summary>More</summary>

  inner
  </details>
- next
""";
            var doc = MarkdownReader.Parse(md);
            var list = Assert.IsType<UnorderedListBlock>(doc.Blocks[0]);
            Assert.Equal(2, list.Items.Count);
            Assert.Contains(list.Items[0].Children, b => b is DetailsBlock);

            var html = doc.ToHtmlFragment();
            Assert.Contains("<details", html);
            Assert.Contains("<summary>More</summary>", html);
            Assert.Contains("inner", html);
        }

        [Fact]
        public void Unordered_List_Allows_Indented_Continuation_Lines() {
            string md = """
- first line
  second line
- next
""";
            var doc = MarkdownReader.Parse(md);
            var list = Assert.IsType<UnorderedListBlock>(doc.Blocks[0]);
            Assert.Equal(2, list.Items.Count);

            var html = doc.ToHtmlFragment();
            Assert.Contains("first line second line", html);
        }

        [Fact]
        public void Ordered_List_Allows_Indented_Continuation_Lines() {
            string md = """
1. first line
   second line
2. next
""";
            var doc = MarkdownReader.Parse(md);
            var list = Assert.IsType<OrderedListBlock>(doc.Blocks[0]);
            Assert.Equal(2, list.Items.Count);

            var html = doc.ToHtmlFragment();
            Assert.Contains("first line second line", html);
        }

        [Fact]
        public void List_Item_Does_Not_Break_Continuation_On_Pipe_Text() {
            string md = """
- outer
  a | b
  c | d
- next
""";
            var doc = MarkdownReader.Parse(md);
            var list = Assert.IsType<UnorderedListBlock>(doc.Blocks[0]);
            Assert.Equal(2, list.Items.Count);
            Assert.Empty(list.Items[0].Children);

            var html = doc.ToHtmlFragment();
            Assert.Contains("a | b c | d", html);
        }
    }
}
