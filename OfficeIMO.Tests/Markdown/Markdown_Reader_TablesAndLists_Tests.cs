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
        public void Table_Parsing_Does_Not_Treat_Unmatched_Single_Backticks_As_Code_Spans() {
            string md = """
| Col1 | Col2 |
| --- | --- |
| `a | b |
""";
            var doc = MarkdownReader.Parse(md);
            var table = Assert.IsType<TableBlock>(doc.Blocks[0]);
            Assert.Equal(2, table.Headers.Count);
            Assert.Single(table.Rows);
            Assert.Equal(2, table.Rows[0].Count);
            Assert.Equal("`a", table.Rows[0][0]);
            Assert.Equal("b", table.Rows[0][1]);

            var html = doc.ToHtmlFragment();
            Assert.DoesNotContain("<code>`a</code>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Table_Parsing_Does_Not_Treat_Unmatched_Multi_Backticks_As_Code_Spans() {
            string md = """
| Col1 | Col2 |
| --- | --- |
| ``a | b |
""";
            var doc = MarkdownReader.Parse(md);
            var table = Assert.IsType<TableBlock>(doc.Blocks[0]);
            Assert.Equal(2, table.Headers.Count);
            Assert.Single(table.Rows);
            Assert.Equal(2, table.Rows[0].Count);
            Assert.Equal("``a", table.Rows[0][0]);
            Assert.Equal("b", table.Rows[0][1]);

            var html = doc.ToHtmlFragment();
            Assert.DoesNotContain("<code>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Table_Headers_Render_Inline_Markup() {
            string md = """
| **H** | `C` |
| --- | --- |
| a | b |
""";
            var doc = MarkdownReader.Parse(md);
            var table = Assert.IsType<TableBlock>(doc.Blocks[0]);
            Assert.Equal(2, table.HeaderInlines.Count);
            Assert.Single(table.RowInlines);
            Assert.Equal("**H**", table.HeaderInlines[0].RenderMarkdown());
            Assert.Equal("`C`", table.HeaderInlines[1].RenderMarkdown());
            var html = doc.ToHtmlFragment();
            Assert.Contains("<th><strong>H</strong></th>", html);
            Assert.Contains("<th><code>C</code></th>", html);
        }

        [Fact]
        public void Table_Cells_Are_Exposed_As_Typed_Block_Content() {
            string md = """
| **H** | `C` |
| --- | --- |
| [a](https://example.com) | b |
""";
            var doc = MarkdownReader.Parse(md);
            var table = Assert.IsType<TableBlock>(doc.Blocks[0]);

            Assert.Collection(table.HeaderCells,
                cell => Assert.Equal("**H**", Assert.IsType<ParagraphBlock>(Assert.Single(cell.Blocks)).Inlines.RenderMarkdown()),
                cell => Assert.Equal("`C`", Assert.IsType<ParagraphBlock>(Assert.Single(cell.Blocks)).Inlines.RenderMarkdown()));

            Assert.Single(table.RowCells);
            Assert.Collection(table.RowCells[0],
                cell => Assert.Equal("[a](https://example.com)", Assert.IsType<ParagraphBlock>(Assert.Single(cell.Blocks)).Inlines.RenderMarkdown()),
                cell => Assert.Equal("b", Assert.IsType<ParagraphBlock>(Assert.Single(cell.Blocks)).Inlines.RenderMarkdown()));
        }

        [Fact]
        public void Table_Cells_Can_Parse_MultiBlock_Markdown_Content_From_Cell_Body() {
            string md = """
| Section | Notes |
| --- | --- |
| Alpha | Intro<br><br>> Quoted |
""";
            var doc = MarkdownReader.Parse(md);
            var table = Assert.IsType<TableBlock>(Assert.Single(doc.Blocks));

            Assert.Collection(table.RowCells[0][1].Blocks,
                block => Assert.Equal("Intro", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
                block => Assert.IsType<QuoteBlock>(block));

            var markdown = doc.ToMarkdown().Replace("\r\n", "\n");
            Assert.Contains("Intro<br><br>> Quoted", markdown, StringComparison.Ordinal);

            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("<td><p>Intro</p><blockquote><p>Quoted</p></blockquote></td>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Table_Cells_Keep_SingleBreak_Content_As_A_Paragraph() {
            string md = """
| Notes |
| --- |
| First<br>Second |
""";
            var doc = MarkdownReader.Parse(md);
            var table = Assert.IsType<TableBlock>(Assert.Single(doc.Blocks));

            var cell = table.RowCells[0][0];
            var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(cell.Blocks));
            Assert.DoesNotContain(cell.Blocks, block => block is QuoteBlock or UnorderedListBlock or OrderedListBlock);
            Assert.Contains("First", paragraph.Inlines.RenderMarkdown(), StringComparison.Ordinal);
            Assert.Contains("Second", paragraph.Inlines.RenderMarkdown(), StringComparison.Ordinal);

            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("First<br", html, StringComparison.Ordinal);
            Assert.Contains("Second", html, StringComparison.Ordinal);
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
            var table = Assert.IsType<TableBlock>(doc.Blocks[0]);
            Assert.Single(table.RowInlines);
            Assert.Equal("[x](https://example.com)", table.RowInlines[0][0].RenderMarkdown());
            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("<a href=\"https://example.com\">x</a>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Table_RenderHtml_Falls_Back_To_Current_StringCells_After_Mutation() {
            string md = """
| Header |
| --- |
| value |
""";
            var doc = MarkdownReader.Parse(md);
            var table = Assert.IsType<TableBlock>(doc.Blocks[0]);

            table.Headers[0] = "**Changed**";
            table.Rows[0] = new[] { "[fresh](https://example.com)" };

            var html = ((IMarkdownBlock)table).RenderHtml();

            Assert.Contains("<th><strong>Changed</strong></th>", html, StringComparison.Ordinal);
            Assert.Contains("<a href=\"https://example.com\">fresh</a>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Table_InlineCells_Follow_Current_String_Content_After_Mutation() {
            string md = """
| Header |
| --- |
| value |
""";
            var doc = MarkdownReader.Parse(md);
            var table = Assert.IsType<TableBlock>(doc.Blocks[0]);

            table.Headers[0] = "**Changed**";
            table.Rows[0] = new[] { "[fresh](https://example.com)" };

            Assert.Equal("**Changed**", table.HeaderInlines[0].RenderMarkdown());
            Assert.Single(table.RowInlines);
            Assert.Equal("[fresh](https://example.com)", table.RowInlines[0][0].RenderMarkdown());
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
            Assert.Equal(2, list.Items[0].ParagraphBlocks.Count);
            Assert.Equal("first paragraph", list.Items[0].ParagraphBlocks[0].Inlines.RenderMarkdown());
            Assert.Equal("second paragraph", list.Items[0].ParagraphBlocks[1].Inlines.RenderMarkdown());
            Assert.Collection(
                list.Items[0].BlockChildren,
                block => Assert.Equal("first paragraph", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
                block => Assert.Equal("second paragraph", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()));

            var html = doc.ToHtmlFragment();
            Assert.Contains("<li><p>first paragraph</p><p>second paragraph</p></li>", html);
        }

        [Fact]
        public void Unordered_List_Item_Allows_Lazy_Continuation() {
            const string md = """
- item
continuation
""";

            var doc = MarkdownReader.Parse(md);
            var list = Assert.IsType<UnorderedListBlock>(doc.Blocks[0]);
            Assert.Single(list.Items);

            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("<ul><li>item continuation</li></ul>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Ordered_List_Item_Allows_Lazy_Continuation() {
            const string md = """
1. item
continuation
""";

            var doc = MarkdownReader.Parse(md);
            var list = Assert.IsType<OrderedListBlock>(doc.Blocks[0]);
            Assert.Single(list.Items);

            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("<ol><li>item continuation</li></ol>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Unordered_List_Item_Second_Paragraph_Allows_Lazy_Continuation() {
            const string md = """
- item

    code
after
""";

            var doc = MarkdownReader.Parse(md);
            var list = Assert.IsType<UnorderedListBlock>(doc.Blocks[0]);
            Assert.Single(list.Items);

            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("<ul><li><p>item</p><p>code after</p></li></ul>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Unordered_List_Becomes_Loose_When_Later_Item_Has_Second_Paragraph() {
            const string md = """
- a
- b

  second paragraph
""";

            var doc = MarkdownReader.Parse(md);

            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("<ul><li><p>a</p></li><li><p>b</p><p>second paragraph</p></li></ul>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Ordered_List_Becomes_Loose_When_Later_Item_Has_Second_Paragraph() {
            const string md = """
10. a
11. b

    second paragraph
""";

            var doc = MarkdownReader.Parse(md);

            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("<ol start=\"10\"><li><p>a</p></li><li><p>b</p><p>second paragraph</p></li></ol>", html, StringComparison.Ordinal);
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
        public void List_Item_Can_Contain_Nested_Unordered_List() {
            string md = """
- outer
  - one
  - two
- next
""";
            var doc = MarkdownReader.Parse(md);
            var list = Assert.IsType<UnorderedListBlock>(doc.Blocks[0]);
            Assert.Equal(2, list.Items.Count);
            Assert.Single(list.Items[0].Children);
            Assert.IsType<UnorderedListBlock>(list.Items[0].Children[0]);

            var html = doc.ToHtmlFragment();
            Assert.Contains("<li>outer<ul><li>one</li><li>two</li></ul></li>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Ordered_List_Item_Can_Contain_Nested_Ordered_List() {
            string md = """
1. outer
   1. one
   2. two
2. next
""";
            var doc = MarkdownReader.Parse(md);
            var list = Assert.IsType<OrderedListBlock>(doc.Blocks[0]);
            Assert.Equal(2, list.Items.Count);
            Assert.Single(list.Items[0].Children);
            Assert.IsType<OrderedListBlock>(list.Items[0].Children[0]);

            var html = doc.ToHtmlFragment();
            Assert.Contains("<li>outer<ol><li>one</li><li>two</li></ol></li>", html, StringComparison.Ordinal);
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
        public void List_Item_Can_Contain_BlankLine_Separated_Nested_Blockquote() {
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
            Assert.True(list.Items[0].ForceLoose);

            var html = doc.ToHtmlFragment();
            Assert.Contains("<li><p>outer</p><blockquote>", html, StringComparison.Ordinal);
            Assert.Contains("quote 1", html);
        }

        [Fact]
        public void List_Item_Can_Contain_Nested_Unordered_List_After_Nested_Blockquote() {
            string md = """
- item
  > quote
  continuation
  - nested
""";
            var doc = MarkdownReader.Parse(md);
            var list = Assert.IsType<UnorderedListBlock>(doc.Blocks[0]);
            Assert.Single(list.Items);
            Assert.Contains(list.Items[0].Children, b => b is QuoteBlock);
            Assert.Contains(list.Items[0].Children, b => b is UnorderedListBlock);

            var html = doc.ToHtmlFragment();
            Assert.Contains("<blockquote><p>quote continuation</p></blockquote><ul><li>nested</li></ul>", html, StringComparison.Ordinal);
            Assert.DoesNotContain("<p>- nested</p>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void List_Item_Nested_Blockquote_Can_Lazily_Continue_Within_Tight_List() {
            string md = """
- item
  > quote
  continuation
""";
            var doc = MarkdownReader.Parse(md);
            var list = Assert.IsType<UnorderedListBlock>(doc.Blocks[0]);
            Assert.Single(list.Items);
            var quote = Assert.IsType<QuoteBlock>(list.Items[0].Children[0]);

            var html = doc.ToHtmlFragment();
            Assert.Contains("<li>item<blockquote><p>quote continuation</p></blockquote></li>", html, StringComparison.Ordinal);
            Assert.Single(quote.Children);
        }

        [Fact]
        public void List_Item_Nested_Blockquote_Can_Lazily_Continue_Within_Loose_List() {
            string md = """
- item

  > quote
  continuation
""";
            var doc = MarkdownReader.Parse(md);
            var list = Assert.IsType<UnorderedListBlock>(doc.Blocks[0]);
            Assert.Single(list.Items);
            var quote = Assert.IsType<QuoteBlock>(list.Items[0].Children[0]);

            var html = doc.ToHtmlFragment();
            Assert.Contains("<blockquote><p>quote continuation</p></blockquote>", html, StringComparison.Ordinal);
            Assert.Single(quote.Children);
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
        public void List_Item_Can_Contain_BlankLine_Separated_Nested_Table() {
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
            Assert.True(list.Items[0].ForceLoose);

            var html = doc.ToHtmlFragment();
            Assert.Contains("<li><p>outer</p><table>", html, StringComparison.Ordinal);
            Assert.Contains(">A<", html);
        }

        [Fact]
        public void List_Item_Can_Contain_BlankLine_Separated_Nested_Indented_Code_Block() {
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
            Assert.True(list.Items[0].ForceLoose);

            var html = doc.ToHtmlFragment();
            Assert.Contains("<li><p>outer</p><pre><code>", html, StringComparison.Ordinal);
            Assert.Contains("line1", html);
        }

        [Fact]
        public void List_Item_Keeps_Trailing_Paragraph_After_Nested_Blockquote() {
            string md = """
- item
  > quote

  trailing
""";
            var doc = MarkdownReader.Parse(md);
            var list = Assert.IsType<UnorderedListBlock>(doc.Blocks[0]);
            Assert.Single(list.Items);
            Assert.Equal(2, list.Items[0].ChildBlocks.Count);
            Assert.IsType<QuoteBlock>(list.Items[0].ChildBlocks[0]);
            Assert.IsType<ParagraphBlock>(list.Items[0].ChildBlocks[1]);
            Assert.Collection(
                list.Items[0].BlockChildren,
                block => Assert.Equal("item", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
                block => Assert.IsType<QuoteBlock>(block),
                block => Assert.Equal("trailing", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()));
            Assert.True(list.Items[0].ForceLoose);

            var html = doc.ToHtmlFragment();
            Assert.Contains("<li><p>item</p><blockquote><p>quote</p></blockquote><p>trailing</p></li>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void List_Item_Keeps_Trailing_Paragraph_After_BlankLine_Separated_Nested_Blockquote() {
            string md = """
- item

  > quote

  trailing
""";
            var doc = MarkdownReader.Parse(md);
            var list = Assert.IsType<UnorderedListBlock>(doc.Blocks[0]);
            Assert.Single(list.Items);
            Assert.Equal(2, list.Items[0].ChildBlocks.Count);
            Assert.IsType<QuoteBlock>(list.Items[0].ChildBlocks[0]);
            Assert.IsType<ParagraphBlock>(list.Items[0].ChildBlocks[1]);
            Assert.Collection(
                list.Items[0].BlockChildren,
                block => Assert.Equal("item", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
                block => Assert.IsType<QuoteBlock>(block),
                block => Assert.Equal("trailing", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()));
            Assert.True(list.Items[0].ForceLoose);

            var html = doc.ToHtmlFragment();
            Assert.Contains("<li><p>item</p><blockquote><p>quote</p></blockquote><p>trailing</p></li>", html, StringComparison.Ordinal);
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
            Assert.Contains(list.Items[0].ChildBlocks, b => b is DetailsBlock);

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
        public void Unordered_List_Allows_Tab_Indented_Continuation_Lines() {
            string md = "- first line\n\tsecond line\n- next";
            var doc = MarkdownReader.Parse(md);
            var list = Assert.IsType<UnorderedListBlock>(doc.Blocks[0]);
            Assert.Equal(2, list.Items.Count);
            Assert.Empty(list.Items[0].AdditionalParagraphs);

            var html = doc.ToHtmlFragment();
            Assert.Contains("first line second line", html, StringComparison.Ordinal);
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
        public void Ordered_List_Allows_Tab_Indented_Continuation_Lines() {
            string md = "1. first line\n\tsecond line\n2. next";
            var doc = MarkdownReader.Parse(md);
            var list = Assert.IsType<OrderedListBlock>(doc.Blocks[0]);
            Assert.Equal(2, list.Items.Count);
            Assert.Empty(list.Items[0].AdditionalParagraphs);

            var html = doc.ToHtmlFragment();
            Assert.Contains("first line second line", html, StringComparison.Ordinal);
        }

        [Theory]
        [InlineData("- item\n  heading\n  -------", "<ul><li><h2", ">item heading</h2></li></ul>")]
        [InlineData("1. item\n   heading\n   -------", "<ol><li><h2", ">item heading</h2></li></ol>")]
        public void List_Item_Can_Render_Setext_Heading(string md, string expectedPrefix, string expectedSuffix) {
            var doc = MarkdownReader.Parse(md);

            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains(expectedPrefix, html, StringComparison.Ordinal);
            Assert.Contains(expectedSuffix, html, StringComparison.Ordinal);
        }

        [Fact]
        public void List_Item_Setext_Heading_Parses_Inline_Markup() {
            const string md = """
                - **item**
                  `heading`
                  -------
                """;

            var doc = MarkdownReader.Parse(md);

            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("<li><h2 id=\"item-heading\"><strong>item</strong> <code>heading</code></h2></li>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void List_Item_Setext_Heading_Does_Not_Emit_Empty_Paragraphs_Before_Nested_Blocks() {
            string md = """
- item
  heading
  -------

  > quote
""";
            var doc = MarkdownReader.Parse(md);

            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("<li><h2", html, StringComparison.Ordinal);
            Assert.Contains("<blockquote><p>quote</p></blockquote>", html, StringComparison.Ordinal);
            Assert.DoesNotContain("<p></p>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void List_Item_Setext_Heading_Can_Be_Followed_By_Paragraph_In_Same_Group() {
            const string md = """
- item
  heading
  -------
  after
""";

            var doc = MarkdownReader.Parse(md);

            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("<li><h2", html, StringComparison.Ordinal);
            Assert.Contains(">item heading</h2>after</li>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void List_Item_Blank_Line_Then_Setext_Heading_Starts_A_New_Block() {
            const string md = """
- item

  Heading
  ---
  text
""";

            var doc = MarkdownReader.Parse(md);

            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("<ul><li><p>item</p><h2", html, StringComparison.Ordinal);
            Assert.Contains(">Heading</h2><p>text</p></li></ul>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void NonOne_Ordered_Marker_Does_Not_Interrupt_List_Item_Paragraph() {
            string md = """
- outer
  10. item
      continuation
""";
            var doc = MarkdownReader.Parse(md);
            var list = Assert.IsType<UnorderedListBlock>(doc.Blocks[0]);
            Assert.Single(list.Items);
            Assert.Empty(list.Items[0].Children);

            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("<ul><li>outer 10. item continuation</li></ul>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void List_Item_Nested_Blockquote_Keeps_Lazy_NonOne_Ordered_Continuation_As_Paragraph() {
            string md = """
- outer
  > alpha
  10. beta
      gamma
""";
            var doc = MarkdownReader.Parse(md);
            var list = Assert.IsType<UnorderedListBlock>(doc.Blocks[0]);
            Assert.Single(list.Items);
            var quote = Assert.IsType<QuoteBlock>(list.Items[0].Children[0]);

            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("<blockquote><p>alpha 10. beta gamma</p></blockquote>", html, StringComparison.Ordinal);
            Assert.DoesNotContain("<pre><code>", html, StringComparison.Ordinal);
            Assert.Single(quote.Children);
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
