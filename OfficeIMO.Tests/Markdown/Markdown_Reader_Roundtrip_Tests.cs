using System;
using System.Linq;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite {
    public class Markdown_Reader_Roundtrip_Tests {
        [Fact]
        public void Reader_Roundtrips_Basic_Document() {
            var md = MarkdownDoc.Create()
                .FrontMatter(new { title = "Doc", tags = new [] { "a", "b" }, published = true })
                .H1("Doc")
                .P("Intro paragraph with a link to docs.")
                .H2("Install")
                .Code("bash", "dotnet tool install -g Example")
                .Caption("Install the global tool")
                .H2("Features")
                .Ul(ul => ul
                    .Item("Tables from sequences")
                    .Item("Callouts and TOC")
                    .ItemTask("Works offline", true))
                .Image("https://example.com/logo.png", alt: "Logo", title: "Example").Caption("Our logo")
                .Table(t => t.Headers("Col1", "Col2").Row("A", "1").Row("B", "2").Align(ColumnAlignment.Left, ColumnAlignment.Right))
                .Callout("info", "Heads up", "Early access APIs may change.")
                .Dl(d => d.Item("Term", "Definition"));

            var text = md.ToMarkdown();
            var parsed = MarkdownReader.Parse(text);

            // Basic block shape assertions
            Assert.True(parsed.Blocks.Count >= 7);
            Assert.IsType<HeadingBlock>(parsed.Blocks[0] is FrontMatterBlock ? parsed.Blocks[1] : parsed.Blocks[0]);

            // Find code block and validate caption
            var code = parsed.Blocks.OfType<CodeBlock>().FirstOrDefault();
            Assert.NotNull(code);
            Assert.Equal("bash", code!.Language);
            Assert.Equal("Install the global tool", code.Caption);

            // Validate image + caption
            var img = parsed.Blocks.OfType<ImageBlock>().FirstOrDefault();
            Assert.NotNull(img);
            Assert.Contains("logo.png", img!.Path);
            Assert.Equal("Logo", img.Alt);
            Assert.Equal("Example", img.Title);
            Assert.Equal("Our logo", img.Caption);

            // Validate table header/alignments
            var table = parsed.Blocks.OfType<TableBlock>().FirstOrDefault();
            Assert.NotNull(table);
            Assert.Equal(new [] { "Col1", "Col2" }, table!.Headers);
            Assert.Equal(new [] { ColumnAlignment.Left, ColumnAlignment.Right }, table.Alignments);

            // Validate list kinds: unordered with a task item
            var ul = parsed.Blocks.OfType<UnorderedListBlock>().FirstOrDefault();
            Assert.NotNull(ul);
            Assert.True(ul!.Items.Count >= 3);
        }

        [Fact]
        public void Document_Add_Treats_FrontMatter_As_Document_Header_Block() {
            var md = MarkdownDoc.Create()
                .Add(FrontMatterBlock.FromObject(new { title = "Doc", published = true }))
                .H1("Doc");

            var text = md.ToMarkdown().Replace("\r", "");

            Assert.StartsWith("---\n", text);
            Assert.Contains("title: Doc", text);
            Assert.Contains("published: true", text);
            Assert.Contains("\n\n# Doc", text);
            Assert.Single(md.Blocks);
            Assert.IsType<HeadingBlock>(md.Blocks[0]);
        }

        [Fact]
        public void Reader_Exposes_FrontMatter_Entries_As_Structured_Data() {
            const string markdown = """
---
title: Doc
published: true
tags: [a, b]
---

# Heading
""";

            var parsed = MarkdownReader.Parse(markdown);
            var frontMatter = Assert.IsType<FrontMatterBlock>(parsed.DocumentHeader!);

            Assert.Collection(
                frontMatter.Entries,
                entry => {
                    Assert.Equal("title", entry.Key);
                    Assert.Equal("Doc", Assert.IsType<string>(entry.Value));
                },
                entry => {
                    Assert.Equal("published", entry.Key);
                    Assert.True(Assert.IsType<bool>(entry.Value));
                },
                entry => {
                    Assert.Equal("tags", entry.Key);
                    var tags = Assert.IsAssignableFrom<IEnumerable<string>>(entry.Value);
                    Assert.Equal(new[] { "a", "b" }, tags.ToArray());
                });
        }

        [Fact]
        public void Reader_Exposes_TopLevel_Blocks_In_Document_Order() {
            const string markdown = """
---
title: Doc
---

# Heading

Paragraph
""";

            var parsed = MarkdownReader.Parse(markdown);

            Assert.Collection(
                parsed.TopLevelBlocks,
                block => Assert.IsType<FrontMatterBlock>(block),
                block => Assert.IsType<HeadingBlock>(block),
                block => Assert.IsType<ParagraphBlock>(block));

            Assert.Collection(
                parsed.Blocks,
                block => Assert.IsType<HeadingBlock>(block),
                block => Assert.IsType<ParagraphBlock>(block));
        }

        [Fact]
        public void Reader_Enumerates_Blocks_Depth_First() {
            const string markdown = """
---
title: Doc
---

> Quote
>
> - item

Paragraph
""";

            var parsed = MarkdownReader.Parse(markdown);
            var kinds = parsed.DescendantsAndSelf().Select(block => block.GetType()).ToArray();

            Assert.Equal(
                new[] {
                    typeof(FrontMatterBlock),
                    typeof(QuoteBlock),
                    typeof(ParagraphBlock),
                    typeof(UnorderedListBlock),
                    typeof(ParagraphBlock),
                    typeof(ParagraphBlock)
                },
                kinds);
        }

        [Fact]
        public void Reader_Enumerates_List_Items_In_Document_Order() {
            const string markdown = """
- outer
  - inner

> - quoted
""";

            var parsed = MarkdownReader.Parse(markdown);
            var items = parsed.DescendantListItems().ToArray();

            Assert.Equal(3, items.Length);
            Assert.Equal("outer", items[0].Content.RenderMarkdown());
            Assert.Equal("inner", items[1].Content.RenderMarkdown());
            Assert.Equal("quoted", items[2].Content.RenderMarkdown());

            var topList = Assert.IsType<UnorderedListBlock>(parsed.Blocks[0]);
            Assert.Equal(topList.Items, topList.ListItems);

            var quotedList = Assert.IsType<UnorderedListBlock>(Assert.IsType<QuoteBlock>(parsed.Blocks[1]).ChildBlocks[0]);
            Assert.Equal(quotedList.Items, quotedList.ListItems);
        }
    }
}
