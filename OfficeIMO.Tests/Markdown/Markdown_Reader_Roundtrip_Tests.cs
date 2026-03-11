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

            var published = frontMatter.FindEntry("published");
            Assert.NotNull(published);
            Assert.True(Assert.IsType<bool>(published!.Value));

            Assert.True(frontMatter.TryGetValue<string>("title", out var title));
            Assert.Equal("Doc", title);

            Assert.True(frontMatter.TryGetValue<bool>("published", out var isPublished));
            Assert.True(isPublished);

            Assert.True(frontMatter.TryGetValue<IEnumerable<string>>("tags", out var tagValues));
            Assert.Equal(new[] { "a", "b" }, tagValues!.ToArray());

            Assert.False(frontMatter.TryGetValue<int>("title", out _));
            Assert.Null(frontMatter.FindEntry("missing"));

            Assert.True(parsed.HasDocumentHeader);
            Assert.Equal("published", parsed.FindFrontMatterEntry("published")!.Key);
            Assert.True(parsed.TryGetFrontMatterValue<string>("title", out var documentTitle));
            Assert.Equal("Doc", documentTitle);
            Assert.False(parsed.TryGetFrontMatterValue<int>("title", out _));
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

            Assert.Collection(
                parsed.TopLevelBlocksOfType<HeadingBlock>(),
                block => Assert.Equal("Heading", block.Text));

            Assert.Collection(
                parsed.TopLevelBlocksOfType<FrontMatterBlock>(),
                block => Assert.Equal("Doc", block.Entries[0].Value));

            var topLevelHeading = parsed.FindFirstTopLevelBlockOfType<HeadingBlock>();
            Assert.NotNull(topLevelHeading);
            Assert.Equal("Heading", topLevelHeading!.Text);
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

            Assert.Equal(
                new[] { "Quote", "item", "Paragraph" },
                parsed.DescendantsOfType<ParagraphBlock>()
                    .Select(block => block.Inlines.RenderMarkdown())
                    .ToArray());

            var firstNestedList = parsed.FindFirstDescendantOfType<UnorderedListBlock>();
            Assert.NotNull(firstNestedList);
            Assert.Equal("item", firstNestedList!.Items[0].Content.RenderMarkdown());
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

        [Fact]
        public void Reader_Enumerates_Headings_And_Resolved_Anchors() {
            const string markdown = """
# Title

## Repeat

## Repeat
""";

            var parsed = MarkdownReader.Parse(markdown);
            var headings = parsed.DescendantHeadings().ToArray();

            Assert.Equal(new[] { "Title", "Repeat", "Repeat" }, headings.Select(h => h.Text).ToArray());
            Assert.Equal("title", parsed.GetHeadingAnchor(headings[0]));
            Assert.Equal("repeat", parsed.GetHeadingAnchor(headings[1]));
            Assert.Equal("repeat-1", parsed.GetHeadingAnchor(headings[2]));

            var infos = parsed.GetHeadingInfos();
            Assert.Equal(new[] { "Title", "Repeat", "Repeat" }, infos.Select(info => info.Text).ToArray());
            Assert.Equal(new[] { "title", "repeat", "repeat-1" }, infos.Select(info => info.Anchor).ToArray());
            Assert.Equal(new[] { 1, 2, 2 }, infos.Select(info => info.Level).ToArray());
            Assert.Same(headings[1], infos[1].Block);

            var byAnchor = parsed.FindHeadingByAnchor("#repeat-1");
            Assert.NotNull(byAnchor);
            Assert.Equal("Repeat", byAnchor!.Text);
            Assert.Equal("repeat-1", byAnchor.Anchor);

            Assert.Null(parsed.FindHeadingByAnchor("missing"));

            var byText = parsed.FindHeadings("repeat");
            Assert.Equal(2, byText.Count);
            Assert.Equal(new[] { "repeat", "repeat-1" }, byText.Select(info => info.Anchor).ToArray());
        }
    }
}
