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
    }
}

