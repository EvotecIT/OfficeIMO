using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite {
    public class Markdown_Reader_IndentedCode_Tests {
        [Fact]
        public void Parses_Indented_Code_Block_As_CodeBlock() {
            string md = """
    line1
    line2

Paragraph
""";
            var doc = MarkdownReader.Parse(md);
            Assert.IsType<CodeBlock>(doc.Blocks[0]);
            Assert.IsType<ParagraphBlock>(doc.Blocks[1]);

            var html = doc.ToHtmlFragment();
            Assert.Contains("<pre><code>", html);
            Assert.Contains("line1", html);
            Assert.Contains("line2", html);
        }

        [Fact]
        public void Indented_Code_Can_Be_Disabled() {
            string md = "    not code";
            var doc = MarkdownReader.Parse(md, new MarkdownReaderOptions { IndentedCodeBlocks = false });
            Assert.IsType<ParagraphBlock>(doc.Blocks[0]);
        }

        [Fact]
        public void Parses_Tab_Indented_Code_Block_As_CodeBlock() {
            string md = "\tline1\n\tline2\n\nParagraph";
            var doc = MarkdownReader.Parse(md);

            var code = Assert.IsType<CodeBlock>(doc.Blocks[0]);
            Assert.Equal("line1\nline2", code.Content.Replace("\r\n", "\n"));
            Assert.IsType<ParagraphBlock>(doc.Blocks[1]);

            var html = doc.ToHtmlFragment();
            Assert.Contains("<pre><code>line1", html, StringComparison.Ordinal);
            Assert.Contains("line2", html, StringComparison.Ordinal);
        }

        [Fact]
        public void FourSpaceIndented_Fence_Is_Not_Parsed_As_Fenced_Code_Block() {
            string md = """
    ```csharp
    var x = 1;
    ```
""";

            var doc = MarkdownReader.Parse(md);

            var block = Assert.IsType<CodeBlock>(Assert.Single(doc.Blocks));
            Assert.False(block.IsFenced);
            Assert.Equal("```csharp\nvar x = 1;\n```", block.Content.Replace("\r\n", "\n"));
        }

        [Fact]
        public void Closing_Fence_Indented_By_Four_Spaces_Does_Not_Close_Fenced_Block() {
            string md = """
```csharp
var x = 1;
    ```
after
""";

            var doc = MarkdownReader.Parse(md);

            var block = Assert.IsType<CodeBlock>(Assert.Single(doc.Blocks));
            Assert.True(block.IsFenced);
            Assert.Equal("csharp", block.Language);
            Assert.Equal("var x = 1;\n    ```\nafter", block.Content.Replace("\r\n", "\n"));
        }
    }
}

