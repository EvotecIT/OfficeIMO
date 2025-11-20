using System;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite {
    public class Markdown_CodeBlock_Fence_Tests {
        [Fact]
        public void CodeBlock_Expands_Fence_For_Inner_Backticks() {
            var snippet = "echo ```value````";
            var md = MarkdownDoc.Create().Code("bash", snippet);

            var lines = md.ToMarkdown().Replace("\r", string.Empty)
                .Split('\n', StringSplitOptions.RemoveEmptyEntries);

            Assert.StartsWith("`````bash", lines[0]);
            Assert.Contains(snippet, lines);
            Assert.Equal("`````", lines[^1]);
        }

        [Fact]
        public void CodeBlock_Handles_Backticks_With_Trailing_Spaces() {
            var snippet = "```   ";
            var md = MarkdownDoc.Create().Code("text", snippet);

            var lines = md.ToMarkdown().Replace("\r", string.Empty)
                .Split('\n', StringSplitOptions.RemoveEmptyEntries);

            Assert.StartsWith("````text", lines[0]);
            Assert.Equal(snippet, lines[1]);
            Assert.Equal("````", lines[2]);
        }
    }
}
