using System;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite {
    public class Markdown_CodeBlock_Fence_Tests {
        [Fact]
        public void CodeBlock_Expands_Fence_For_Inner_Backticks() {
            var snippet = "echo ```value````";
            var expectedFence = MarkdownFence.BuildSafeFence(snippet);
            var md = MarkdownDoc.Create().Code("bash", snippet);

            var lines = md.ToMarkdown().Replace("\r", string.Empty)
                .Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);

            Assert.StartsWith($"{expectedFence}bash", lines[0]);
            Assert.Contains(snippet, lines);
            Assert.Equal(expectedFence, lines[lines.Length - 1]);
        }

        [Fact]
        public void CodeBlock_Handles_Backticks_With_Trailing_Spaces() {
            var snippet = "```   ";
            var expectedFence = MarkdownFence.BuildSafeFence(snippet);
            var md = MarkdownDoc.Create().Code("text", snippet);

            var lines = md.ToMarkdown().Replace("\r", string.Empty)
                .Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);

            Assert.StartsWith($"{expectedFence}text", lines[0]);
            Assert.Equal(snippet, lines[1]);
            Assert.Equal(expectedFence, lines[2]);
        }

        [Fact]
        public void CodeBlock_RoundTrips_With_Adaptive_Fence_Lengths() {
            var snippet = "echo ```value````"; // contains a run of 4 backticks
            var original = MarkdownDoc.Create().Code("bash", snippet);

            var markdown = original.ToMarkdown();
            var parsed = MarkdownReader.Parse(markdown);

            var block = Assert.Single(parsed.Blocks);
            var code = Assert.IsType<CodeBlock>(block);
            Assert.Equal("bash", code.Language);
            Assert.Equal(snippet, code.Content);
            Assert.Equal(markdown, parsed.ToMarkdown());
        }
    }
}
