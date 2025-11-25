using System.Linq;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite {
    public class Markdown_BaseUri_Tests {
        [Fact]
        public void BaseUri_ResolvesRelativeLinks() {
            var options = new MarkdownReaderOptions { BaseUri = "https://docs.example.com/articles/" };
            var doc = MarkdownReader.Parse("See [Guide](getting-started/index.html).", options);

            var paragraph = Assert.IsType<ParagraphBlock>(doc.Blocks[0]);
            var link = Assert.Single(paragraph.Inlines.Items.OfType<LinkInline>());
            Assert.Equal("https://docs.example.com/articles/getting-started/index.html", link.Url);
        }

        [Fact]
        public void BaseUri_ResolvesRelativeImages() {
            var options = new MarkdownReaderOptions { BaseUri = "https://cdn.example.com/static/" };
            var doc = MarkdownReader.Parse("![Logo](images/logo.png)", options);

            var image = Assert.IsType<ImageBlock>(doc.Blocks[0]);
            Assert.Equal("https://cdn.example.com/static/images/logo.png", image.Path);
        }

        [Fact]
        public void BaseUri_ResolvesReferenceLinks() {
            var md = string.Join("\n", new[] {
                "See [Docs][docs].",
                "",
                "[docs]: ./guide/index.html \"Docs\""
            });
            var options = new MarkdownReaderOptions { BaseUri = "https://docs.example.com/" };
            var doc = MarkdownReader.Parse(md, options);

            var paragraph = Assert.IsType<ParagraphBlock>(doc.Blocks[0]);
            var link = Assert.Single(paragraph.Inlines.Items.OfType<LinkInline>());
            Assert.Equal("https://docs.example.com/guide/index.html", link.Url);
        }

        [Fact]
        public void BaseUri_SkipsAbsoluteUrls() {
            var options = new MarkdownReaderOptions { BaseUri = "https://docs.example.com/base/" };
            var doc = MarkdownReader.Parse("Visit [Site](https://example.net/x) and [Cdn](//cdn.example.net/app.js) and [Mail](mailto:hello@example.net) and [Fragment](#intro).", options);

            var urls = doc.Blocks
                .OfType<ParagraphBlock>()
                .SelectMany(p => p.Inlines.Items.OfType<LinkInline>())
                .Select(l => l.Url)
                .ToList();

            Assert.Contains("https://example.net/x", urls);
            Assert.Contains("//cdn.example.net/app.js", urls);
            Assert.Contains("mailto:hello@example.net", urls);
            Assert.Contains("#intro", urls);
        }

        [Fact]
        public void BaseUri_SkipsDataUrls() {
            var options = new MarkdownReaderOptions { BaseUri = "https://assets.example.com/" };
            var doc = MarkdownReader.Parse("![Inline](data:image/png;base64,AAA)", options);

            var image = Assert.IsType<ImageBlock>(doc.Blocks[0]);
            Assert.Equal("data:image/png;base64,AAA", image.Path);
        }

        [Fact]
        public void BaseUri_AllowsEmptyHref() {
            var options = new MarkdownReaderOptions { BaseUri = "https://docs.example.com/base/" };
            var doc = MarkdownReader.Parse("[Empty]()", options);

            var paragraph = Assert.IsType<ParagraphBlock>(doc.Blocks[0]);
            var link = Assert.Single(paragraph.Inlines.Items.OfType<LinkInline>());
            Assert.Equal(string.Empty, link.Url);
        }

        [Fact]
        public void BaseUri_HandlesInvalidUris() {
            var options = new MarkdownReaderOptions { BaseUri = "http://exa mple.com" }; // invalid base should be ignored
            var doc = MarkdownReader.Parse("![Photo](images/photo.png)", options);

            var image = Assert.IsType<ImageBlock>(doc.Blocks[0]);
            Assert.Equal("images/photo.png", image.Path);
        }
    }
}
