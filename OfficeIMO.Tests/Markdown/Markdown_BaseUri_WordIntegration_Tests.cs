using System;
using System.Linq;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite {
    public class Markdown_BaseUri_WordIntegration_Tests {
        [Fact]
        public void LoadFromMarkdown_ResolvesRelativeLinks_WithBaseUri() {
            string md = "See [Docs](./guide/index.html).";
            var options = new MarkdownToWordOptions { BaseUri = "https://docs.example.com/" };

            using var doc = md.LoadFromMarkdown(options);

            var link = Assert.Single(doc.HyperLinks);
            Assert.Equal(new Uri("https://docs.example.com/guide/index.html"), link.Uri);
        }

        [Fact]
        public void LoadFromMarkdown_Ignores_NonHttp_BaseUri() {
            string md = "Go [Home](./index.html)";
            var options = new MarkdownToWordOptions { BaseUri = "file:///tmp/docs/" };

            using var doc = md.LoadFromMarkdown(options);

            var link = Assert.Single(doc.HyperLinks);
            // Should remain relative because file:// base is rejected
            Assert.Equal(new Uri("./index.html", UriKind.Relative), link.Uri);
        }
    }
}
