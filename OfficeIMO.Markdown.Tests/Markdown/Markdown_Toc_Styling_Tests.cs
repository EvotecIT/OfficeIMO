using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite {
    public class Markdown_Toc_Styling_Tests {
        [Fact]
        public void Toc_ScrollSpy_Adds_Nav_And_Script() {
            var md = MarkdownDocument.Create()
                .H1("Doc")
                .H2("Install").P("...")
                .H2("Usage").P("...")
                .TocHere(o => { o.MinLevel = 2; o.MaxLevel = 2; o.Layout = MarkdownTocLayout.Panel; o.ScrollSpy = true; o.IncludeTitle = true; o.Title = "Contents"; });

            var html = md.ToHtml(new MarkdownHtmlExportOptions { Title = "Doc", Style = MarkdownHtmlStyle.Clean, Kind = global::OfficeIMO.Markdown.MarkdownHtmlKind.Document });

            Assert.Contains("<nav role=\"navigation\"", html);
            Assert.Contains("class=\"md-toc", html);
            Assert.Contains("data-md-scrollspy=\"1\"", html);
            // Our ScrollSpy script uses IntersectionObserver
            Assert.Contains("IntersectionObserver", html);
        }

        [Fact]
        public void Toc_SidebarRight_Renders_Classes() {
            var md = MarkdownDocument.Create()
                .H1("Doc").H2("One").H2("Two")
                .TocHere(o => { o.MinLevel = 2; o.MaxLevel = 3; o.Layout = MarkdownTocLayout.SidebarRight; o.Sticky = true; o.IncludeTitle = true; o.Title = "On this page"; });
            var html = md.ToHtml(new MarkdownHtmlExportOptions { Style = MarkdownHtmlStyle.GithubAuto, Kind = global::OfficeIMO.Markdown.MarkdownHtmlKind.Fragment });
            Assert.Contains("md-toc sidebar right sticky", html);
            Assert.Contains("On this page", html);
        }

        [Fact]
        public void Toc_SidebarRight_Keeps_Footnotes_In_Content_Column() {
            var markdown = """
# Doc

See note[^1].

## One

[TOC layout=sidebar-right min=2 max=2 title="On this page"]

[^1]: Footnote text
""";

            var doc = OfficeIMO.Markdown.MarkdownDocument.Parse(markdown);
            var html = doc.ToHtml(new MarkdownHtmlExportOptions { Style = MarkdownHtmlStyle.Clean, Kind = global::OfficeIMO.Markdown.MarkdownHtmlKind.Fragment });

            Assert.Contains("class=\"md-layout two-col right\"", html);
            Assert.Contains("class=\"md-content\"", html);
            Assert.Contains("class=\"footnotes\"", html);
            Assert.Contains("Footnote text", html);
        }
    }
}

