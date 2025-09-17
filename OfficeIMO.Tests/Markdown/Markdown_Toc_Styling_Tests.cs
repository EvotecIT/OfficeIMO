using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite {
    public class Markdown_Toc_Styling_Tests {
        [Fact]
        public void Toc_ScrollSpy_Adds_Nav_And_Script() {
            var md = MarkdownDoc.Create()
                .H1("Doc")
                .H2("Install").P("...")
                .H2("Usage").P("...")
                .TocHere(o => { o.MinLevel = 2; o.MaxLevel = 2; o.Layout = TocLayout.Panel; o.ScrollSpy = true; o.IncludeTitle = true; o.Title = "Contents"; });

            var html = md.ToHtmlDocument(new HtmlOptions { Title = "Doc", Style = HtmlStyle.Clean });

            Assert.Contains("<nav role=\"navigation\"", html);
            Assert.Contains("class=\"md-toc", html);
            Assert.Contains("data-md-scrollspy=\"1\"", html);
            // Our ScrollSpy script uses IntersectionObserver
            Assert.Contains("IntersectionObserver", html);
        }

        [Fact]
        public void Toc_SidebarRight_Renders_Classes() {
            var md = MarkdownDoc.Create()
                .H1("Doc").H2("One").H2("Two")
                .TocHere(o => { o.MinLevel = 2; o.MaxLevel = 3; o.Layout = TocLayout.SidebarRight; o.Sticky = true; o.IncludeTitle = true; o.Title = "On this page"; });
            var html = md.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.GithubAuto });
            Assert.Contains("md-toc sidebar right sticky", html);
            Assert.Contains("On this page", html);
        }
    }
}

