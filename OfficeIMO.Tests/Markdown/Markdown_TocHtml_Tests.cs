using System.Text.RegularExpressions;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite {
    public class Markdown_TocHtml_Tests {
        [Fact]
        public void Toc_Html_NestedLists_AreInsideParentLi() {
            var md = MarkdownDoc.Create()
                .H2("H2")
                .TocHere(o => { o.MinLevel = 2; o.MaxLevel = 6; })
                .H3("H3");

            var html = md.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Clean });
            var normalized = Regex.Replace(html, "\\s+", "");

            // Correct structure: nested <ul> appears before closing the parent <li>
            Assert.Contains("<li><ahref=\"#h2\">H2</a><ul>", normalized);

            // Broken structure must not occur: closing </li> immediately followed by nested <ul>
            Assert.DoesNotContain("</li><ul>", normalized);
        }
    }
}

