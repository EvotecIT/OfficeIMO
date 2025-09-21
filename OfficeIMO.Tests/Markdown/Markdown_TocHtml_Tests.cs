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

        [Fact]
        public void Toc_Html_Clamps_OutOfRange_Level_Settings() {
            var md = MarkdownDoc.Create()
                .TocHere(o => {
                    o.IncludeTitle = true;
                    o.Title = "Contents";
                    o.MinLevel = 0;
                    o.MaxLevel = 99;
                    o.TitleLevel = 9;
                })
                .H1("Intro")
                .H6("Deep");

            var html = md.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Clean });

            Assert.Contains("<h6>Contents</h6>", html);
            Assert.DoesNotContain("<h9", html);
            Assert.Contains("href=\"#intro\">Intro</a>", html);
            Assert.Contains("href=\"#deep\">Deep</a>", html);
        }

        [Fact]
        public void Toc_Html_Normalizes_OutOfRange_MinAndMaxLevels() {
            var md = MarkdownDoc.Create()
                .TocHere(o => {
                    o.IncludeTitle = false;
                    o.MinLevel = 8;
                    o.MaxLevel = 9;
                    o.RequireTopLevel = false;
                })
                .H6("Deep");

            var html = md.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Clean });

            Assert.Contains("href=\"#deep\">Deep</a>", html);
        }
    }
}

