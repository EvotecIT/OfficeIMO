using System;
using System.IO;
using System.Text;
using OfficeIMO.Markdown;

namespace OfficeIMO.Examples.Markdown {
    internal static class Markdown03_Anchors_Theme {
        public static void Example_AnchorsAndTheme(string folderPath, bool open) {
            Console.WriteLine("[*] Markdown: Anchors, Back-to-Top, Theme Toggle");
            string mdFolder = Path.Combine(folderPath, "Markdown");
            Directory.CreateDirectory(mdFolder);
            string mdPath = Path.Combine(mdFolder, "AnchorsAndTheme.md");
            string htmlPath = Path.ChangeExtension(mdPath, ".html");

            var md = MarkdownDoc.Create()
                .FrontMatter(new { title = "Markdown Anchors & Theme" })
                .H1("Report Title")
                // Collapsible TOC at top (HTML only); avoid duplicate heading by setting IncludeTitle=false
                .Toc(opts => { opts.MinLevel = 1; opts.MaxLevel = 3; opts.Ordered = false; opts.IncludeTitle = false; opts.Collapsible = true; opts.Collapsed = false; }, placeAtTop: true)
                .H2("Intro").P(p => p
                    .Text("This paragraph demonstrates ")
                    .Bold("bold")
                    .Text(", ")
                    .Italic("italic")
                    .Text(", and ")
                    .Underline("underlined")
                    .Text(" text. Here's a ")
                    .Link("link", "https://evotec.xyz", "Evotec")
                    .Text("."))
                .H2("Section A").P("Lorem ipsum dolor sit amet.")
                .H3("Section A.1").P("Subsection with anchor icon.")
                .H2("Section B").P("Another section.")
                .H3("Section B.1").P("Try the theme toggle in the top-right.")
                .H3("Links in Tables")
                .Table(t => t
                    .Headers("Name","URL")
                    .Row("OfficeIMO", "[GitHub](https://github.com/EvotecIT/OfficeIMO)")
                    .Row("DomainDetective", "[Docs](https://evotec.xyz/hub/)")
                );

            File.WriteAllText(mdPath, md.ToMarkdown(), Encoding.UTF8);

            var html = md.ToHtmlDocument(new HtmlOptions {
                Title = "Markdown Anchors & Theme",
                Style = HtmlStyle.GithubAuto,
                CssDelivery = CssDelivery.Inline,
                ShowAnchorIcons = true,
                AnchorIcon = "🔗",
                CopyHeadingLinkOnClick = true,
                BackToTopLinks = true,
                BackToTopMinLevel = 2,
                BackToTopText = "Back to top",
                ThemeToggle = true
            });
            File.WriteAllText(htmlPath, html, Encoding.UTF8);
            Console.WriteLine($"✓ Markdown saved: {mdPath}");
            Console.WriteLine($"✓ HTML saved:     {htmlPath}");
        }
    }
}
