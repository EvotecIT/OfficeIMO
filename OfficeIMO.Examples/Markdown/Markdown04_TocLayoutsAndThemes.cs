using System;
using System.IO;
using System.Text;
using OfficeIMO.Markdown;

namespace OfficeIMO.Examples.Markdown {
    internal static class Markdown04_TocLayoutsAndThemes {
        private static (string mdPath, string htmlPath) Paths(string folderPath, string name) {
            string mdFolder = Path.Combine(folderPath, "Markdown");
            Directory.CreateDirectory(mdFolder);
            string mdPath = Path.Combine(mdFolder, name + ".md");
            string htmlPath = Path.ChangeExtension(mdPath, ".html");
            return (mdPath, htmlPath);
        }

        private static MarkdownDoc BaseDoc() => MarkdownDoc.Create()
            .H1("OfficeIMO.Markdown — TOC Demos")
            .P("Demonstrates panel and sidebar TOCs with optional ScrollSpy and themes.")
            .H2("Install").P("Installation instructions...")
            .H2("Usage").H3("Tables").P("Table helpers...").H3("Lists").P("List helpers...")
            .H2("Advanced").H3("Prism").P("Code highlighting...").H3("Assets").P("Online/Offline...")
            .H2("FAQ").P("Common questions...")
            .H2("Appendix").H3("Extra").P("Additional material...");

        public static void Example_Toc_PanelTop(string folderPath, bool open) {
            Console.WriteLine("[*] Markdown: TOC Panel (top)");
            var (mdPath, htmlPath) = Paths(folderPath, "Markdown_Toc_PanelTop");
            var md = BaseDoc();
            md.Toc(o => { o.MinLevel = 1; o.MaxLevel = 3; o.Layout = TocLayout.Panel; o.IncludeTitle = true; o.Title = "On this page"; }, placeAtTop: true);
            File.WriteAllText(mdPath, md.ToMarkdown(), Encoding.UTF8);
            var html = md.ToHtmlDocument(new HtmlOptions { Title = "TOC Panel", Style = HtmlStyle.Word, Kind = HtmlKind.Document, ThemeToggle = true, IncludeAnchorLinks = true, BackToTopLinks = true });
            File.WriteAllText(htmlPath, html, Encoding.UTF8);
            Console.WriteLine($"✓ HTML: {htmlPath}");
        }

        public static void Example_Toc_SidebarLeft(string folderPath, bool open) {
            Console.WriteLine("[*] Markdown: TOC Sidebar Left (ScrollSpy, matched to right)");
            var (mdPath, htmlPath) = Paths(folderPath, "Markdown_Toc_SidebarLeft");
            var md = BaseDoc();
            md = MarkdownDoc.Create().Toc(o => {
                o.MinLevel = 1; o.MaxLevel = 3;
                o.Layout = TocLayout.SidebarLeft;
                o.Sticky = true;
                o.ScrollSpy = true;
                o.Chrome = TocChrome.Outline;
                o.IncludeTitle = true; o.Title = "On this page";
                o.WidthPx = 260;
                o.HideOnNarrow = true;
            }, placeAtTop: true);
            foreach (var b in BaseDoc().Blocks) md.Add(b);
            File.WriteAllText(mdPath, md.ToMarkdown(), Encoding.UTF8);
            var html = md.ToHtmlDocument(new HtmlOptions { Title = "TOC Left", Style = HtmlStyle.Word, Kind = HtmlKind.Document, ThemeToggle = true, IncludeAnchorLinks = true, BackToTopLinks = true });
            File.WriteAllText(htmlPath, html, Encoding.UTF8);
            Console.WriteLine($"✓ HTML: {htmlPath}");
        }

        public static void Example_Toc_SidebarRight_ScrollSpy(string folderPath, bool open) {
            Console.WriteLine("[*] Markdown: TOC Sidebar Right (ScrollSpy)");
            var (mdPath, htmlPath) = Paths(folderPath, "Markdown_Toc_SidebarRight_ScrollSpy");
            var baseDoc = BaseDoc();
            var md = MarkdownDoc.Create();
            md.Toc(o => { o.MinLevel = 2; o.MaxLevel = 3; o.Layout = TocLayout.SidebarRight; o.Sticky = true; o.ScrollSpy = true; o.IncludeTitle = true; o.Title = "On this page"; }, placeAtTop: true);
            foreach (var b in baseDoc.Blocks) md.Add(b);
            File.WriteAllText(mdPath, md.ToMarkdown(), Encoding.UTF8);
            var html = md.ToHtmlDocument(new HtmlOptions { Title = "TOC Right + ScrollSpy", Style = HtmlStyle.Word, Kind = HtmlKind.Document, ThemeToggle = true, IncludeAnchorLinks = true, BackToTopLinks = true });
            File.WriteAllText(htmlPath, html, Encoding.UTF8);
            Console.WriteLine($"✓ HTML: {htmlPath}");
        }

        public static void Example_Toc_ScrollSpy_Long_IndigoTheme(string folderPath, bool open) {
            Console.WriteLine("[*] Markdown: ScrollSpy Long (Indigo theme)");
            var (mdPath, htmlPath) = Paths(folderPath, "Markdown_Toc_ScrollSpy_Indigo");
            var md = MarkdownDoc.Create();
            md.Toc(o => { o.MinLevel = 2; o.MaxLevel = 3; o.Layout = TocLayout.SidebarLeft; o.Sticky = true; o.ScrollSpy = true; o.IncludeTitle = true; o.Title = "On this page"; }, placeAtTop: true)
              .H1("ScrollSpy Demo");
            for (int i = 1; i <= 8; i++) {
                md.H2($"Section {i}");
                for (int j = 1; j <= 3; j++) {
                    md.H3($"Subsection {i}.{j}");
                    md.P(new string('x', 1400));
                }
            }
            File.WriteAllText(mdPath, md.ToMarkdown(), Encoding.UTF8);
            var indigo = new ThemeColors {
                AccentLight = "#4f46e5", AccentDark = "#8b9cfb",
                HeadingLight = "#111827", HeadingDark = "#e5e7eb",
                TocBgLight = "#eef2ff", TocBorderLight = "#c7d2fe",
                TocBgDark = "#1f2937", TocBorderDark = "#374151",
                ActiveLinkLight = "#4338ca", ActiveLinkDark = "#a5b4fc"
            };
            var html = md.ToHtmlDocument(new HtmlOptions { Title = "ScrollSpy Indigo", Style = HtmlStyle.Word, Kind = HtmlKind.Document, ThemeToggle = true, IncludeAnchorLinks = true, BackToTopLinks = true, Theme = indigo });
            File.WriteAllText(htmlPath, html, Encoding.UTF8);
            Console.WriteLine($"✓ HTML: {htmlPath}");
        }
    }
}
