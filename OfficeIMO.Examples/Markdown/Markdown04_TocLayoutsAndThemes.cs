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

        private static MarkdownDocument BaseDoc() => MarkdownDocument.Create()
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
            md.Toc(o => { o.MinLevel = 1; o.MaxLevel = 3; o.Layout = MarkdownTocLayout.Panel; o.IncludeTitle = true; o.Title = "On this page"; }, placeAtTop: true);
            File.WriteAllText(mdPath, md.ToMarkdown(), Encoding.UTF8);
            var html = md.ToHtml(new MarkdownHtmlExportOptions { Title = "TOC Panel", Style = MarkdownHtmlStyle.Word, Kind = MarkdownHtmlKind.Document, ThemeToggle = true, IncludeAnchorLinks = true, BackToTopLinks = true });
            File.WriteAllText(htmlPath, html, Encoding.UTF8);
            Console.WriteLine($"✓ HTML: {htmlPath}");
        }

        public static void Example_Toc_SidebarLeft(string folderPath, bool open) {
            Console.WriteLine("[*] Markdown: TOC Sidebar Left (ScrollSpy, matched to right)");
            var (mdPath, htmlPath) = Paths(folderPath, "Markdown_Toc_SidebarLeft");
            var md = BaseDoc();
            md = MarkdownDocument.Create().Toc(o => {
                o.MinLevel = 1; o.MaxLevel = 3;
                o.Layout = MarkdownTocLayout.SidebarLeft;
                o.Sticky = true;
                o.ScrollSpy = true;
                o.Chrome = MarkdownTocChrome.Outline;
                o.IncludeTitle = true; o.Title = "On this page";
                o.WidthPx = 260;
                o.HideOnNarrow = true;
            }, placeAtTop: true);
            foreach (var b in BaseDoc().Blocks) md.Add(b);
            File.WriteAllText(mdPath, md.ToMarkdown(), Encoding.UTF8);
            var html = md.ToHtml(new MarkdownHtmlExportOptions { Title = "TOC Left", Style = MarkdownHtmlStyle.Word, Kind = MarkdownHtmlKind.Document, ThemeToggle = true, IncludeAnchorLinks = true, BackToTopLinks = true });
            File.WriteAllText(htmlPath, html, Encoding.UTF8);
            Console.WriteLine($"✓ HTML: {htmlPath}");
        }

        public static void Example_Toc_SidebarRight_ScrollSpy(string folderPath, bool open) {
            Console.WriteLine("[*] Markdown: TOC Sidebar Right (ScrollSpy)");
            var (mdPath, htmlPath) = Paths(folderPath, "Markdown_Toc_SidebarRight_ScrollSpy");
            var baseDoc = BaseDoc();
            var md = MarkdownDocument.Create();
            md.Toc(o => { o.MinLevel = 2; o.MaxLevel = 3; o.Layout = MarkdownTocLayout.SidebarRight; o.Sticky = true; o.ScrollSpy = true; o.IncludeTitle = true; o.Title = "On this page"; }, placeAtTop: true);
            foreach (var b in baseDoc.Blocks) md.Add(b);
            File.WriteAllText(mdPath, md.ToMarkdown(), Encoding.UTF8);
            var html = md.ToHtml(new MarkdownHtmlExportOptions { Title = "TOC Right + ScrollSpy", Style = MarkdownHtmlStyle.Word, Kind = MarkdownHtmlKind.Document, ThemeToggle = true, IncludeAnchorLinks = true, BackToTopLinks = true });
            File.WriteAllText(htmlPath, html, Encoding.UTF8);
            Console.WriteLine($"✓ HTML: {htmlPath}");
        }

        public static void Example_Toc_ScrollSpy_Long_IndigoTheme(string folderPath, bool open) {
            Console.WriteLine("[*] Markdown: ScrollSpy Long (Indigo theme)");
            var (mdPath, htmlPath) = Paths(folderPath, "Markdown_Toc_ScrollSpy_Indigo");
            var md = MarkdownDocument.Create();
            md.Toc(o => { o.MinLevel = 2; o.MaxLevel = 3; o.Layout = MarkdownTocLayout.SidebarLeft; o.Sticky = true; o.ScrollSpy = true; o.IncludeTitle = true; o.Title = "On this page"; }, placeAtTop: true)
              .H1("ScrollSpy Demo");
            for (int i = 1; i <= 8; i++) {
                md.H2($"Section {i}");
                for (int j = 1; j <= 3; j++) {
                    md.H3($"Subsection {i}.{j}");
                    md.P(new string('x', 1400));
                }
            }
            File.WriteAllText(mdPath, md.ToMarkdown(), Encoding.UTF8);
            var indigo = MarkdownVisualTheme.WordLike().WithColorScheme(MarkdownColorSchemeKind.Indigo);
            var html = md.ToHtml(new MarkdownHtmlExportOptions { Title = "ScrollSpy Indigo", Kind = MarkdownHtmlKind.Document, ThemeToggle = true, IncludeAnchorLinks = true, BackToTopLinks = true, Theme = indigo });
            File.WriteAllText(htmlPath, html, Encoding.UTF8);
            Console.WriteLine($"✓ HTML: {htmlPath}");
        }
    }
}
