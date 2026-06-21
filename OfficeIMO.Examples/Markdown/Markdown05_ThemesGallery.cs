using System;
using System.IO;
using System.Text;
using OfficeIMO.Markdown;

namespace OfficeIMO.Examples.Markdown {
    internal static class Markdown05_ThemesGallery {
        private static (string mdPath, string htmlPath) Paths(string folderPath, string name) {
            string mdFolder = Path.Combine(folderPath, "Markdown");
            Directory.CreateDirectory(mdFolder);
            string mdPath = Path.Combine(mdFolder, name + ".md");
            string htmlPath = Path.ChangeExtension(mdPath, ".html");
            return (mdPath, htmlPath);
        }

        private static MarkdownDoc BaseDoc() => MarkdownDoc.Create()
            .H1("Theme Gallery")
            .Toc(o => { o.MinLevel = 2; o.MaxLevel = 3; o.Layout = TocLayout.Panel; o.Title = "Contents"; }, placeAtTop: true)
            .P("Quick visual check of built-in HtmlStyle presets and a custom theme.")
            .H2("Links & Inline Code").P(p => p.Text("See ").Link("GitHub", "https://github.com/EvotecIT/OfficeIMO").Text(" and ").Code("inline code").Text("."))
            .H2("List")
            .Ul(ul => ul.Item("First").Item("Second").Item("Third"))
            .H2("Zebra Table")
            .Table(t => t.Headers("Col A","Col B").Row("A1","B1").Row("A2","B2").Row("A3","B3"))
            .H2("Callout")
            .Callout("info", "Heads up", "This callout left edge and background reflect the theme accent.")
            .H2("Code Block")
            .Code("csharp", "Console.WriteLine(\"Hello\");");

        public static void Example_Themes(string folderPath, bool open) {
            Console.WriteLine("[*] Markdown: Themes Gallery");
            var baseDoc = BaseDoc();

            // Clean (light)
            var (md1, html1) = Paths(folderPath, "Markdown_Theme_Clean");
            File.WriteAllText(md1, baseDoc.ToMarkdown(), Encoding.UTF8);
            File.WriteAllText(html1, baseDoc.ToHtmlDocument(new HtmlOptions { Title = "Theme: Clean", Style = HtmlStyle.Clean, Kind = HtmlKind.Document, ThemeToggle = true, IncludeAnchorLinks = true, BackToTopLinks = true }), Encoding.UTF8);

            // Word (document-centric)
            var (md1b, html1b) = Paths(folderPath, "Markdown_Theme_Word");
            File.WriteAllText(md1b, baseDoc.ToMarkdown(), Encoding.UTF8);
            File.WriteAllText(html1b, baseDoc.ToHtmlDocument(new HtmlOptions { Title = "Theme: Word", Style = HtmlStyle.Word, Kind = HtmlKind.Document, ThemeToggle = false, IncludeAnchorLinks = true, BackToTopLinks = true }), Encoding.UTF8);

            // GitHub Light (static) — theme toggle off to avoid confusion
            var (md2, html2) = Paths(folderPath, "Markdown_Theme_GithubLight");
            File.WriteAllText(md2, baseDoc.ToMarkdown(), Encoding.UTF8);
            File.WriteAllText(html2, baseDoc.ToHtmlDocument(new HtmlOptions { Title = "Theme: GitHub Light", Style = HtmlStyle.GithubLight, Kind = HtmlKind.Document, ThemeToggle = false, IncludeAnchorLinks = true, BackToTopLinks = true }), Encoding.UTF8);

            // GitHub Dark (static) — theme toggle off to avoid confusion
            var (md3, html3) = Paths(folderPath, "Markdown_Theme_GithubDark");
            File.WriteAllText(md3, baseDoc.ToMarkdown(), Encoding.UTF8);
            File.WriteAllText(html3, baseDoc.ToHtmlDocument(new HtmlOptions { Title = "Theme: GitHub Dark", Style = HtmlStyle.GithubDark, Kind = HtmlKind.Document, ThemeToggle = false, IncludeAnchorLinks = true, BackToTopLinks = true }), Encoding.UTF8);

            // GitHub Auto
            var (md4, html4) = Paths(folderPath, "Markdown_Theme_GithubAuto");
            File.WriteAllText(md4, baseDoc.ToMarkdown(), Encoding.UTF8);
            File.WriteAllText(html4, baseDoc.ToHtmlDocument(new HtmlOptions { Title = "Theme: GitHub Auto", Style = HtmlStyle.GithubAuto, Kind = HtmlKind.Document, ThemeToggle = true, IncludeAnchorLinks = true, BackToTopLinks = true }), Encoding.UTF8);

            // Shared Indigo accents (toggle enabled)
            var (md5, html5) = Paths(folderPath, "Markdown_Theme_Indigo");
            var indigo = MarkdownVisualTheme.Report().WithColorScheme(MarkdownColorSchemeKind.Indigo);
            File.WriteAllText(md5, baseDoc.ToMarkdown(), Encoding.UTF8);
            File.WriteAllText(html5, baseDoc.ToHtmlDocument(new HtmlOptions { Title = "Theme: Indigo", Kind = HtmlKind.Document, ThemeToggle = true, IncludeAnchorLinks = true, BackToTopLinks = true, VisualTheme = indigo }), Encoding.UTF8);

            // Shared Blue accents
            var (md6, html6) = Paths(folderPath, "Markdown_Theme_Blue");
            var blue = MarkdownVisualTheme.Report().WithColorScheme(MarkdownColorSchemeKind.Blue);
            File.WriteAllText(md6, baseDoc.ToMarkdown(), Encoding.UTF8);
            File.WriteAllText(html6, baseDoc.ToHtmlDocument(new HtmlOptions { Title = "Theme: Blue", Kind = HtmlKind.Document, ThemeToggle = true, IncludeAnchorLinks = true, BackToTopLinks = true, VisualTheme = blue }), Encoding.UTF8);

            // Shared Emerald accents
            var (md7, html7) = Paths(folderPath, "Markdown_Theme_Emerald");
            var emerald = MarkdownVisualTheme.Report().WithColorScheme(MarkdownColorSchemeKind.Emerald);
            File.WriteAllText(md7, baseDoc.ToMarkdown(), Encoding.UTF8);
            File.WriteAllText(html7, baseDoc.ToHtmlDocument(new HtmlOptions { Title = "Theme: Emerald", Kind = HtmlKind.Document, ThemeToggle = true, IncludeAnchorLinks = true, BackToTopLinks = true, VisualTheme = emerald }), Encoding.UTF8);

            // Shared Rose accents
            var (md8, html8) = Paths(folderPath, "Markdown_Theme_Rose");
            var rose = MarkdownVisualTheme.Report().WithColorScheme(MarkdownColorSchemeKind.Rose);
            File.WriteAllText(md8, baseDoc.ToMarkdown(), Encoding.UTF8);
            File.WriteAllText(html8, baseDoc.ToHtmlDocument(new HtmlOptions { Title = "Theme: Rose", Kind = HtmlKind.Document, ThemeToggle = true, IncludeAnchorLinks = true, BackToTopLinks = true, VisualTheme = rose }), Encoding.UTF8);

            Console.WriteLine($"✓ HTML (Clean):        {html1}");
            Console.WriteLine($"✓ HTML (Word):        {html1b}");
            Console.WriteLine($"✓ HTML (GitHubLight):  {html2}");
            Console.WriteLine($"✓ HTML (GitHubDark):   {html3}");
            Console.WriteLine($"✓ HTML (GitHubAuto):   {html4}");
            Console.WriteLine($"✓ HTML (Indigo):       {html5}");
            Console.WriteLine($"✓ HTML (Blue):         {html6}");
            Console.WriteLine($"✓ HTML (Emerald):      {html7}");
            Console.WriteLine($"✓ HTML (Rose):         {html8}");
        }
    }
}
