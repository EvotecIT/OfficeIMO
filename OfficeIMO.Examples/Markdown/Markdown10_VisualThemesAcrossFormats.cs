using System;
using System.IO;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

namespace OfficeIMO.Examples.Markdown {
    internal static class Markdown10_VisualThemesAcrossFormats {
        public static void Example_SharedVisualTheme(string folderPath, bool openWord) {
            Console.WriteLine("[*] Markdown: Shared visual theme across HTML, PDF, and Word");
            string mdFolder = Path.Combine(folderPath, "Markdown");
            Directory.CreateDirectory(mdFolder);

            MarkdownVisualTheme theme = MarkdownVisualTheme.Report()
                .WithColorScheme(MarkdownColorSchemeKind.Emerald)
                .WithColors(accent: "SeaGreen", heading: "#064e3b")
                .WithTable(table => {
                    table.BorderWidth = 0.9;
                    table.CellPaddingX = 8;
                    table.CellPaddingY = 6;
                    table.UseRowStripes = true;
                });

            MarkdownDoc doc = MarkdownDoc.Create()
                .FrontMatter(new { title = "Shared Theme Report", theme = "report", scheme = "emerald" })
                .H1("Shared Theme Report")
                .Toc(o => {
                    o.MinLevel = 2;
                    o.MaxLevel = 3;
                    o.Layout = TocLayout.Panel;
                    o.Title = "Contents";
                }, placeAtTop: true)
                .P("One Markdown AST can now drive HTML, PDF, and Word with a single shared visual theme.")
                .H2("Evidence")
                .Ul(ul => ul
                    .Item("Headings use the same accent family.")
                    .Item("Tables share borders, header colors, padding, and row stripes.")
                    .Item("Code, quotes, callouts, and links follow the same palette."))
                .Table(t => t
                    .Headers("Surface", "Save API", "Theme source")
                    .Row("Markdown", "SaveAsMarkdown", "semantic source")
                    .Row("HTML", "SaveAsHtml", "HtmlOptions.Theme")
                    .Row("PDF", "SaveAsPdf", "MarkdownPdfSaveOptions.Theme")
                    .Row("Word", "ToWordDocument", "MarkdownToWordOptions.Theme"))
                .Callout("success", "Consistent visuals", "The same theme object controls the conversion-specific renderer details.")
                .H2("Code")
                .Code("csharp", """
var theme = MarkdownVisualTheme.Report()
    .WithColorScheme(MarkdownColorSchemeKind.Emerald)
    .WithColors(accent: "SeaGreen", heading: "#064e3b")
    .WithTable(table => table.BorderWidth = 0.9);

doc.SaveAsHtml(htmlPath, new HtmlOptions { Theme = theme });
doc.SaveAsPdf(pdfPath, new MarkdownPdfSaveOptions { Theme = theme });
using var word = doc.ToWordDocument(new MarkdownToWordOptions { Theme = theme });
word.SaveCopy(docxPath);
""");

            string mdPath = Path.Combine(mdFolder, "Markdown_SharedVisualTheme.md");
            string htmlPath = Path.Combine(mdFolder, "Markdown_SharedVisualTheme.html");
            string pdfPath = Path.Combine(mdFolder, "Markdown_SharedVisualTheme.pdf");
            string docxPath = Path.Combine(mdFolder, "Markdown_SharedVisualTheme.docx");

            File.WriteAllText(mdPath, doc.ToMarkdown(), Encoding.UTF8);
            doc.SaveAsHtml(htmlPath, new HtmlOptions {
                Title = "Shared Theme Report",
                Kind = HtmlKind.Document,
                Theme = theme,
                ThemeToggle = true,
                IncludeAnchorLinks = true,
                BackToTopLinks = true
            });
            doc.SaveAsPdf(pdfPath, new MarkdownPdfSaveOptions {
                Theme = theme
            });

            using var word = doc.ToWordDocument(new MarkdownToWordOptions {
                Theme = theme,
                FontFamily = "Aptos"
            });
            word.SaveCopy(docxPath, new WordSaveOptions { OpenAfterSave = openWord });

            Console.WriteLine($"✓ Markdown: {mdPath}");
            Console.WriteLine($"✓ HTML:     {htmlPath}");
            Console.WriteLine($"✓ PDF:      {pdfPath}");
            Console.WriteLine($"✓ Word:     {docxPath}");
        }
    }
}
