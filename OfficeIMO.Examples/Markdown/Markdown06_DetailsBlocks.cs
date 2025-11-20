using System;
using System.IO;
using OfficeIMO.Markdown;

namespace OfficeIMO.Examples.Markdown {
    internal static class Markdown06_DetailsBlocks {
        public static void Example_Details_Block(string folderPath, bool open) {
            Console.WriteLine("[*] Markdown details block");

            string mdFolder = Path.Combine(folderPath, "Markdown");
            Directory.CreateDirectory(mdFolder);
            string path = Path.Combine(mdFolder, "Details.md");

            MarkdownDoc md = MarkdownDoc.Create()
                .H1("Collapsible sections")
                .Details("Show more", body => {
                    body.P("This paragraph is hidden by default.");
                    body.Ul(list => list
                        .Item("Supports paragraphs and lists")
                        .Item("Content is parsed and rendered to GitHub-friendly HTML"));
                });

            File.WriteAllText(path, md.ToMarkdown());
            Console.WriteLine($"✓ Markdown saved: {path}");
        }
    }
}
