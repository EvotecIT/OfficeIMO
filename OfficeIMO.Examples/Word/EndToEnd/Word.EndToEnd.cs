using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Markdown;
using OfficeIMO.Word.Fluent;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Examples.Word.EndToEnd {
    internal static class Word_EndToEnd {
        public static void Example(string folderPath, bool openWord) {
            Console.WriteLine("[*] Word ⇄ Markdown ⇄ HTML — End-to-End" );
            string outDir = Path.Combine(folderPath, "Word", "EndToEnd");
            Directory.CreateDirectory(outDir);

            // 1) Build a small Word document
            using (var doc = WordDocument.Create()) {
                // Headers/Footers + page numbering
                doc.AddHeadersAndFooters();
                doc.Header.Default.AddParagraph("End-to-End Demo");
                doc.Footer.Default.AddParagraph().AddPageNumber(includeTotalPages: true);

                // TOC at top (updates on open)
                new WordFluentDocument(doc).TocAtTop("Contents", minLevel: 1, maxLevel: 3, titleLevel: 2);

                // Large multi-section body (~10+ pages depending on printer metrics)
                string lorem = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed non risus. Suspendisse lectus tortor, dignissim sit amet, adipiscing nec, ultricies sed, dolor. Cras elementum ultrices diam. Maecenas ligula massa, varius a, semper congue, euismod non, mi.";

                for (int ch = 1; ch <= 6; ch++) {
                    var h1 = doc.AddParagraph($"Chapter {ch}");
                    h1.Style = WordParagraphStyles.Heading1;
                    if (ch > 1) h1.PageBreakBefore = true;

                    for (int sec = 1; sec <= 3; sec++) {
                        doc.AddParagraph($"Topic {ch}.{sec}").Style = WordParagraphStyles.Heading2;
                        // a few paragraphs per topic
                        for (int p = 0; p < 4; p++) doc.AddParagraph(lorem);
                        // Demonstrate lists: bulleted + nested numbered under this topic
                        var list = doc.AddList(WordListStyle.Bulleted);
                        list.AddItem($"Point A{ch}{sec}");
                        list.AddItem($"Point B{ch}{sec}");
                        list.AddItem($"Point C{ch}{sec}");
                        // Nested numbered list under the last bullet for variety
                        var fluent = new WordFluentDocument(doc);
                        fluent.List(l => l.Numbered()
                            .Item($"Nested 1 for {ch}.{sec}")
                            .Indent()
                            .Item($"Nested 1.1 {ch}.{sec}")
                            .Item($"Nested 1.2 {ch}.{sec}")
                            .Outdent()
                            .Item($"Nested 2 for {ch}.{sec}"));
                        // a table
                        var t = doc.AddTable(6, 3);
                        t.Rows[0].Cells[0].Paragraphs[0].Text = "Col1";
                        t.Rows[0].Cells[1].Paragraphs[0].Text = "Col2";
                        t.Rows[0].Cells[2].Paragraphs[0].Text = "Col3";
                        for (int r = 1; r < 6; r++) {
                            for (int c = 0; c < 3; c++) t.Rows[r].Cells[c].Paragraphs[0].Text = $"R{r}C{c+1}";
                        }
                    }
                    // New section every 2 chapters
                    if (ch % 2 == 0) doc.AddSection(SectionMarkValues.NextPage);
                }

                // 2) Save as .docx
                string docx = Path.Combine(outDir, "EndToEnd.docx");
                doc.Save(docx);
                Console.WriteLine($"✓ Word: {docx}");

                // 3) Export to Markdown
                string mdPath = Path.Combine(outDir, "EndToEnd.md");
                doc.SaveAsMarkdown(mdPath, new WordToMarkdownOptions());
                Console.WriteLine($"✓ Markdown: {mdPath}");

                // 4) Export to styled HTML (via Markdown) with TOC at top using simple options
                string htmlPath = Path.Combine(outDir, "EndToEnd.html");
                doc.SaveAsHtmlViaMarkdown(htmlPath, new HtmlOptions {
                    Style = HtmlStyle.Word,
                    Title = "End-to-End Demo",
                    IncludeAnchorLinks = true,
                    BackToTopLinks = true,
                    BackToTopMinLevel = 2,
                    InjectTocAtTop = true,
                    InjectTocTitle = "Contents",
                    InjectTocMinLevel = 1,
                    InjectTocMaxLevel = 3,
                    InjectTocOrdered = false,
                    InjectTocTitleLevel = 2
                });
                Console.WriteLine($"✓ Styled HTML (via Markdown+Word style): {htmlPath}");
            }

            // 5) Load the generated Markdown back into Word
            string mdIn = Path.Combine(outDir, "EndToEnd.md");
            string fromMdDocx = Path.Combine(outDir, "EndToEnd.FromMarkdown.docx");
            using (var fromMarkdown = File.ReadAllText(mdIn).LoadFromMarkdown(new MarkdownToWordOptions { FontFamily = "Calibri" })) {
                fromMarkdown.Save(fromMdDocx);
                Console.WriteLine($"✓ Markdown → Word: {fromMdDocx}");
            }

            // 6) Quick sanity: convert the round-tripped doc back to Markdown/HTML
            using (var round = WordDocument.Load(fromMdDocx)) {
                string md2 = round.ToMarkdown();
                File.WriteAllText(Path.Combine(outDir, "EndToEnd.RoundTrip.md"), md2);
                string html2 = round.ToHtml();
                File.WriteAllText(Path.Combine(outDir, "EndToEnd.RoundTrip.html"), html2);
                Console.WriteLine("✓ Round-trip: Word ⇄ Markdown and Word ⇄ HTML written");
            }
        }
    }
}
