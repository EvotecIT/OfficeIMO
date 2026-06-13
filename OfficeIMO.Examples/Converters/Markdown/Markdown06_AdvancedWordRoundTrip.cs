using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

namespace OfficeIMO.Examples.Word.Converters {
    internal static class Markdown06_AdvancedWordRoundTrip {
        public static void Example(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating advanced Word -> Markdown -> Word example");

            string sourcePath = Path.Combine(folderPath, "MarkdownAdvanced.Source.docx");
            string markdownPath = Path.Combine(folderPath, "MarkdownAdvanced.VisualFallback.md");
            string roundTripPath = Path.Combine(folderPath, "MarkdownAdvanced.RoundTrip.docx");
            string resourcesPath = Path.Combine(folderPath, "MarkdownAdvanced.VisualFallback.assets");

            using var document = WordDocument.Create(sourcePath);
            BuildSourceDocument(document);
            document.Save();

            var warnings = new List<string>();
            var exportOptions = new WordToMarkdownOptions {
                IncludeHeadersAndFootersAsSemanticBlocks = true,
                PageBreakMode = MarkdownPageBreakMode.SemanticBlock,
                UnsupportedContentMode = MarkdownUnsupportedContentMode.Placeholder,
                VisualFallbackMode = MarkdownVisualFallbackMode.SvgFile,
                OnWarning = warnings.Add
            };

            document.SaveAsMarkdown(markdownPath, exportOptions);

            string markdown = File.ReadAllText(markdownPath);
            string resourcesUri = new Uri(Path.GetDirectoryName(markdownPath)!.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar).AbsoluteUri;
            var importOptions = new MarkdownToWordOptions {
                FontFamily = "Calibri",
                BaseUri = resourcesUri,
                AllowLocalImages = true
            };
            importOptions.AllowedImageDirectories.Add(resourcesPath);
            using var roundTrip = markdown.LoadFromMarkdown(importOptions);
            roundTrip.Save(roundTripPath);

            Console.WriteLine($"  Source Word:      {sourcePath}");
            Console.WriteLine($"  Markdown:         {markdownPath}");
            Console.WriteLine($"  Resources:        {resourcesPath}");
            Console.WriteLine($"  Round-trip Word:  {roundTripPath}");
            Console.WriteLine($"  Chart SVG files:  {Directory.GetFiles(resourcesPath, "*.svg").Length}");
            Console.WriteLine($"  Page break block: {markdown.Contains(WordMarkdownSemanticBlocks.PageBreakFenceLanguage, StringComparison.Ordinal)}");
            Console.WriteLine($"  Header block:     {markdown.Contains(WordMarkdownSemanticBlocks.HeaderFenceLanguage, StringComparison.Ordinal)}");
            Console.WriteLine($"  TOC marker:       {markdown.Contains("[TOC", StringComparison.Ordinal)}");
            foreach (string warning in warnings) {
                Console.WriteLine($"  Warning:          {warning}");
            }

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(sourcePath) { UseShellExecute = true });
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(markdownPath) { UseShellExecute = true });
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(roundTripPath) { UseShellExecute = true });
            }
        }

        private static void BuildSourceDocument(WordDocument document) {
            document.BuiltinDocumentProperties.Title = "Advanced Markdown conversion sample";
            document.BuiltinDocumentProperties.Subject = "Word to Markdown conversion with visual fallback";
            document.BuiltinDocumentProperties.Creator = "OfficeIMO";

            document.AddHeadersAndFooters();
            document.Header!.Default!.AddParagraph("OfficeIMO advanced Markdown conversion");
            document.Footer!.Default!.AddParagraph("Generated from OfficeIMO.Word.Markdown");

            var toc = document.AddTableOfContent(minLevel: 1, maxLevel: 3);
            toc.Text = "Contents";

            document.AddParagraph("Advanced Markdown Conversion").Style = WordParagraphStyles.Heading1;
            var intro = document.AddParagraph("This sample exports a Word report to Markdown while preserving ");
            intro.AddText("layout cues").Bold = true;
            intro.AddText(", tables, lists, page breaks, headers, footers, and a chart rendered as an SVG image fallback.");

            document.AddParagraph("What the exporter is configured to do").Style = WordParagraphStyles.Heading2;
            var list = document.AddList(WordListStyle.Bulleted);
            list.AddItem("Keep page breaks as semantic fenced blocks so Markdown can round-trip them back to Word.");
            list.AddItem("Render supported Word charts as SVG data URI images when visual fidelity is preferred.");
            list.AddItem("Emit placeholders for unsupported content instead of silently dropping it.");

            document.AddParagraph("Quarterly summary").Style = WordParagraphStyles.Heading2;
            var table = document.AddTable(4, 4);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Region";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "Q1";
            table.Rows[0].Cells[2].Paragraphs[0].Text = "Q2";
            table.Rows[0].Cells[3].Paragraphs[0].Text = "Q3";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "EMEA";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "11";
            table.Rows[1].Cells[2].Paragraphs[0].Text = "16";
            table.Rows[1].Cells[3].Paragraphs[0].Text = "21";
            table.Rows[2].Cells[0].Paragraphs[0].Text = "APAC";
            table.Rows[2].Cells[1].Paragraphs[0].Text = "9";
            table.Rows[2].Cells[2].Paragraphs[0].Text = "14";
            table.Rows[2].Cells[3].Paragraphs[0].Text = "19";
            table.Rows[3].Cells[0].Paragraphs[0].Text = "AMER";
            table.Rows[3].Cells[1].Paragraphs[0].Text = "13";
            table.Rows[3].Cells[2].Paragraphs[0].Text = "17";
            table.Rows[3].Cells[3].Paragraphs[0].Text = "24";

            document.AddParagraph("Chart visual fallback").Style = WordParagraphStyles.Heading2;
            WordChart chart = document.AddChart("Regional pipeline", roundedCorners: false, width: 640, height: 320);
            chart.AddCategories(new List<string> { "Q1", "Q2", "Q3" });
            chart.AddBar("EMEA", new List<int> { 11, 16, 21 }, OfficeColor.CornflowerBlue);
            chart.AddBar("APAC", new List<int> { 9, 14, 19 }, OfficeColor.SeaGreen);
            chart.AddBar("AMER", new List<int> { 13, 17, 24 }, OfficeColor.Orange);
            chart.ApplyPalette(WordChart.WordChartPalette.ColorBlindSafe)
                .SetWidthToPageContent(1.0, 320);

            document.AddParagraph("Pie point-color fallback").Style = WordParagraphStyles.Heading2;
            WordChart pie = document.AddChart("Rules outcome", roundedCorners: false, width: 480, height: 260);
            pie.AddPie("Passed", 42);
            pie.AddPie("Failed", 30);
            pie.AddPie("Skipped", 5);
            pie.ApplyPalette(WordChart.WordChartPalette.Professional, semanticOutcomes: true, applyToPies: true, applyToSeries: false)
                .SetWidthToPageContent(0.8, 260);

            document.AddPageBreak();
            document.AddParagraph("Second Page").Style = WordParagraphStyles.Heading1;
            var details = document.AddParagraph("The page break above is emitted as an ");
            details.AddText("officeimo-word-page-break").SetFontFamily("Consolas");
            details.AddText(" semantic block in Markdown, then restored as a real Word page break on import.");
        }
    }
}
