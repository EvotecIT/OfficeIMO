using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word.Html;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word {
    internal static class ConversionExamples {
        public static void Example_UnifiedConversions(string folderPath, bool openWord) {
            Console.WriteLine("[*] Demonstrating unified conversion methods");
            
            using (WordDocument document = WordDocument.Create()) {
                document.AddParagraph("# Heading 1").Style = WordParagraphStyles.Heading1;
                document.AddParagraph("This is **bold** and *italic* text with some content.");
                
                var list = document.AddList(WordListStyle.Bulleted);
                list.AddItem("First item");
                list.AddItem("Second item");
                list.AddItem("Third item");
                
                var table = document.AddTable(2, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "A1";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "B1";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "A2";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "B2";
                
                string docPath = Path.Combine(folderPath, "UnifiedExample.docx");
                document.Save(docPath);
                Console.WriteLine($"✓ Saved as Word: {docPath}");
                
                string pdfPath = Path.Combine(folderPath, "UnifiedExample.pdf");
                document.SaveAsPdf(pdfPath, new PdfSaveOptions { 
                    Orientation = PdfPageOrientation.Portrait 
                });
                Console.WriteLine($"✓ Saved as PDF: {pdfPath}");
                
                string markdownPath = Path.Combine(folderPath, "UnifiedExample.md");
                document.SaveAsMarkdown(markdownPath, new WordToMarkdownOptions());
                Console.WriteLine($"✓ Saved as Markdown: {markdownPath}");
                
                string htmlPath = Path.Combine(folderPath, "UnifiedExample.html");
                document.SaveAsHtml(htmlPath, new WordToHtmlOptions { 
                    IncludeFontStyles = true 
                });
                Console.WriteLine($"✓ Saved as HTML: {htmlPath}");
                
                string markdown = document.ToMarkdown();
                Console.WriteLine("\nMarkdown output:");
                Console.WriteLine(markdown);
                
                string html = document.ToHtml();
                Console.WriteLine("\nHTML output (first 200 chars):");
                Console.WriteLine(html.Substring(0, Math.Min(200, html.Length)) + "...");
            }
        }
        
        public static void Example_RoundTripConversions(string folderPath, bool openWord) {
            Console.WriteLine("[*] Demonstrating round-trip conversions");
            
            string markdown = @"# Main Title

This is a paragraph with **bold** and *italic* text.

## Subtitle

- Item 1
- Item 2
- Item 3

1. First numbered
2. Second numbered
3. Third numbered";
            
            Console.WriteLine("Original Markdown:");
            Console.WriteLine(markdown);
            Console.WriteLine();
            
            using (WordDocument fromMarkdown = markdown.LoadFromMarkdown()) {
                string docPath = Path.Combine(folderPath, "FromMarkdown.docx");
                fromMarkdown.Save(docPath);
                Console.WriteLine($"✓ Markdown → Word: {docPath}");
                
                string backToMarkdown = fromMarkdown.ToMarkdown();
                Console.WriteLine("\nRound-trip Markdown:");
                Console.WriteLine(backToMarkdown);
            }
            
            string html = @"<h1>Main Title</h1>
<p>This is a paragraph with <b>bold</b> and <i>italic</i> text.</p>
<h2>Subtitle</h2>
<ul>
    <li>Item 1</li>
    <li>Item 2</li>
    <li>Item 3</li>
</ul>
<ol>
    <li>First numbered</li>
    <li>Second numbered</li>
    <li>Third numbered</li>
</ol>";
            
            Console.WriteLine("\nOriginal HTML:");
            Console.WriteLine(html);
            Console.WriteLine();
            
            using (WordDocument fromHtml = html.LoadFromHtml()) {
                string docPath = Path.Combine(folderPath, "FromHtml.docx");
                fromHtml.Save(docPath);
                Console.WriteLine($"✓ HTML → Word: {docPath}");
                
                string backToHtml = fromHtml.ToHtml();
                Console.WriteLine("\nRound-trip HTML (first 200 chars):");
                Console.WriteLine(backToHtml.Substring(0, Math.Min(200, backToHtml.Length)) + "...");
            }
        }
        
        public static void Example_StreamConversions(string folderPath, bool openWord) {
            Console.WriteLine("[*] Demonstrating stream-based conversions");
            
            using (WordDocument document = WordDocument.Create()) {
                document.AddParagraph("Stream-based conversion example");
                document.AddParagraph("This demonstrates conversion using streams instead of files.");
                
                using (MemoryStream pdfStream = new MemoryStream()) {
                    document.SaveAsPdf(pdfStream);
                    byte[] pdfBytes = pdfStream.ToArray();
                    Console.WriteLine($"✓ PDF stream size: {pdfBytes.Length} bytes");
                    File.WriteAllBytes(Path.Combine(folderPath, "StreamExample.pdf"), pdfBytes);
                }
                
                using (MemoryStream markdownStream = new MemoryStream()) {
                    document.SaveAsMarkdown(markdownStream);
                    markdownStream.Position = 0;
                    using (StreamReader reader = new StreamReader(markdownStream)) {
                        string markdown = reader.ReadToEnd();
                        Console.WriteLine($"✓ Markdown stream content:");
                        Console.WriteLine(markdown);
                    }
                }
                
                using (MemoryStream htmlStream = new MemoryStream()) {
                    document.SaveAsHtml(htmlStream);
                    htmlStream.Position = 0;
                    using (StreamReader reader = new StreamReader(htmlStream)) {
                        string html = reader.ReadToEnd();
                        Console.WriteLine($"✓ HTML stream size: {html.Length} characters");
                    }
                }
            }
        }
    }
}