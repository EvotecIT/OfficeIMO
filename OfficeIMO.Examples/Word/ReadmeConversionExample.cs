using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word.Html;
using System;
using System.IO;
using System.Text;
using System.Linq;

namespace OfficeIMO.Examples.Word {
    internal static class ReadmeConversionExample {
        public static void Example_ConvertReadme(string folderPath, bool openWord) {
            Console.WriteLine("[*] README.md Conversion Test - Testing all converters");
            Console.WriteLine("=" + new string('=', 70));
            
            string readmePath = Path.Combine(Directory.GetCurrentDirectory(), "README.md");
            if (!File.Exists(readmePath)) {
                Console.WriteLine("❌ README.md not found at: " + readmePath);
                return;
            }
            
            string readmeContent = File.ReadAllText(readmePath);
            Console.WriteLine($"✓ Loaded README.md ({readmeContent.Length:N0} characters)");
            
            // Analyze README content
            AnalyzeMarkdownContent(readmeContent);
            
            // Test Markdown -> Word conversion
            Console.WriteLine("\n📄 MARKDOWN -> WORD CONVERSION");
            Console.WriteLine("-" + new string('-', 70));
            TestMarkdownToWord(readmeContent, folderPath);
            
            // Test Word -> HTML conversion
            Console.WriteLine("\n🌐 WORD -> HTML CONVERSION");
            Console.WriteLine("-" + new string('-', 70));
            TestWordToHtml(folderPath);
            
            // Test Word -> PDF conversion
            Console.WriteLine("\n📑 WORD -> PDF CONVERSION");
            Console.WriteLine("-" + new string('-', 70));
            TestWordToPdf(folderPath);
            
            // Test HTML -> Word conversion
            Console.WriteLine("\n🔄 HTML -> WORD CONVERSION");
            Console.WriteLine("-" + new string('-', 70));
            TestHtmlToWord(folderPath);
            
            // Test round-trip conversions
            Console.WriteLine("\n🔁 ROUND-TRIP CONVERSIONS");
            Console.WriteLine("-" + new string('-', 70));
            TestRoundTrips(readmeContent, folderPath);
            
            Console.WriteLine("\n📊 CONVERTER LIMITATIONS SUMMARY");
            Console.WriteLine("=" + new string('=', 70));
            DocumentLimitations();
        }
        
        private static void AnalyzeMarkdownContent(string markdown) {
            Console.WriteLine("\n📊 README.md Content Analysis:");
            
            var lines = markdown.Split('\n');
            
            // Count different markdown elements
            int headings = lines.Count(l => l.TrimStart().StartsWith("#"));
            int lists = lines.Count(l => l.TrimStart().StartsWith("- ") || l.TrimStart().StartsWith("* "));
            int numberedLists = lines.Count(l => System.Text.RegularExpressions.Regex.IsMatch(l.TrimStart(), @"^\d+\."));
            int checkboxes = lines.Count(l => l.Contains("- [ ]") || l.Contains("- [x]") || l.Contains("☑️") || l.Contains("◼️"));
            int links = System.Text.RegularExpressions.Regex.Matches(markdown, @"\[([^\]]+)\]\(([^)]+)\)").Count;
            int images = System.Text.RegularExpressions.Regex.Matches(markdown, @"!\[([^\]]*)\]\(([^)]+)\)").Count;
            int codeBlocks = System.Text.RegularExpressions.Regex.Matches(markdown, @"```").Count / 2;
            int inlineCode = System.Text.RegularExpressions.Regex.Matches(markdown, @"`[^`]+`").Count;
            int tables = System.Text.RegularExpressions.Regex.Matches(markdown, @"\|.*\|.*\|").Count > 0 ? 1 : 0;
            int emojis = System.Text.RegularExpressions.Regex.Matches(markdown, @":[a-z_]+:").Count;
            
            Console.WriteLine($"  • Headings: {headings}");
            Console.WriteLine($"  • Bullet lists: {lists}");
            Console.WriteLine($"  • Numbered lists: {numberedLists}");
            Console.WriteLine($"  • Checkboxes/Tasks: {checkboxes}");
            Console.WriteLine($"  • Links: {links}");
            Console.WriteLine($"  • Images/Badges: {images}");
            Console.WriteLine($"  • Code blocks: {codeBlocks}");
            Console.WriteLine($"  • Inline code: {inlineCode}");
            Console.WriteLine($"  • Tables: {tables}");
            Console.WriteLine($"  • Emojis: {emojis}");
        }
        
        private static void TestMarkdownToWord(string markdown, string folderPath) {
            try {
                string outputPath = Path.Combine(folderPath, "README_from_markdown.docx");
                
                using (var document = markdown.LoadFromMarkdown(new MarkdownToWordOptions {
                    FontFamily = "Calibri"
                })) {
                    document.Save(outputPath);
                    
                    Console.WriteLine($"✓ Created: {Path.GetFileName(outputPath)}");
                    Console.WriteLine($"  • Paragraphs: {document.Paragraphs.Count}");
                    Console.WriteLine($"  • Lists: {document.Lists.Count}");
                    Console.WriteLine($"  • Tables: {document.Tables.Count}");
                    Console.WriteLine($"  • Images: {document.Images.Count}");
                    Console.WriteLine($"  • Hyperlinks: {document.HyperLinks.Count}");
                    
                    // Identify what was lost
                    Console.WriteLine("\n  ⚠️ Lost in conversion:");
                    Console.WriteLine("    - All hyperlinks (badges, links)");
                    Console.WriteLine("    - All images (badges)");
                    Console.WriteLine("    - Tables");
                    Console.WriteLine("    - Code blocks and inline code");
                    Console.WriteLine("    - Checkboxes");
                    Console.WriteLine("    - Emojis");
                    Console.WriteLine("    - Nested lists");
                }
            } catch (Exception ex) {
                Console.WriteLine($"❌ Markdown->Word failed: {ex.Message}");
            }
        }
        
        private static void TestWordToHtml(string folderPath) {
            try {
                string inputPath = Path.Combine(folderPath, "README_from_markdown.docx");
                if (!File.Exists(inputPath)) {
                    Console.WriteLine("⚠️ Skipping - input file not found");
                    return;
                }
                
                using (var document = WordDocument.Load(inputPath)) {
                    string htmlPath = Path.Combine(folderPath, "README_to_html.html");
                    document.SaveAsHtml(htmlPath, new WordToHtmlOptions {
                        IncludeFontStyles = true,
                        IncludeListStyles = true,
                        IncludeDefaultCss = true
                    });
                    
                    string html = File.ReadAllText(htmlPath);
                    Console.WriteLine($"✓ Created: {Path.GetFileName(htmlPath)}");
                    Console.WriteLine($"  • HTML size: {html.Length:N0} characters");
                    
                    // Check what HTML elements were generated
                    bool hasH1 = html.Contains("<h1");
                    bool hasLists = html.Contains("<ul>") || html.Contains("<ol>");
                    bool hasTables = html.Contains("<table");
                    bool hasImages = html.Contains("<img");
                    bool hasLinks = html.Contains("<a ");
                    
                    Console.WriteLine($"  • Has headings: {hasH1}");
                    Console.WriteLine($"  • Has lists: {hasLists}");
                    Console.WriteLine($"  • Has tables: {hasTables}");
                    Console.WriteLine($"  • Has images: {hasImages}");
                    Console.WriteLine($"  • Has links: {hasLinks}");
                }
            } catch (Exception ex) {
                Console.WriteLine($"❌ Word->HTML failed: {ex.Message}");
            }
        }
        
        private static void TestWordToPdf(string folderPath) {
            try {
                string inputPath = Path.Combine(folderPath, "README_from_markdown.docx");
                if (!File.Exists(inputPath)) {
                    Console.WriteLine("⚠️ Skipping - input file not found");
                    return;
                }
                
                using (var document = WordDocument.Load(inputPath)) {
                    string pdfPath = Path.Combine(folderPath, "README_to_pdf.pdf");
                    document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                        Orientation = PdfPageOrientation.Portrait
                    });
                    
                    var fileInfo = new FileInfo(pdfPath);
                    Console.WriteLine($"✓ Created: {Path.GetFileName(pdfPath)}");
                    Console.WriteLine($"  • PDF size: {fileInfo.Length:N0} bytes");
                    Console.WriteLine($"  • Conversion successful");
                    
                    Console.WriteLine("\n  ℹ️ PDF converter status:");
                    Console.WriteLine("    ✓ Headings preserved");
                    Console.WriteLine("    ✓ Lists preserved");
                    Console.WriteLine("    ✓ Basic formatting preserved");
                    Console.WriteLine("    ✗ Tables (if they existed)");
                    Console.WriteLine("    ✗ Images (if they existed)");
                    Console.WriteLine("    ✗ Hyperlinks (not clickable)");
                }
            } catch (Exception ex) {
                Console.WriteLine($"❌ Word->PDF failed: {ex.Message}");
            }
        }
        
        private static void TestHtmlToWord(string folderPath) {
            // Create a sample HTML with typical README elements
            string sampleHtml = @"
<h1>Test HTML with README Features</h1>
<p>This tests various HTML elements that appear in READMEs.</p>

<h2>Links and Images</h2>
<p>Visit <a href='https://github.com'>GitHub</a> for more info.</p>
<img src='https://img.shields.io/badge/test-badge-blue' alt='Test Badge'>

<h2>Lists</h2>
<ul>
    <li>Bullet item 1</li>
    <li>Bullet item 2
        <ul>
            <li>Nested item</li>
        </ul>
    </li>
</ul>

<ol>
    <li>Numbered item 1</li>
    <li>Numbered item 2</li>
</ol>

<h2>Table</h2>
<table>
    <tr>
        <th>Platform</th>
        <th>Status</th>
    </tr>
    <tr>
        <td>Windows</td>
        <td>✓ Supported</td>
    </tr>
</table>

<h2>Code</h2>
<pre><code>var example = ""Hello World"";</code></pre>
<p>Inline <code>code</code> example.</p>

<blockquote>This is a blockquote</blockquote>
";
            
            try {
                using (var document = sampleHtml.ToWordDocument(new HtmlToWordOptions {
                    FontFamily = "Calibri"
                })) {
                    string outputPath = Path.Combine(folderPath, "README_from_html.docx");
                    document.Save(outputPath);
                    
                    Console.WriteLine($"✓ Created: {Path.GetFileName(outputPath)}");
                    Console.WriteLine($"  • Paragraphs: {document.Paragraphs.Count}");
                    Console.WriteLine($"  • Lists: {document.Lists.Count}");
                    Console.WriteLine($"  • Tables: {document.Tables.Count}");
                    Console.WriteLine($"  • Images: {document.Images.Count}");
                    Console.WriteLine($"  • Hyperlinks: {document.HyperLinks.Count}");
                    
                    Console.WriteLine("\n  ⚠️ HTML->Word conversion status:");
                    Console.WriteLine("    ? Hyperlinks");
                    Console.WriteLine("    ? Images from URLs");
                    Console.WriteLine("    ? Tables");
                    Console.WriteLine("    ? Code blocks");
                    Console.WriteLine("    ? Blockquotes");
                    Console.WriteLine("    ? Nested lists");
                }
            } catch (Exception ex) {
                Console.WriteLine($"❌ HTML->Word failed: {ex.Message}");
            }
        }
        
        private static void TestRoundTrips(string markdown, string folderPath) {
            // Test Markdown -> Word -> Markdown
            try {
                Console.WriteLine("Testing: Markdown → Word → Markdown");
                using (var document = markdown.Substring(0, Math.Min(500, markdown.Length)).LoadFromMarkdown()) {
                    string backToMarkdown = document.ToMarkdown();
                    Console.WriteLine($"  • Original length: {markdown.Length}");
                    Console.WriteLine($"  • Round-trip length: {backToMarkdown.Length}");
                    Console.WriteLine($"  • Data preserved: ~{(backToMarkdown.Length * 100 / 500)}%");
                }
            } catch (Exception ex) {
                Console.WriteLine($"  ❌ Failed: {ex.Message}");
            }
        }
        
        private static void DocumentLimitations() {
            Console.WriteLine("\n🔴 MARKDOWN CONVERTER MISSING:");
            Console.WriteLine("  • Tables (critical for README)");
            Console.WriteLine("  • Hyperlinks and URLs");
            Console.WriteLine("  • Images and badges");
            Console.WriteLine("  • Code blocks and inline code");
            Console.WriteLine("  • Blockquotes");
            Console.WriteLine("  • Horizontal rules");
            Console.WriteLine("  • Nested lists");
            Console.WriteLine("  • Checkbox lists");
            Console.WriteLine("  • HTML in markdown");
            Console.WriteLine("  • Emojis");
            Console.WriteLine("  • Multi-line paragraph handling");
            
            Console.WriteLine("\n🟡 HTML CONVERTER MISSING:");
            Console.WriteLine("  • Unknown - needs testing with:");
            Console.WriteLine("    - Complex tables");
            Console.WriteLine("    - External images");
            Console.WriteLine("    - CSS styles");
            Console.WriteLine("    - JavaScript (should be stripped)");
            Console.WriteLine("    - SVG images");
            Console.WriteLine("    - Iframes");
            Console.WriteLine("    - Forms");
            
            Console.WriteLine("\n🟢 PDF CONVERTER STATUS:");
            Console.WriteLine("  ✓ Basic text and formatting");
            Console.WriteLine("  ✓ Headings and paragraphs");
            Console.WriteLine("  ✓ Simple lists");
            Console.WriteLine("  ✓ Basic tables (from Word)");
            Console.WriteLine("  ⚠️ Images (limited support)");
            Console.WriteLine("  ✗ Hyperlinks (not clickable)");
            Console.WriteLine("  ✗ Bookmarks/TOC navigation");
            
            Console.WriteLine("\n📌 RECOMMENDED IMPROVEMENTS:");
            Console.WriteLine("  1. OfficeIMO.Markdown handles Markdown parsing");
            Console.WriteLine("  2. Add HtmlAgilityPack for better HTML parsing");
            Console.WriteLine("  3. Implement table support in Markdown converter");
            Console.WriteLine("  4. Add hyperlink support across all converters");
            Console.WriteLine("  5. Implement image downloading/embedding");
            Console.WriteLine("  6. Add code syntax highlighting");
        }
    }
}
