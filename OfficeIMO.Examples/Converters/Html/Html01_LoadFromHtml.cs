using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word.Converters {
    internal static class Html01_LoadFromHtml {
        public static void Example(string folderPath, bool openWord) {
            Console.WriteLine("[*] Loading HTML and converting to Word");
            
            string html = @"<!DOCTYPE html>
<html>
<head>
    <title>Test Document</title>
</head>
<body>
    <h1>Main Title</h1>
    <p>This is a paragraph with <strong>bold</strong> and <em>italic</em> text.</p>
    
    <h2>Features</h2>
    <ul>
        <li>Bullet point 1</li>
        <li>Bullet point 2</li>
        <li>Bullet point 3</li>
    </ul>
    
    <h3>Table Example</h3>
    <table>
        <tr>
            <th>Header 1</th>
            <th>Header 2</th>
        </tr>
        <tr>
            <td>Data 1</td>
            <td>Data 2</td>
        </tr>
    </table>
    
    <p>Visit <a href=""https://github.com"">GitHub</a> for more info.</p>
</body>
</html>";
            
            var doc = html.LoadFromHtml();
            string outputPath = Path.Combine(folderPath, "LoadFromHtml.docx");
            doc.Save(outputPath);
            
            Console.WriteLine($"✓ Created: {outputPath}");
            Console.WriteLine($"✓ Paragraphs: {doc.Paragraphs.Count}");
            Console.WriteLine($"✓ Tables: {doc.Tables.Count}");
            Console.WriteLine($"✓ Hyperlinks: {doc.HyperLinks.Count}");
            
            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outputPath) { UseShellExecute = true });
            }
        }
    }
}