using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word.Converters {
    internal static class Markdown01_LoadFromMarkdown {
        public static void Example(string folderPath, bool openWord) {
            Console.WriteLine("[*] Loading Markdown and converting to Word");
            
            string markdown = @"# Main Title

This is a paragraph with **bold** and *italic* text.

## Features

- Bullet point 1
- Bullet point 2
- Bullet point 3

### Code Example

```csharp
var example = ""Hello World"";
```

| Column 1 | Column 2 |
|----------|----------|
| Data 1   | Data 2   |
";
            
            var doc = markdown.LoadFromMarkdown();
            string outputPath = Path.Combine(folderPath, "LoadFromMarkdown.docx");
            doc.Save(outputPath);
            
            Console.WriteLine($"✓ Created: {outputPath}");
            Console.WriteLine($"✓ Paragraphs: {doc.Paragraphs.Count}");
            Console.WriteLine($"✓ Tables: {doc.Tables.Count}");
            
            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outputPath) { UseShellExecute = true });
            }
        }
    }
}