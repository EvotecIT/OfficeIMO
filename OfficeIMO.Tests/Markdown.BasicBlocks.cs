using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public void MarkdownToWord_ConvertsVariousElements() {
            string imagePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png"));
            string md = $@"# Heading 1

Paragraph with **bold** and *italic* and [link](https://example.com).

- Item 1
- Item 2

```c
code
```

|A|B|
|-|-|
|1|2|

> Quote line

---

![Alt]({imagePath})
";
            var doc = md.LoadFromMarkdown(new MarkdownToWordOptions { FontFamily = "Calibri" });

            Assert.Equal(WordParagraphStyles.Heading1, doc.Paragraphs[0].Style);
              var quoteParagraph = doc.Paragraphs.First(p => p.Text.Contains("Quote line"));
              Assert.True(quoteParagraph.IndentationBefore > 0);

            using MemoryStream ms = new();
            doc.Save(ms);
            ms.Position = 0;
              using WordprocessingDocument docx = WordprocessingDocument.Open(ms, false);
              var body = docx.MainDocumentPart!.Document.Body!;

              var codeRun = body.Descendants<Run>().First(r => r.InnerText.Contains("code"));
              Assert.Equal("Consolas", codeRun.RunProperties!.RunFonts!.Ascii);
        }
    }
}
