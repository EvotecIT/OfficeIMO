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

            var codeParagraph = doc.Paragraphs.First(p => p.Text.Contains("code"));
            // New Markdown engine does not assign a language-specific style to code blocks.
            // It uses monospace font on runs instead of paragraph style.
            Assert.Null(codeParagraph.StyleId);

            using MemoryStream ms = new();
            doc.Save(ms);
            ms.Position = 0;
            using WordprocessingDocument docx = WordprocessingDocument.Open(ms, false);
            var body = docx.MainDocumentPart!.Document.Body!;

            var codeRun = body.Descendants<Run>().First(r => r.InnerText.Contains("code"));
            Assert.Equal(FontResolver.Resolve("monospace"), codeRun.RunProperties!.RunFonts!.Ascii);
        }

        [Fact]
        public void Markdown_BlockQuote_Nesting_RoundTrip() {
            string md = @"> Level 1\n> > Level 2";
            var doc = md.LoadFromMarkdown();

            string markdown = doc.ToMarkdown();
            Assert.Contains("> Level 1", markdown);
            Assert.Contains("> > Level 2", markdown);
        }

        [Fact]
        public void MarkdownToWord_Renders_DetailsBlock_As_Structured_Paragraphs() {
            const string md = """
                <details open>
                <summary>More info</summary>

                Hidden text
                </details>
                """;

            using var doc = md.LoadFromMarkdown(new MarkdownToWordOptions());

            var summaryIndex = doc.Paragraphs.FindIndex(p => string.Equals(p.Text.Trim(), "More info", StringComparison.Ordinal));
            var bodyIndex = doc.Paragraphs.FindIndex(p => string.Equals(p.Text.Trim(), "Hidden text", StringComparison.Ordinal));
            Assert.True(summaryIndex >= 0);
            Assert.True(bodyIndex >= 0);

            var summaryParagraph = doc.Paragraphs[summaryIndex];
            Assert.Contains(summaryParagraph.GetRuns(), run => run.Bold);
            Assert.True(bodyIndex > summaryIndex);
        }

        [Fact]
        public void MarkdownToWord_Renders_FrontMatter_Header_Before_Body() {
            const string md = """
                ---
                title: Sample
                tags: [docs, ast]
                ---

                Body text
                """;

            using var doc = md.LoadFromMarkdown(new MarkdownToWordOptions());

            var paragraphs = doc.Paragraphs
                .Select(p => p.Text.Trim())
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .ToList();

            Assert.True(paragraphs.Count >= 5);
            Assert.Equal("---", paragraphs[0]);
            Assert.Equal("title: Sample", paragraphs[1]);
            Assert.Equal("tags: [docs, ast]", paragraphs[2]);
            Assert.Equal("---", paragraphs[3]);
            Assert.Contains("Body text", paragraphs);
        }
    }
}
