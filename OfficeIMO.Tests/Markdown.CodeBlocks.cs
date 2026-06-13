using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public void WordToMarkdown_CodeParagraph_OutputFence() {
            using var doc = WordDocument.Create();
            string mono = FontResolver.Resolve("monospace")!;
            doc.AddParagraph("Console.WriteLine(\"Hello\");").SetFontFamily(mono).SetStyleId("CodeLang_csharp");

            string markdown = doc.ToMarkdown(new WordToMarkdownOptions());

            Assert.Equal("```csharp\nConsole.WriteLine(\"Hello\");\n```", markdown);
        }

        [Fact]
        public void WordToMarkdown_CodeParagraph_OutputFence_OptionFont() {
            using var doc = WordDocument.Create();
            const string codeFont = "Courier New";
            doc.AddParagraph("System.out.println(\"Hello\");").SetFontFamily(codeFont).SetStyleId("CodeLang_java");

            var options = new WordToMarkdownOptions { FontFamily = codeFont };
            string markdown = doc.ToMarkdown(options);

            Assert.Equal("```java\nSystem.out.println(\"Hello\");\n```", markdown);
        }

        [Fact]
        public void WordToMarkdown_CodeParagraph_Preserves_PageBreaks() {
            using var doc = WordDocument.Create();
            var paragraph = doc.AddParagraph();
            paragraph.SetStyleId("CodeLang_csharp");
            paragraph.PageBreakBefore = true;
            paragraph._paragraph.Append(new Run(
                new Text("Before"),
                new Break { Type = BreakValues.Page },
                new Text("After")));

            OfficeIMO.Markdown.MarkdownDoc markdown = doc.ToMarkdownDocument(new WordToMarkdownOptions());

            Assert.Collection(
                markdown.Blocks,
                block => Assert.IsType<OfficeIMO.Markdown.SemanticFencedBlock>(block),
                block => {
                    var code = Assert.IsType<OfficeIMO.Markdown.CodeBlock>(block);
                    Assert.Equal("csharp", code.Language);
                    Assert.Equal("Before", code.Content);
                },
                block => Assert.IsType<OfficeIMO.Markdown.SemanticFencedBlock>(block),
                block => {
                    var code = Assert.IsType<OfficeIMO.Markdown.CodeBlock>(block);
                    Assert.Equal("csharp", code.Language);
                    Assert.Equal("After", code.Content);
                });
        }
    }
}

