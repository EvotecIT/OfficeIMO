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
    }
}

