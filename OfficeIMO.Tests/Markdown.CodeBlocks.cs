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
    }
}

