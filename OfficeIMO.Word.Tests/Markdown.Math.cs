using OfficeIMO.Markdown;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public class MarkdownMathToWordTests {
        [Fact]
        public void TexMath_BecomesNativeOfficeMath_AndUnreadableTexStaysText() {
            const string markdown = "Energy $E = mc^2$ holds.\n\n$$\n\\sqrt{x}\n$$\n\n```math\n\\begin{unknownenv}x\\end{unknownenv}\n```\n\n```math\n\\begin{gathered}a\\\\b\\end{gathered}\n```";
            var warnings = new List<string>();

            using WordDocument document = MarkdownDocument.Parse(markdown, new MarkdownLoadOptions { Math = true })
                .ToWord(new MarkdownWordExportOptions { OnWarning = warnings.Add });

            Assert.Equal(2, document.Equations.Count);
            Assert.Equal("E=mc^(2)", document.Equations[0].ToExpression().ToPlainText());
            Assert.Contains("<m:oMathPara", document.Equations[1].Omml, StringComparison.Ordinal);
            Assert.Contains(document.Paragraphs, paragraph => paragraph.Text.Contains("unknownenv", StringComparison.Ordinal));
            Assert.Single(warnings, warning => warning.Contains("unknownenv", StringComparison.Ordinal));
            Assert.Single(warnings, warning => warning.Contains("gathered", StringComparison.Ordinal));
            Assert.Empty(document.ValidateDocument());
        }
    }
}
