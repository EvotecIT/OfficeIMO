using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class HtmlWordToHtml {
        [Fact]
        public void Test_WordToHtml_OutputBudgetExcludesDisabledStories() {
            using var document = WordDocument.Create();
            document.AddParagraph("Visible body");
            string hiddenText = new string('x', 8192);
            var section = document.Sections[0];
            section.GetOrCreateHeader(HeaderFooterValues.Default).AddParagraph(hiddenText);
            section.GetOrCreateFooter(HeaderFooterValues.Default).AddParagraph(hiddenText);
            document.AddParagraph("Footnote anchor").AddFootNote(hiddenText);
            document.AddParagraph("Endnote anchor").AddEndNote(hiddenText);
            document.AddParagraph("Comment anchor").AddComment("Reviewer", "R", hiddenText);

            var options = new WordToHtmlOptions {
                MaxOutputCharacters = 4096,
                ExportHeadersAndFooters = false,
                ExportFootnotes = false,
                ExportEndnotes = false,
                ExportComments = false
            };

            string html = document.ToHtmlResult(options).RequireValue();

            Assert.Contains("Visible body", html, StringComparison.Ordinal);
            Assert.DoesNotContain(hiddenText, html, StringComparison.Ordinal);
        }
    }
}
