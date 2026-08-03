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

        [Fact]
        public void Test_WordToHtml_OutputBudgetExcludesUnreferencedNotesAndComments() {
            using var document = WordDocument.Create();
            document.AddParagraph("Visible body");
            string orphanedText = new string('x', 8192);
            WordParagraph footnoteReference = document.AddParagraph().AddFootNote(orphanedText);
            footnoteReference._run!.Remove();
            WordParagraph commentParagraph = document.AddParagraph("Unreviewed text");
            commentParagraph.AddComment("Reviewer", "R", orphanedText);
            foreach (CommentReference reference in commentParagraph._paragraph.Descendants<CommentReference>().ToList()) {
                reference.Remove();
            }

            string html = document.ToHtmlResult(new WordToHtmlOptions {
                MaxOutputCharacters = 4096,
                ExportFootnotes = true,
                ExportComments = true,
                IncludeDefaultCss = false
            }).RequireValue();

            Assert.Contains("Visible body", html, StringComparison.Ordinal);
            Assert.DoesNotContain(orphanedText, html, StringComparison.Ordinal);
        }

        [Fact]
        public void Test_WordToHtml_FieldDiagnosticOnlyDescribesExportedStories() {
            using var document = WordDocument.Create();
            document.AddParagraph("Visible body");
            WordParagraph header = document.Sections[0]
                .GetOrCreateHeader(HeaderFooterValues.Default)
                .AddParagraph();
            header._paragraph.Append(new SimpleField(new Run(new Text("Header field result"))) {
                Instruction = "DATE"
            });

            var omitted = document.ToHtmlResult(new WordToHtmlOptions {
                ExportHeadersAndFooters = false
            });
            var exported = document.ToHtmlResult(new WordToHtmlOptions {
                ExportHeadersAndFooters = true
            });

            Assert.Contains(omitted.Report.Diagnostics, diagnostic =>
                diagnostic.Code == "HeadersFootersOmitted");
            Assert.DoesNotContain(omitted.Report.Diagnostics, diagnostic =>
                diagnostic.Code == "FieldInstructionsFlattened");
            Assert.Contains(exported.Report.Diagnostics, diagnostic =>
                diagnostic.Code == "FieldInstructionsFlattened");
        }
    }
}
