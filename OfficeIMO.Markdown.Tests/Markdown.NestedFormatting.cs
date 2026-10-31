using System.Linq;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public void Markdown_NestedEmphasis_BoldItalic() {
            string md = "This ***bolditalic*** text.";
            var doc = OfficeIMO.Markdown.MarkdownDocument.Parse(md).ToWord(new WordMarkdownImportOptions());
            var run = doc.Paragraphs[0].GetRuns().First(r => r.Bold && r.Italic);
            Assert.Equal("bolditalic", run.Text);
        }

        [Fact]
        public void Markdown_MixedFormatting_StrikeBoldItalic() {
            string md = "Text ~~**bold strike**~~ and ***bold italic***.";
            var doc = OfficeIMO.Markdown.MarkdownDocument.Parse(md).ToWord(new WordMarkdownImportOptions());
            // Validate via round-trip because the engine preserves inner markers inside strike
            string roundTrip = doc.ToMarkdown(new WordMarkdownExportOptions { EnableUnderline = true }).ToMarkdown();
            Assert.True(
                roundTrip.Contains("~~**bold strike**~~", StringComparison.Ordinal)
                || roundTrip.Contains("**~~bold strike~~**", StringComparison.Ordinal),
                $"Expected bold+strike content in roundtrip output, got:{Environment.NewLine}{roundTrip}");
            Assert.Contains("***bold italic***", roundTrip);
        }

        [Fact]
        public void Markdown_NestedFormatting_StrikeUnderline_RoundTrip() {
            string md = "Text ~~<u>strike underline</u>~~ and <u>~~underline strike~~</u>.";
            var doc = OfficeIMO.Markdown.MarkdownDocument.Parse(md).ToWord(new WordMarkdownImportOptions());
            // Run-level nesting varies; validate round-trip Markdown output instead
            string roundTrip = doc.ToMarkdown(new WordMarkdownExportOptions { EnableUnderline = true }).ToMarkdown();
            Assert.Contains("~~<u>strike underline</u>~~", roundTrip);
            Assert.True(
                roundTrip.Contains("<u>~~underline strike~~</u>", StringComparison.Ordinal)
                || roundTrip.Contains("~~<u>underline strike</u>~~", StringComparison.Ordinal),
                $"Expected underline+strike content in roundtrip output, got:{Environment.NewLine}{roundTrip}");
        }
    }
}
