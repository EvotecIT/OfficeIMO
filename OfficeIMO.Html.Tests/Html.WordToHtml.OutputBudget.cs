using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class HtmlWordToHtml {
        [Fact]
        public void Test_WordToHtml_OutputBudgetExcludesOmittedRevisionText() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("Visible content").AddInsertedText(new string('x', 16_384), "Reviewer");
            var options = new WordToHtmlOptions { MaxOutputCharacters = 4096 };

            HtmlTextConversionResult result = document.ToHtmlResult(options);

            Assert.True(result.Succeeded);
            Assert.Contains("Visible content", result.RequireValue(), StringComparison.Ordinal);
            Assert.DoesNotContain(new string('x', 256), result.RequireValue(), StringComparison.Ordinal);
            Assert.Contains(result.Report.Diagnostics, diagnostic =>
                diagnostic.Code == "TrackedRevisionTextOmitted" &&
                diagnostic.LossKind == HtmlConversionLossKind.Omission);
        }

        [Fact]
        public void Test_WordToHtml_OutputBudgetStopsEmptyParagraphsBeforeDomConstruction() {
            using WordDocument document = WordDocument.Create();
            for (int index = 0; index < 1024; index++) {
                document.AddParagraph();
            }
            var options = new WordToHtmlOptions {
                IncludeDefaultCss = false,
                MaxOutputCharacters = 512
            };

            HtmlConversionLimitException exception = Assert.Throws<HtmlConversionLimitException>(
                () => document.ToHtmlResult(options));

            Assert.Equal("WordHtmlOutputLimitExceeded", exception.Code);
            Assert.Equal("GeneratedElement:p", exception.LimitSource);
            Assert.True(exception.Actual > exception.Limit);
        }
    }
}
