using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System;
using System.IO;
using System.Text;
using Xunit;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeIMO.Tests {
    public partial class HtmlWordToHtml {
        [Fact]
        public void Test_WordToHtml_OutputBudgetMeasuresParsedMathMl() {
            using WordDocument document = WordDocument.Create();
            string equationText = new string('"', 1024) + new string('\'', 1024);
            document.AddParagraph()._paragraph.Append(
                new M.OfficeMath(new M.Run(new M.Text(equationText))));
            string expected = document.ToHtmlResult(new WordToHtmlOptions {
                IncludeDefaultCss = false
            }).RequireValue();

            HtmlTextConversionResult bounded = document.ToHtmlResult(new WordToHtmlOptions {
                IncludeDefaultCss = false,
                MaxOutputCharacters = expected.Length
            });

            Assert.True(bounded.Succeeded);
            Assert.Equal(expected, bounded.RequireValue());
        }

        [Fact]
        public void Test_WordToHtml_OutputBudgetReservesVisibleRunArtifacts() {
            using WordDocument document = WordDocument.Create();
            var run = new Run();
            for (int index = 0; index < 1024; index++) {
                run.Append(new TabChar(), new CarriageReturn(), new Break(), new NoBreakHyphen(), new SoftHyphen());
            }
            document.AddParagraph()._paragraph.Append(run);

            HtmlConversionLimitException exception = Assert.Throws<HtmlConversionLimitException>(() =>
                document.ToHtmlResult(new WordToHtmlOptions {
                    IncludeDefaultCss = false,
                    MaxOutputCharacters = 4096
                }));

            Assert.Equal("WordHtmlOutputLimitExceeded", exception.Code);
            Assert.Contains("document.xml", exception.LimitSource, StringComparison.Ordinal);
        }

        [Fact]
        public void Test_WordToHtml_OutputBudgetReplacesPrechargedTextControlContent() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph().AddStructuredDocumentTag(new string('"', 2048));
            string expected = document.ToHtmlResult(new WordToHtmlOptions {
                IncludeDefaultCss = false
            }).RequireValue();

            HtmlTextConversionResult bounded = document.ToHtmlResult(new WordToHtmlOptions {
                IncludeDefaultCss = false,
                MaxOutputCharacters = expected.Length
            });

            Assert.True(bounded.Succeeded);
            Assert.Equal(expected, bounded.RequireValue());
        }

        [Fact]
        public void Test_WordToHtml_OutputBudgetDoesNotPrechargeTransformedRunFont() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph()._paragraph.Append(new Run(
                new RunProperties(new RunFonts { Ascii = new string('F', 2048) }),
                new Text("Visible content")));
            string expected = document.ToHtmlResult(new WordToHtmlOptions {
                IncludeDefaultCss = false,
                IncludeFontStyles = true
            }).RequireValue();

            HtmlTextConversionResult bounded = document.ToHtmlResult(new WordToHtmlOptions {
                IncludeDefaultCss = false,
                IncludeFontStyles = true,
                MaxOutputCharacters = expected.Length
            });

            Assert.True(bounded.Succeeded);
            Assert.Equal(expected, bounded.RequireValue());
        }

        [Fact]
        public void Test_WordToHtml_OutputBudgetMeasuresParsedInlineSvgMarkup() {
            using WordDocument document = WordDocument.Create();
            string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"10\" height=\"10\" viewBox=\"0 0 10 10\">" +
                         string.Concat(Enumerable.Repeat("<g/>", 512)) +
                         "</svg>";
            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(svg))) {
                document.AddParagraph().AddImage(stream, "compact.svg", width: null, height: null);
            }
            string expected = document.ToHtmlResult(new WordToHtmlOptions {
                IncludeDefaultCss = false
            }).RequireValue();

            HtmlConversionLimitException exception = Assert.Throws<HtmlConversionLimitException>(() =>
                document.ToHtmlResult(new WordToHtmlOptions {
                    IncludeDefaultCss = false,
                    MaxOutputCharacters = expected.Length - 1
                }));

            Assert.Equal("WordHtmlOutputLimitExceeded", exception.Code);
            Assert.Equal("InlineSvg:serialized", exception.LimitSource);
            Assert.True(exception.Actual > exception.Limit);
        }
    }
}
