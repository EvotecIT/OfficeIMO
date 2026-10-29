using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.Html;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_Paragraph_MixedUnits() {
            string html = "<p style=\"margin-left:1.5em;padding-top:10px;text-align:right\"><span style=\"font-size:24px;color:#123456\">Test</span></p>";
            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWord(new WordHtmlImportOptions());
            var paragraph = doc.Paragraphs[0];
            Assert.Equal(OfficeIMO.Word.WordParagraphAlignment.Right, paragraph.ParagraphAlignment);
            Assert.Equal(18d, paragraph.IndentationBefore?.ToPoints());
            Assert.Equal(7.5d, paragraph.LineSpacingBefore?.ToPoints());
            var run = paragraph.GetRuns().First();
            Assert.Equal(24, run.FontSize?.ToPoints());
            Assert.Equal("123456", run.ColorHex);
        }

        [Fact]
        public void HtmlToWord_Paragraph_TextIndent_Positive() {
            string html = "<p style=\"text-indent:18pt\">Indented</p>";
            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWord(new WordHtmlImportOptions());
            var paragraph = doc.Paragraphs[0];
            Assert.Equal(18d, paragraph.IndentationFirstLine?.ToPoints());
            Assert.Null(paragraph.IndentationHanging?.ToPoints());
        }

        [Fact]
        public void HtmlToWord_Paragraph_TextIndent_Negative() {
            string html = "<p style=\"text-indent:-12pt\">Hanging</p>";
            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWord(new WordHtmlImportOptions());
            var paragraph = doc.Paragraphs[0];
            Assert.Equal(12d, paragraph.IndentationHanging?.ToPoints());
            Assert.Null(paragraph.IndentationFirstLine?.ToPoints());
        }

        [Fact]
        public void HtmlToWord_Paragraph_PaddingShorthand() {
            string html = "<p style=\"padding:12pt 0 0 6pt\">Padded</p>";
            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWord(new WordHtmlImportOptions());
            var paragraph = doc.Paragraphs[0];
            Assert.Equal(6d, paragraph.IndentationBefore?.ToPoints());
            Assert.Equal(12d, paragraph.LineSpacingBefore?.ToPoints());
        }

        [Fact]
        public void HtmlToWord_SpanStyles_MultipleDeclarations() {
            string html = "<p><span style=\"font-weight:bold;font-style:italic;text-decoration:underline line-through;font-size:16pt;color:rgb(0,128,0)\">Styled</span></p>";
            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWord(new WordHtmlImportOptions());
            var run = doc.Paragraphs[0].GetRuns().First();
            Assert.True(run.Bold);
            Assert.True(run.Italic);
            Assert.Equal(OfficeIMO.Word.WordUnderlineStyle.Single, run.Underline);
            Assert.True(run.Strike);
            Assert.Equal(16, run.FontSize?.ToPoints());
            Assert.Equal("008000", run.ColorHex);
        }

        [Fact]
        public void HtmlToWord_SpanStyles_ModernRgbColorSyntax() {
            string html = "<p><span style=\"color:rgb(100% 0% 50% / 0.5)\">Modern</span></p>";
            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWord(new WordHtmlImportOptions());
            var run = doc.Paragraphs[0].GetRuns().First();

            Assert.Equal("FF0080", run.ColorHex);
        }

        [Fact]
        public void HtmlToWord_NestedInheritance() {
            string html = "<div style=\"color:#ff0000;font-size:20px;\">A<span style=\"font-size:10px;\">B</span><span>C</span></div>";
            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWord(new WordHtmlImportOptions());
            var runs = doc.Paragraphs[0].GetRuns().ToArray();
            Assert.Equal("FF0000", runs[0].ColorHex);
            Assert.Equal(20, runs[0].FontSize?.ToPoints());
            Assert.Equal("FF0000", runs[1].ColorHex);
            Assert.Equal(10, runs[1].FontSize?.ToPoints());
            Assert.Equal("FF0000", runs[2].ColorHex);
            Assert.Equal(20, runs[2].FontSize?.ToPoints());
        }

        [Fact]
        public void HtmlToWord_BodyStylesheet_InheritsTextFormatting() {
            string html = "<style>body { color:#123456; font-size:22px; } .special { font-size:11px; }</style><p>One</p><p class=\"special\">Two</p>";
            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWord(new WordHtmlImportOptions());
            var bodyRuns = doc._wordprocessingDocument!.MainDocumentPart!.Document!.Body!.Descendants<Run>();
            var first = bodyRuns.Single(run => run.InnerText == "One");
            var second = bodyRuns.Single(run => run.InnerText == "Two");

            Assert.Equal("123456", first.RunProperties!.Color!.Val!.Value);
            Assert.Equal("44", first.RunProperties!.FontSize!.Val!.Value);
            Assert.Equal("123456", second.RunProperties!.Color!.Val!.Value);
            Assert.Equal("22", second.RunProperties!.FontSize!.Val!.Value);
        }

        [Fact]
        public void HtmlToWord_DirectStylesheetRule_OverridesInheritedStylesheetRule() {
            string html = "<style>body.theme { color:#ff0000 } p { color:#0000ff }</style><body class=\"theme\"><p>Text</p></body>";
            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWord(new WordHtmlImportOptions());
            var run = doc.Paragraphs[0].GetRuns().First();

            Assert.Equal("0000FF", run.ColorHex);
        }

        [Fact]
        public void HtmlToWord_StylesheetClass_AppliesRunFormatting() {
            string html = "<style>.special { color:#abcdef; font-size:11px; }</style><p class=\"special\">Text</p>";
            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWord(new WordHtmlImportOptions());
            var run = doc.Paragraphs[0].GetRuns().First();

            Assert.Equal("ABCDEF", run.ColorHex);
            Assert.Equal(11, run.FontSize?.ToPoints());
        }

        [Fact]
        public void HtmlToWord_BodyStyle_DoesNotInheritLayoutMargins() {
            string html = "<body style=\"margin-left:72pt;color:#123456\"><p>Text</p></body>";
            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWord(new WordHtmlImportOptions());
            var paragraph = doc.Paragraphs[0];
            var run = paragraph.GetRuns().First();

            Assert.Equal("123456", run.ColorHex);
            Assert.Null(paragraph.IndentationBefore?.ToPoints());
        }
    }
}
