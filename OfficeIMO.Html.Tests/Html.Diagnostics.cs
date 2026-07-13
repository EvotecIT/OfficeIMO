using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using OfficeIMO.Word.Html;
using System;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public class HtmlDiagnostics {
        [Fact]
        public void HtmlToWord_UnsupportedInlineCss_AddsDiagnostics() {
            var options = new HtmlToWordOptions();

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse("<p style=\"color:red;display:grid;position:absolute\">Text</p>").ToWordDocumentResult(options);
            using var document = conversion.Value;

            var diagnostics = conversion.Report.Diagnostics.Where(diagnostic => diagnostic.Code == "UnsupportedCssDeclaration").ToList();
            Assert.Equal(2, diagnostics.Count);
            Assert.Contains(diagnostics, diagnostic => string.Equals(diagnostic.Source, "p:display", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(diagnostics, diagnostic => string.Equals(diagnostic.Source, "p:position", StringComparison.OrdinalIgnoreCase));
            Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.Source?.EndsWith(":color", StringComparison.OrdinalIgnoreCase) == true);
        }

        [Fact]
        public void HtmlToWord_UnsupportedStylesheetCss_AddsDistinctDiagnostics() {
            var options = new HtmlToWordOptions();
            string html = "<style>.warn{display:flex;position:absolute;color:#333}</style><p class=\"warn\">One</p><p class=\"warn\">Two</p>";

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            using var document = conversion.Value;

            var diagnostics = conversion.Report.Diagnostics.Where(diagnostic => diagnostic.Code == "UnsupportedCssDeclaration").ToList();
            Assert.Equal(2, diagnostics.Count);
            Assert.Contains(diagnostics, diagnostic => string.Equals(diagnostic.Source, "p:display", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(diagnostics, diagnostic => string.Equals(diagnostic.Source, "p:position", StringComparison.OrdinalIgnoreCase));
            Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.Source?.EndsWith(":color", StringComparison.OrdinalIgnoreCase) == true);
        }

        [Fact]
        public void HtmlToWord_UnsupportedCssValues_AddDiagnostics() {
            var options = new HtmlToWordOptions();
            string html = "<p style=\"font-size:clamp(12px,2vw,20px);text-align:match-parent;color:not-a-color\">Text</p>";

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            using var document = conversion.Value;

            var diagnostics = conversion.Report.Diagnostics.Where(diagnostic => diagnostic.Code == "UnsupportedCssValue").ToList();
            Assert.Equal(3, diagnostics.Count);
            Assert.All(diagnostics, diagnostic => Assert.Equal(HtmlDiagnosticSeverity.Warning, diagnostic.Severity));
            Assert.Contains(diagnostics, diagnostic => string.Equals(diagnostic.Source, "p:font-size", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(diagnostics, diagnostic => string.Equals(diagnostic.Source, "p:text-align", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(diagnostics, diagnostic => string.Equals(diagnostic.Source, "p:color", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(diagnostics, diagnostic => diagnostic.Detail?.Contains("clamp(12px,2vw,20px)", StringComparison.OrdinalIgnoreCase) == true);
        }

        [Theory]
        [InlineData("transparent")]
        [InlineData("currentColor")]
        public void HtmlToWord_UnmappedColorKeywords_CanStopConversion(string colorValue) {
            var options = new HtmlToWordOptions {
                UnsupportedCssHandling = HtmlUnsupportedCssHandling.Error
            };

            var exception = Assert.Throws<HtmlUnsupportedCssException>(() =>
                OfficeIMO.Html.HtmlConversionDocument.Parse($"<p style=\"color:{colorValue}\">Text</p>").ToWordDocument(options));

            Assert.Equal("UnsupportedCssValue", exception.Code);
            Assert.Equal("p:color", exception.CssSource);
            Assert.Contains(colorValue, exception.Detail, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void HtmlToWord_UnsupportedCssDiagnostics_CanBeIgnored() {
            var options = new HtmlToWordOptions {
                UnsupportedCssHandling = HtmlUnsupportedCssHandling.Ignore
            };

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse("<p style=\"display:grid;text-align:match-parent\">Text</p>").ToWordDocumentResult(options);
            using var document = conversion.Value;

            Assert.DoesNotContain(conversion.Report.Diagnostics, diagnostic => diagnostic.Code.StartsWith("UnsupportedCss", StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void HtmlToWord_UnsupportedCssDiagnostics_CanStopConversion() {
            var options = new HtmlToWordOptions {
                UnsupportedCssHandling = HtmlUnsupportedCssHandling.Error
            };

            var exception = Assert.Throws<HtmlUnsupportedCssException>(() =>
                OfficeIMO.Html.HtmlConversionDocument.Parse("<p style=\"text-align:match-parent\">Text</p>").ToWordDocument(options));

            Assert.Equal("UnsupportedCssValue", exception.Code);
            Assert.Equal("p:text-align", exception.CssSource);
            Assert.Contains("match-parent", exception.Detail, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void HtmlToWord_UnsupportedStylesheetCssValues_CanStopConversion() {
            var options = new HtmlToWordOptions {
                UnsupportedCssHandling = HtmlUnsupportedCssHandling.Error
            };

            var exception = Assert.Throws<HtmlUnsupportedCssException>(() =>
                OfficeIMO.Html.HtmlConversionDocument.Parse("<style>p{font-size:calc(1em + 1px)}</style><p>Text</p>").ToWordDocument(options));

            Assert.Equal("UnsupportedCssValue", exception.Code);
            Assert.Equal("p:font-size", exception.CssSource);
            Assert.Contains("calc", exception.Detail, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void HtmlToWord_UnsupportedWidthAndHeightValues_AddDiagnostics() {
            var options = new HtmlToWordOptions();

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse("<p style=\"width:calc(100% - 1em);height:calc(10px + 1px)\">Text</p>").ToWordDocumentResult(options);
            using var document = conversion.Value;

            var diagnostics = conversion.Report.Diagnostics.Where(diagnostic => diagnostic.Code == "UnsupportedCssValue").ToList();
            Assert.Contains(diagnostics, diagnostic => string.Equals(diagnostic.Source, "p:width", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(diagnostics, diagnostic => string.Equals(diagnostic.Source, "p:height", StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void HtmlToWord_UnsupportedMappedCssValues_AddDiagnostics() {
            var options = new HtmlToWordOptions();

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse("<p style=\"vertical-align:middle\">Text</p><table style=\"border-collapse:discard;border-spacing:calc(1px + 1px)\"><tr><td style=\"direction:sideways;vertical-align:middle\">Cell</td></tr></table>").ToWordDocumentResult(options);
            using var document = conversion.Value;

            var diagnostics = conversion.Report.Diagnostics.Where(diagnostic => diagnostic.Code == "UnsupportedCssValue").ToList();
            Assert.Contains(diagnostics, diagnostic =>
                string.Equals(diagnostic.Source, "p:vertical-align", StringComparison.OrdinalIgnoreCase) &&
                diagnostic.Detail?.Contains("middle", StringComparison.OrdinalIgnoreCase) == true);
            Assert.DoesNotContain(diagnostics, diagnostic =>
                string.Equals(diagnostic.Source, "td:vertical-align", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(diagnostics, diagnostic =>
                string.Equals(diagnostic.Source, "table:border-collapse", StringComparison.OrdinalIgnoreCase) &&
                diagnostic.Detail?.Contains("discard", StringComparison.OrdinalIgnoreCase) == true);
            Assert.Contains(diagnostics, diagnostic =>
                string.Equals(diagnostic.Source, "table:border-spacing", StringComparison.OrdinalIgnoreCase) &&
                diagnostic.Detail?.Contains("calc", StringComparison.OrdinalIgnoreCase) == true);
            Assert.Contains(diagnostics, diagnostic =>
                string.Equals(diagnostic.Source, "td:direction", StringComparison.OrdinalIgnoreCase) &&
                diagnostic.Detail?.Contains("sideways", StringComparison.OrdinalIgnoreCase) == true);
        }

        [Fact]
        public void HtmlToWord_TextDecorationShorthandColor_AddsValueDiagnosticButMapsStyle() {
            var options = new HtmlToWordOptions();

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse("<p style=\"text-decoration:underline wavy red\">Text</p>").ToWordDocumentResult(options);
            using var doc = conversion.Value;

            var run = doc.Paragraphs[0].GetRuns().Single();
            Assert.Equal(UnderlineValues.Wave, run.Underline);
            var diagnostic = Assert.Single(conversion.Report.Diagnostics, diagnostic =>
                string.Equals(diagnostic.Code, "UnsupportedCssValue", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(diagnostic.Source, "p:text-decoration", StringComparison.OrdinalIgnoreCase));
            Assert.Contains("color token 'red'", diagnostic.Detail, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void HtmlToWord_TextDecorationLonghandDiagnostics_MatchMappedProperties() {
            var options = new HtmlToWordOptions();

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse("<p style=\"text-decoration-line:underline;text-decoration-style:wavy;text-decoration-color:red\">Text</p>").ToWordDocumentResult(options);
            using var document = conversion.Value;

            Assert.DoesNotContain(conversion.Report.Diagnostics, diagnostic => string.Equals(diagnostic.Source, "p:text-decoration-line", StringComparison.OrdinalIgnoreCase));
            Assert.DoesNotContain(conversion.Report.Diagnostics, diagnostic => string.Equals(diagnostic.Source, "p:text-decoration-style", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(conversion.Report.Diagnostics, diagnostic =>
                string.Equals(diagnostic.Code, "UnsupportedCssDeclaration", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(diagnostic.Source, "p:text-decoration-color", StringComparison.OrdinalIgnoreCase));
        }

        [Theory]
        [InlineData("border-collapse:discard", "table:border-collapse", "discard")]
        [InlineData("border-spacing:calc(1px + 1px)", "table:border-spacing", "calc")]
        [InlineData("direction:sideways", "table:direction", "sideways")]
        public void HtmlToWord_UnsupportedMappedCssValues_CanStopConversion(string style, string expectedSource, string expectedValue) {
            var options = new HtmlToWordOptions {
                UnsupportedCssHandling = HtmlUnsupportedCssHandling.Error
            };

            var exception = Assert.Throws<HtmlUnsupportedCssException>(() =>
                OfficeIMO.Html.HtmlConversionDocument.Parse($"<table style=\"{style}\"><tr><td>Cell</td></tr></table>").ToWordDocument(options));

            Assert.Equal("UnsupportedCssValue", exception.Code);
            Assert.Equal(expectedSource, exception.CssSource);
            Assert.Contains(expectedValue, exception.Detail, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void HtmlToWord_BorderSideDiagnostics_MatchMappedProperties() {
            var options = new HtmlToWordOptions();

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse("<table style=\"border-left:1px solid #123456;border-image:linear-gradient(red, blue) 1\"><tr><td>Cell</td></tr></table>").ToWordDocumentResult(options);
            using var document = conversion.Value;

            Assert.DoesNotContain(conversion.Report.Diagnostics, diagnostic => string.Equals(diagnostic.Source, "table:border-left", StringComparison.OrdinalIgnoreCase));
            var diagnostics = conversion.Report.Diagnostics.Where(diagnostic => diagnostic.Code == "UnsupportedCssDeclaration").ToList();
            Assert.Contains(diagnostics, diagnostic => string.Equals(diagnostic.Source, "table:border-image", StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void HtmlToWord_BorderLonghandDiagnostics_ReportDroppedProperties() {
            var options = new HtmlToWordOptions {
                UnsupportedCssHandling = HtmlUnsupportedCssHandling.Error
            };

            var exception = Assert.Throws<HtmlUnsupportedCssException>(() =>
                OfficeIMO.Html.HtmlConversionDocument.Parse("<table><tr><td style=\"border-left-color:red\">Cell</td></tr></table>").ToWordDocument(options));

            Assert.Equal("UnsupportedCssDeclaration", exception.Code);
            Assert.Equal("td:border-left-color", exception.CssSource);
        }

        [Theory]
        [InlineData("caption")]
        [InlineData("italic 12pt/14pt Arial")]
        public void HtmlToWord_UnsupportedFontShorthandValues_CanStopConversion(string fontValue) {
            var options = new HtmlToWordOptions {
                UnsupportedCssHandling = HtmlUnsupportedCssHandling.Error
            };

            var exception = Assert.Throws<HtmlUnsupportedCssException>(() =>
                OfficeIMO.Html.HtmlConversionDocument.Parse($"<p style=\"font:{fontValue}\">Text</p>").ToWordDocument(options));

            Assert.Equal("UnsupportedCssValue", exception.Code);
            Assert.Equal("p:font", exception.CssSource);
            Assert.Contains(fontValue, exception.Detail, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void HtmlToWord_DocumentStylesheetLinks_AreSkippedByDefault() {
            var options = new HtmlToWordOptions();

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse("<html><head><link rel=\"stylesheet\" href=\"https://example.invalid/site.css\"></head><body><p>Text</p></body></html>").ToWordDocumentResult(options);
            using var document = conversion.Value;

            var diagnostic = Assert.Single(conversion.Report.Diagnostics, diagnostic => diagnostic.Code == "HtmlStylesheetLinkSkipped");
            Assert.Equal(HtmlDiagnosticSeverity.Warning, diagnostic.Severity);
        }
    }
}
