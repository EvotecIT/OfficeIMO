using OfficeIMO.Word.Html;
using System;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public class HtmlDiagnostics {
        [Fact]
        public void HtmlToWord_UnsupportedInlineCss_AddsDiagnostics() {
            var callbackDiagnostics = new System.Collections.Generic.List<HtmlConversionDiagnostic>();
            var options = new HtmlToWordOptions {
                DiagnosticHandler = diagnostic => callbackDiagnostics.Add(diagnostic)
            };

            "<p style=\"color:red;display:grid;position:absolute\">Text</p>".LoadFromHtml(options);

            var diagnostics = options.Diagnostics.Where(diagnostic => diagnostic.Code == "UnsupportedCssDeclaration").ToList();
            Assert.Equal(2, diagnostics.Count);
            Assert.Contains(diagnostics, diagnostic => string.Equals(diagnostic.Source, "p:display", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(diagnostics, diagnostic => string.Equals(diagnostic.Source, "p:position", StringComparison.OrdinalIgnoreCase));
            Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.Source?.EndsWith(":color", StringComparison.OrdinalIgnoreCase) == true);
            Assert.Equal(diagnostics.Count, callbackDiagnostics.Count(diagnostic => diagnostic.Code == "UnsupportedCssDeclaration"));
        }

        [Fact]
        public void HtmlToWord_UnsupportedStylesheetCss_AddsDistinctDiagnostics() {
            var options = new HtmlToWordOptions();
            string html = "<style>.warn{display:flex;position:absolute;color:#333}</style><p class=\"warn\">One</p><p class=\"warn\">Two</p>";

            html.LoadFromHtml(options);

            var diagnostics = options.Diagnostics.Where(diagnostic => diagnostic.Code == "UnsupportedCssDeclaration").ToList();
            Assert.Equal(2, diagnostics.Count);
            Assert.Contains(diagnostics, diagnostic => string.Equals(diagnostic.Source, "p:display", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(diagnostics, diagnostic => string.Equals(diagnostic.Source, "p:position", StringComparison.OrdinalIgnoreCase));
            Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.Source?.EndsWith(":color", StringComparison.OrdinalIgnoreCase) == true);
        }

        [Fact]
        public void HtmlToWord_UnsupportedCssValues_AddDiagnostics() {
            var options = new HtmlToWordOptions();
            string html = "<p style=\"font-size:clamp(12px,2vw,20px);text-align:start;color:not-a-color\">Text</p>";

            html.LoadFromHtml(options);

            var diagnostics = options.Diagnostics.Where(diagnostic => diagnostic.Code == "UnsupportedCssValue").ToList();
            Assert.Equal(3, diagnostics.Count);
            Assert.All(diagnostics, diagnostic => Assert.Equal(HtmlConversionDiagnosticSeverity.Warning, diagnostic.Severity));
            Assert.Contains(diagnostics, diagnostic => string.Equals(diagnostic.Source, "p:font-size", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(diagnostics, diagnostic => string.Equals(diagnostic.Source, "p:text-align", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(diagnostics, diagnostic => string.Equals(diagnostic.Source, "p:color", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(diagnostics, diagnostic => diagnostic.Detail?.Contains("clamp(12px,2vw,20px)", StringComparison.OrdinalIgnoreCase) == true);
        }

        [Fact]
        public void HtmlToWord_UnsupportedCssDiagnostics_CanBeIgnored() {
            var options = new HtmlToWordOptions {
                UnsupportedCssHandling = HtmlUnsupportedCssHandling.Ignore
            };

            "<p style=\"display:grid;text-align:start\">Text</p>".LoadFromHtml(options);

            Assert.DoesNotContain(options.Diagnostics, diagnostic => diagnostic.Code.StartsWith("UnsupportedCss", StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void HtmlToWord_UnsupportedCssDiagnostics_CanStopConversion() {
            var options = new HtmlToWordOptions {
                UnsupportedCssHandling = HtmlUnsupportedCssHandling.Error
            };

            var exception = Assert.Throws<HtmlUnsupportedCssException>(() =>
                "<p style=\"text-align:start\">Text</p>".LoadFromHtml(options));

            Assert.Equal("UnsupportedCssValue", exception.Code);
            Assert.Equal("p:text-align", exception.CssSource);
            Assert.Contains("start", exception.Detail, StringComparison.OrdinalIgnoreCase);
            var diagnostic = Assert.Single(options.Diagnostics, d => d.Code == "UnsupportedCssValue");
            Assert.Equal(HtmlConversionDiagnosticSeverity.Error, diagnostic.Severity);
            Assert.Equal("p:text-align", diagnostic.Source);
        }

        [Fact]
        public void HtmlToWord_UnsupportedStylesheetCssValues_CanStopConversion() {
            var options = new HtmlToWordOptions {
                UnsupportedCssHandling = HtmlUnsupportedCssHandling.Error
            };

            var exception = Assert.Throws<HtmlUnsupportedCssException>(() =>
                "<style>p{font-size:calc(1em + 1px)}</style><p>Text</p>".LoadFromHtml(options));

            Assert.Equal("UnsupportedCssValue", exception.Code);
            Assert.Equal("p:font-size", exception.CssSource);
            Assert.Contains("calc", exception.Detail, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void HtmlToWord_UnsupportedWidthAndHeightValues_AddDiagnostics() {
            var options = new HtmlToWordOptions();

            "<p style=\"width:calc(100% - 1em);height:calc(10px + 1px)\">Text</p>".LoadFromHtml(options);

            var diagnostics = options.Diagnostics.Where(diagnostic => diagnostic.Code == "UnsupportedCssValue").ToList();
            Assert.Contains(diagnostics, diagnostic => string.Equals(diagnostic.Source, "p:width", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(diagnostics, diagnostic => string.Equals(diagnostic.Source, "p:height", StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void HtmlToWord_BorderSideDiagnostics_MatchMappedProperties() {
            var options = new HtmlToWordOptions();

            "<table style=\"border-left:1px solid #123456;border-image:linear-gradient(red, blue) 1\"><tr><td>Cell</td></tr></table>".LoadFromHtml(options);

            Assert.DoesNotContain(options.Diagnostics, diagnostic => string.Equals(diagnostic.Source, "table:border-left", StringComparison.OrdinalIgnoreCase));
            var diagnostics = options.Diagnostics.Where(diagnostic => diagnostic.Code == "UnsupportedCssDeclaration").ToList();
            Assert.Contains(diagnostics, diagnostic => string.Equals(diagnostic.Source, "table:border-image", StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void HtmlToWord_BorderLonghandDiagnostics_ReportDroppedProperties() {
            var options = new HtmlToWordOptions {
                UnsupportedCssHandling = HtmlUnsupportedCssHandling.Error
            };

            var exception = Assert.Throws<HtmlUnsupportedCssException>(() =>
                "<table><tr><td style=\"border-left-color:red\">Cell</td></tr></table>".LoadFromHtml(options));

            Assert.Equal("UnsupportedCssDeclaration", exception.Code);
            Assert.Equal("td:border-left-color", exception.CssSource);
        }
    }
}
