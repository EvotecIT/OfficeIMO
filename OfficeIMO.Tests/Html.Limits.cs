using OfficeIMO.Word.Html;
using System;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_MaxHtmlNodes_StopsConversionWithDiagnostic() {
            var options = new HtmlToWordOptions { MaxHtmlNodes = 1 };

            var exception = Assert.Throws<HtmlConversionLimitException>(() => "<p>One</p><p>Two</p>".LoadFromHtml(options));

            Assert.Equal("HtmlNodeLimitExceeded", exception.Code);
            Assert.Equal("MaxHtmlNodes", exception.LimitSource);
            Assert.Equal(1, exception.Limit);
            Assert.True(exception.Actual > exception.Limit);
            var diagnostic = Assert.Single(options.Diagnostics);
            Assert.Equal("HtmlNodeLimitExceeded", diagnostic.Code);
            Assert.Equal(HtmlConversionDiagnosticSeverity.Error, diagnostic.Severity);
            Assert.Contains("Actual=", diagnostic.Detail!, StringComparison.Ordinal);
            Assert.Contains("Limit=1", diagnostic.Detail!, StringComparison.Ordinal);
        }

        [Fact]
        public void HtmlToWord_MaxHtmlDepth_StopsConversionWithDiagnostic() {
            var options = new HtmlToWordOptions { MaxHtmlDepth = 2 };

            var exception = Assert.Throws<HtmlConversionLimitException>(() => "<div><div><p>Deep</p></div></div>".LoadFromHtml(options));

            Assert.Equal("HtmlDepthLimitExceeded", exception.Code);
            Assert.Equal("MaxHtmlDepth", exception.LimitSource);
            Assert.Equal(2, exception.Limit);
            Assert.True(exception.Actual > exception.Limit);
            var diagnostic = Assert.Single(options.Diagnostics);
            Assert.Equal("HtmlDepthLimitExceeded", diagnostic.Code);
            Assert.Equal(HtmlConversionDiagnosticSeverity.Error, diagnostic.Severity);
        }

        [Fact]
        public void HtmlToWord_MaxCssBytes_StopsConversionBeforeParsingStylesheet() {
            var options = new HtmlToWordOptions { MaxCssBytes = 8 };
            string html = "<style>.a { color: red; }</style><p class=\"a\">Text</p>";

            var exception = Assert.Throws<HtmlConversionLimitException>(() => html.LoadFromHtml(options));

            Assert.Equal("CssSizeLimitExceeded", exception.Code);
            Assert.Equal(8, exception.Limit);
            Assert.True(exception.Actual > exception.Limit);
            var diagnostic = Assert.Single(options.Diagnostics);
            Assert.Equal("CssSizeLimitExceeded", diagnostic.Code);
            Assert.Equal(HtmlConversionDiagnosticSeverity.Error, diagnostic.Severity);
        }

        [Fact]
        public void HtmlToWord_MaxTableCells_StopsConversionBeforeAllocatingTable() {
            var options = new HtmlToWordOptions { MaxTableCells = 3 };
            string html = "<table><tr><td>A</td><td>B</td></tr><tr><td>C</td><td>D</td></tr></table>";

            var exception = Assert.Throws<HtmlConversionLimitException>(() => html.LoadFromHtml(options));

            Assert.Equal("TableSizeLimitExceeded", exception.Code);
            Assert.Equal("MaxTableCells", exception.LimitSource);
            Assert.Equal(3, exception.Limit);
            Assert.Equal(4, exception.Actual);
            var diagnostic = Assert.Single(options.Diagnostics);
            Assert.Equal("TableSizeLimitExceeded", diagnostic.Code);
            Assert.Equal(HtmlConversionDiagnosticSeverity.Error, diagnostic.Severity);
        }

        [Fact]
        public void HtmlToWord_MaxTableCells_StopsConversionWhileExpandingLargeSpans() {
            var options = new HtmlToWordOptions { MaxTableCells = 10 };
            string html = "<table><tr><td colspan=\"1000000\">Wide</td></tr></table>";

            var exception = Assert.Throws<HtmlConversionLimitException>(() => html.LoadFromHtml(options));

            Assert.Equal("TableSizeLimitExceeded", exception.Code);
            Assert.Equal("MaxTableCells", exception.LimitSource);
            Assert.Equal(10, exception.Limit);
            Assert.True(exception.Actual > exception.Limit);
            var diagnostic = Assert.Single(options.Diagnostics);
            Assert.Equal("TableSizeLimitExceeded", diagnostic.Code);
        }

        [Fact]
        public void HtmlToWord_DefaultMaxTableCells_StopsHostileColumnSpan() {
            var options = new HtmlToWordOptions();
            string html = "<table><tr><td colspan=\"50001\">Wide</td></tr></table>";

            var exception = Assert.Throws<HtmlConversionLimitException>(() => html.LoadFromHtml(options));

            Assert.Equal("TableSizeLimitExceeded", exception.Code);
            Assert.Equal("MaxTableCells", exception.LimitSource);
            Assert.Equal(50000, exception.Limit);
            Assert.True(exception.Actual > exception.Limit);
            var diagnostic = Assert.Single(options.Diagnostics);
            Assert.Equal("TableSizeLimitExceeded", diagnostic.Code);
        }

        [Fact]
        public void HtmlToWord_ColumnGroupSpan_DoesNotExpandBeyondResolvedColumns() {
            var options = new HtmlToWordOptions();
            string html = "<table><colgroup><col span=\"1000000\" style=\"width:10px\"></colgroup><tr><td>Cell</td></tr></table>";

            using var document = html.LoadFromHtml(options);

            Assert.Single(document.Tables);
            Assert.Empty(options.Diagnostics);
        }
    }
}
