using System.Globalization;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlRender_LayoutCallbacksSeeEachCallersCultureAndFlowedContext() {
        var flowed = new AsyncLocal<string?>();
        var observed = new List<(string Culture, string? Flowed)>();
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 40D,
            Margins = HtmlRenderMargins.All(0D),
            TextHyphenationCallback = token => {
                lock (observed) observed.Add((CultureInfo.CurrentCulture.Name, flowed.Value));
                return Array.Empty<int>();
            }
        };
        const string html = "<p style='margin:0;width:40px;font-size:16px;hyphens:auto'>typography</p>";

        CultureInfo previous = CultureInfo.CurrentCulture;
        try {
            // Consecutive conversions may share a rendering thread; each must run with its own caller's state.
            foreach (string culture in new[] { "fr-FR", "de-DE", "fr-FR" }) {
                CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo(culture);
                flowed.Value = culture;
                HtmlRenderTestDriver.Render(html, options);
            }
        } finally {
            CultureInfo.CurrentCulture = previous;
        }

        Assert.Contains(observed, call => call.Culture == "de-DE");
        Assert.Contains(observed, call => call.Culture == "fr-FR");
        Assert.All(observed, call => Assert.Equal(call.Culture, call.Flowed));
    }
}
