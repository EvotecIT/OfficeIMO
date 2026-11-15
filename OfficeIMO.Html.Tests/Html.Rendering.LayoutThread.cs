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

    [Fact]
    public void HtmlRender_LayoutCallbacksDoNotSeeAnEarlierCallersContextOrTheCallersSynchronizationContext() {
        var flowed = new AsyncLocal<string?>();
        var observed = new List<(string? Flowed, SynchronizationContext? Context)>();
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 40D,
            Margins = HtmlRenderMargins.All(0D),
            TextHyphenationCallback = token => {
                lock (observed) observed.Add((flowed.Value, SynchronizationContext.Current));
                return Array.Empty<int>();
            }
        };
        const string html = "<p style='margin:0;width:40px;font-size:16px;hyphens:auto'>typography</p>";

        flowed.Value = "first";
        HtmlRenderTestDriver.Render(html, options);
        int firstCalls = observed.Count;

        // A caller that suppresses flow gets a clean context, as on a new thread, even when the
        // rendering thread was started by an earlier conversion.
        flowed.Value = "second";
        using (ExecutionContext.SuppressFlow()) {
            HtmlRenderTestDriver.Render(html, options);
        }
        int suppressedCalls = observed.Count;

        // Layout blocks its caller, so the caller's synchronization context must not reach callbacks.
        SynchronizationContext? previous = SynchronizationContext.Current;
        try {
            SynchronizationContext.SetSynchronizationContext(new SynchronizationContext());
            HtmlRenderTestDriver.Render(html, options);
        } finally {
            SynchronizationContext.SetSynchronizationContext(previous);
        }

        Assert.True(firstCalls > 0 && suppressedCalls > firstCalls && observed.Count > suppressedCalls);
        Assert.All(observed.Take(firstCalls), call => Assert.Equal("first", call.Flowed));
        Assert.All(observed.Skip(firstCalls).Take(suppressedCalls - firstCalls), call => Assert.Null(call.Flowed));
        Assert.All(observed, call => Assert.Null(call.Context));
    }
}
