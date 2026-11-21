using System.Text.Json;
using HtmlTinkerX;
using Microsoft.Playwright;

if (args.Length is < 1 or > 2 || !Uri.TryCreate(args[0], UriKind.Absolute, out var origin)
    || !origin.IsLoopback || origin.Scheme != Uri.UriSchemeHttp) {
    throw new ArgumentException("Usage: WebsiteBrowserContract <loopback-http-origin> [evidence-directory]");
}

using var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(5));
await using var browser = await HtmlBrowser.OpenSessionAsync("about:blank", new HtmlBrowserLaunchOptions {
    Headless = true,
    ChromiumSandbox = false,
    LoadState = HtmlBrowserLoadState.DomContentLoaded,
    Timeout = 30000,
    ViewportWidth = 1365,
    ViewportHeight = 900
}, deadline.Token);

var errors = new System.Collections.Concurrent.ConcurrentQueue<string>();
browser.Page.PageError += (_, message) => errors.Enqueue(message);
try {
    // The per-area index the build writes beside the pages says which comparisons each area declares.
    using var http = new HttpClient { BaseAddress = origin };
    using var index = JsonDocument.Parse(await http.GetStringAsync("/data/benchmarks/index.json"));
    var areas = index.RootElement.GetProperty("areas").EnumerateArray()
        .Select(area => (Id: area.GetProperty("id").GetString()!, Count: area.GetProperty("comparisons").GetArrayLength()))
        .ToArray();
    if (areas.Length == 0) throw new InvalidOperationException("The benchmark index lists no areas.");

    // The hub loads an area's results on demand: opening it must bring in every comparison of that area.
    await OpenAsync("/benchmarks/");
    var first = areas[0];
    await browser.Page.Locator($"#benchmarks-{first.Id} details[data-pf-bench-area-page] > summary").ClickAsync();
    await browser.Page.WaitForFunctionAsync("""
        ({id, count}) => {
          const details = document.querySelector(`#benchmarks-${id} details[data-pf-bench-area-page]`);
          return details?.getAttribute('data-pf-bench-loaded') === 'true' &&
            details.querySelectorAll('[data-pf-bench]').length === count;
        }
        """, new { id = first.Id, count = first.Count }, new PageWaitForFunctionOptions { Timeout = 15000 });
    await CheckLayoutAsync("hub");

    // Every area page renders every declared comparison of its area, at desktop and phone width.
    foreach (var (id, count) in areas) {
        await OpenAsync($"/benchmarks/{id}/");
        var rendered = await browser.Page.Locator("[data-pf-bench]").CountAsync();
        if (rendered != count) {
            throw new InvalidOperationException($"/benchmarks/{id}/ renders {rendered} comparisons; its area declares {count}.");
        }
        await CheckLayoutAsync(id);
    }
    if (!errors.IsEmpty) throw new InvalidOperationException("The benchmark pages raised script errors.");
    Console.WriteLine($"Benchmark hub and {areas.Length} area pages verified at desktop and phone width.");
} catch (Exception error) {
    Console.Error.WriteLine($"Benchmark browser contract failed at {browser.Page.Url}: {error.Message}");
    foreach (var message in errors) Console.Error.WriteLine($"Page error: {message}");
    try { await CaptureAsync("failure"); }
    catch (Exception captureError) { Console.Error.WriteLine($"Failure capture: {captureError.Message}"); }
    return 1;
}
return 0;

async Task OpenAsync(string path) {
    await browser.Page.SetViewportSizeAsync(1365, 900);
    await browser.Page.GotoAsync(new Uri(origin, path).AbsoluteUri, new PageGotoOptions {
        WaitUntil = WaitUntilState.DOMContentLoaded,
        Timeout = 30000
    });
}

// No page may scroll sideways; wide result tables scroll inside their own container.
async Task CheckLayoutAsync(string name) {
    foreach (var (width, height, label) in new[] { (1365, 900, "desktop"), (390, 844, "phone") }) {
        await browser.Page.SetViewportSizeAsync(width, height);
        var overflow = await browser.Page.EvaluateAsync<int>(
            "() => document.documentElement.scrollWidth - document.documentElement.clientWidth");
        if (overflow > 1) {
            throw new InvalidOperationException($"{browser.Page.Url} scrolls sideways by {overflow}px at {label} width.");
        }
        foreach (var scheme in new[] { ColorScheme.Light, ColorScheme.Dark }) {
            await browser.Page.EmulateMediaAsync(new PageEmulateMediaOptions { ColorScheme = scheme });
            await browser.Page.EvaluateAsync("theme => document.documentElement.setAttribute('data-theme', theme)",
                scheme == ColorScheme.Dark ? "dark" : "light");
            // Links fade between theme colours; capture after the transition.
            await browser.Page.WaitForTimeoutAsync(500);
            await CaptureAsync($"{name}-{label}-{scheme.ToString().ToLowerInvariant()}");
        }
    }
}

async Task CaptureAsync(string name) {
    if (args.Length != 2) return;
    Directory.CreateDirectory(args[1]);
    await browser.Page.ScreenshotAsync(new PageScreenshotOptions {
        Path = Path.Combine(args[1], name + ".png"), FullPage = true, Timeout = 20000
    });
}
