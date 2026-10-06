using HtmlTinkerX;
using Microsoft.Playwright;

if (args.Length is < 1 or > 2 || !Uri.TryCreate(args[0], UriKind.Absolute, out var origin)
    || !origin.IsLoopback || origin.Scheme != Uri.UriSchemeHttp) {
    throw new ArgumentException("Usage: WebsiteBrowserContract <loopback-http-origin> [evidence-directory]");
}

using var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(2));
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
    var url = new Uri(origin, "/benchmarks/?benchmark-workload=pdf-structured-generation-net10.0"
        + "&benchmark-os=windows&benchmark-mode=full&benchmark-cpu=0xffff");
    await browser.Page.GotoAsync(url.AbsoluteUri, new PageGotoOptions {
        WaitUntil = WaitUntilState.DOMContentLoaded,
        Timeout = 30000
    });
    await WaitForEvidenceAsync("0xffff", "0xFFFF", false);
    await CaptureAsync("first-cpu-domain");
    await browser.Page.Locator("[data-library-comparison-affinity=\"0xffff0000\"]").ClickAsync();
    await WaitForEvidenceAsync("0xffff0000", "0xFFFF0000", true);
    await CaptureAsync("second-cpu-domain");
    Console.WriteLine("Rendered benchmark dependency version and both CPU-domain selections verified.");
} catch (Exception error) {
    Console.Error.WriteLine($"Benchmark browser contract failed at {browser.Page.Url}: {error.Message}");
    foreach (var message in errors) Console.Error.WriteLine($"Page error: {message}");
    try { await CaptureAsync("failure"); }
    catch (Exception captureError) { Console.Error.WriteLine($"Failure capture: {captureError.Message}"); }
    return 1;
}
return 0;

async Task WaitForEvidenceAsync(string affinity, string label, bool requireUrl) {
    await browser.Page.WaitForFunctionAsync("""
        ({affinity, label, requireUrl}) => {
          const selected = document.querySelector(`[data-library-comparison-affinity="${affinity}"]`);
          const table = document.querySelector('[data-library-comparison-table]');
          const rows = document.querySelector('[data-library-comparison-rows]')?.textContent || '';
          const meta = document.querySelector('[data-library-comparison-meta]')?.textContent || '';
          return selected?.getAttribute('aria-pressed') === 'true' && table?.hidden === false &&
            (!requireUrl || new URL(location.href).searchParams.get('benchmark-cpu') === affinity) &&
            rows.includes('QuestPDF 2026.5.0') && rows.includes('OfficeIMO') &&
            meta.includes('QuestPDF 2026.5.0') && meta.includes(`CPU affinity ${label}`);
        }
        """, new { affinity, label, requireUrl }, new PageWaitForFunctionOptions { Timeout = 15000 });
}

async Task CaptureAsync(string name) {
    if (args.Length != 2) return;
    Directory.CreateDirectory(args[1]);
    await browser.Page.ScreenshotAsync(new PageScreenshotOptions {
        Path = Path.Combine(args[1], name + ".png"), Timeout = 10000
    });
    if (await browser.Page.Locator("[data-library-comparison-benchmarks]").CountAsync() != 0) {
        await browser.Page.Locator("[data-library-comparison-benchmarks]").ScreenshotAsync(new LocatorScreenshotOptions {
            Path = Path.Combine(args[1], name + "-comparison.png"), Timeout = 10000
        });
    }
    await File.WriteAllTextAsync(Path.Combine(args[1], name + ".html"), await browser.Page.ContentAsync());
}
