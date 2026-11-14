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
    // Each CPU-domain workload is its own section; both must draw their tables from the published results.
    var url = new Uri(origin, "/benchmarks/pdf/");
    await browser.Page.GotoAsync(url.AbsoluteUri, new PageGotoOptions {
        WaitUntil = WaitUntilState.DOMContentLoaded,
        Timeout = 30000
    });
    await WaitForEvidenceAsync("pdf-structured-generation-net10.0-affinity-0xffff", "0xFFFF");
    await CaptureAsync("first-cpu-domain");
    await WaitForEvidenceAsync("pdf-structured-generation-net10.0-affinity-0xffff0000", "0xFFFF0000");
    await CaptureAsync("second-cpu-domain");
    Console.WriteLine("Rendered benchmark dependency version and both CPU-domain workloads verified.");
} catch (Exception error) {
    Console.Error.WriteLine($"Benchmark browser contract failed at {browser.Page.Url}: {error.Message}");
    foreach (var message in errors) Console.Error.WriteLine($"Page error: {message}");
    try { await CaptureAsync("failure"); }
    catch (Exception captureError) { Console.Error.WriteLine($"Failure capture: {captureError.Message}"); }
    return 1;
}
return 0;

async Task WaitForEvidenceAsync(string workload, string label) {
    await browser.Page.Locator($"[id=\"{workload}\"]").ScrollIntoViewIfNeededAsync();
    await browser.Page.WaitForFunctionAsync("""
        ({workload, label}) => {
          const section = document.getElementById(workload);
          const chart = section?.querySelector('[data-benchmark-chart]');
          const text = chart?.textContent || '';
          return section?.querySelector('.imo-bench-run[aria-pressed="true"]') &&
            chart.querySelector('table') && text.includes('QuestPDF') && text.includes('2026.5.0') &&
            text.includes('OfficeIMO') && text.includes(`CPU affinity ${label}`);
        }
        """, new { workload, label }, new PageWaitForFunctionOptions { Timeout = 15000 });
}

async Task CaptureAsync(string name) {
    if (args.Length != 2) return;
    Directory.CreateDirectory(args[1]);
    await browser.Page.ScreenshotAsync(new PageScreenshotOptions {
        Path = Path.Combine(args[1], name + ".png"), Timeout = 10000
    });
    await File.WriteAllTextAsync(Path.Combine(args[1], name + ".html"), await browser.Page.ContentAsync());
}
