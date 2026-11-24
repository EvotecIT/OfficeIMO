using System.Collections.Concurrent;
using System.Net;
using System.Net.Sockets;
using HtmlTinkerX;
using Microsoft.Playwright;

if (args is ["--install-browser"]) {
    await HtmlBrowser.EnsureInstalledAsync(HtmlBrowserEngine.Chromium);
    return;
}

var chart = args is ["--chart", _, ("Office" or "Nbconvert"), ..] && args.Length <= 4;
var chartNotebook = args is ["--chart-notebook", _, ..] && args.Length <= 4;
var native = chart && args[2] == "Office";
var compactPreview = chartNotebook && args.Length == 4 ? Path.GetFullPath(args[3]) : null;
if (chart) args = args.Length == 4 ? [args[1], args[3]] : [args[1]];
if (chartNotebook) args = args[1..Math.Min(args.Length, 3)];
if (args.Length is < 1 or > 2) {
    throw new ArgumentException("Usage: WidgetVerifier <artifact.html> [preview.png], --chart-notebook <artifact.html> [wide.png] [compact.png], or --chart <artifact.html> <Office|Nbconvert> [preview.png]");
}
var artifact = Path.GetFullPath(args[0]);
var root = Path.GetDirectoryName(artifact)!;
if (!File.Exists(artifact)) throw new FileNotFoundException("HTML artifact missing.", artifact);
using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(40));
using var server = new HttpListener();
int port;
using (var reservation = new TcpListener(IPAddress.Loopback, 0)) {
    reservation.Start();
    port = ((IPEndPoint)reservation.LocalEndpoint).Port;
}
var origin = new Uri($"http://127.0.0.1:{port}/");
server.Prefixes.Add(origin.AbsoluteUri);
server.Start();
Task serve = ServeAsync(timeout.Token);
try {
    await using var browser = await HtmlBrowser.OpenSessionAsync("about:blank", new HtmlBrowserLaunchOptions {
        Headless = true,
        ViewportWidth = 1080,
        ViewportHeight = 780,
        LoadState = HtmlBrowserLoadState.DomContentLoaded,
        Timeout = 15000
    }, timeout.Token);
    var incidents = new ConcurrentQueue<string>();
    var blocked = new ConcurrentQueue<string>();
    var page = browser.Page;
    page.Response += (_, result) => { if (!result.Ok) Console.Error.WriteLine($"HTTP {result.Status} {result.Url}"); };
    page.PageError += (_, error) => { incidents.Enqueue(error); Console.Error.WriteLine(error); };
    page.Console += (_, message) => {
        if (message.Type is "error" or "warning") Console.Error.WriteLine($"Browser {message.Type}: {message.Text}");
    };
    await page.Context.RouteAsync("**/*", async route => {
        var request = route.Request;
        var uri = new Uri(request.Url);
        if (uri.Scheme is "data" or "blob" || origin.IsBaseOf(uri) && request.Method is ("GET" or "HEAD")) {
            await route.ContinueAsync();
        } else {
            blocked.Enqueue($"{request.Method} {request.Url}");
            Console.Error.WriteLine($"External request blocked: {request.Url}");
            await route.AbortAsync();
        }
    });
    var response = await page.GotoAsync(new Uri(origin, Uri.EscapeDataString(Path.GetFileName(artifact))).AbsoluteUri,
        new PageGotoOptions { WaitUntil = WaitUntilState.DOMContentLoaded, Timeout = 15000 });
    Require(response?.Ok == true, "The artifact did not load.");
    var preview = args.Length == 2 ? Path.GetFullPath(args[1]) : null;
    if (chart || chartNotebook) {
        if (chartNotebook) await ChartProof.VerifyNotebookAsync(page, preview, compactPreview);
        else await ChartProof.VerifyHeatmapAsync(page, native, preview);
        Require(blocked.IsEmpty, $"Chart attempted requests outside its offline artifact: {string.Join(", ", blocked)}");
    } else {
        await VerifyWidgetAsync(page, preview);
    }
    Require(incidents.IsEmpty, string.Join("\n", incidents));
    Console.WriteLine($"Blocked unrelated remote requests: {blocked.Count}.");
} finally {
    timeout.Cancel();
    server.Stop();
    await serve;
}

static async Task VerifyWidgetAsync(IPage page, string? preview) {
    var slider = page.Locator("input[type='range'], [role='slider'], .ui-slider-handle").First;
    await slider.WaitForAsync(new LocatorWaitForOptions { State = WaitForSelectorState.Visible, Timeout = 12000 });
    Require(await SliderValue() == 50, "Expected the saved slider value of 50.");
    if (preview is not null) await page.ScreenshotAsync(new PageScreenshotOptions { Path = preview, FullPage = true });
    await slider.PressAsync("ArrowRight", new LocatorPressOptions { Timeout = 4000 });
    Require(await SliderValue() == 51, "Could not update the range from the keyboard.");
    var play = page.Locator("button[title='Play']:visible, button[aria-label='Play']:visible, button.mod-play:visible, .widget-play button:has(.fa-play):visible, button[data-officeimo-widget-kind='play']:visible, [data-officeimo-widget-kind='play'] button:visible").First;
    await play.ClickAsync(new LocatorClickOptions { Timeout = 5000 });
    await page.WaitForFunctionAsync("""
        () => {
          const el = document.querySelector("input[type='range'], [role='slider']");
          const readout = document.querySelector(".widget-slider .widget-readout");
          return Number(el?.getAttribute("aria-valuenow") ?? el?.value ?? readout?.textContent) > 51;
        }
        """, null, new PageWaitForFunctionOptions { Timeout = 5000 });
    int afterPlay = await SliderValue();
    Require(afterPlay is > 51 and <= 100, "Play did not advance the linked slider.");
    Console.WriteLine($"Available offline. Saved value 50, keyboard value 51, linked play advanced to {afterPlay}.");

    Task<int> SliderValue() => page.EvaluateAsync<int>("""
        () => {
          const el = document.querySelector("input[type='range'], [role='slider']");
          return Number(el?.getAttribute("aria-valuenow") ?? el?.value ?? document.querySelector(".widget-slider .widget-readout")?.textContent);
        }
        """);
}

static void Require(bool condition, string message) {
    if (!condition) throw new InvalidOperationException(message);
}

async Task ServeAsync(CancellationToken token) {
    while (!token.IsCancellationRequested) {
        HttpListenerContext context;
        try {
            context = await server.GetContextAsync().WaitAsync(token);
        } catch (Exception error) when (error is OperationCanceledException or HttpListenerException or ObjectDisposedException) {
            break;
        }
        try {
            var relative = Uri.UnescapeDataString(context.Request.Url!.AbsolutePath).TrimStart('/');
            var file = Path.GetFullPath(Path.Combine(root, relative));
            if (!file.StartsWith(root + Path.DirectorySeparatorChar, StringComparison.Ordinal) || !File.Exists(file)) {
                context.Response.StatusCode = 404;
                continue;
            }
            context.Response.ContentType = Path.GetExtension(file) switch {
                ".html" => "text/html; charset=utf-8",
                ".js" => "text/javascript; charset=utf-8",
                _ => "application/octet-stream"
            };
            using var content = File.OpenRead(file);
            context.Response.ContentLength64 = content.Length;
            if (context.Request.HttpMethod != "HEAD") await content.CopyToAsync(context.Response.OutputStream, token);
        } finally {
            context.Response.Close();
        }
    }
}
