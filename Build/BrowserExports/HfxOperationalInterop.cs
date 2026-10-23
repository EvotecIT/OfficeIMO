using System.Security.Cryptography;
using System.Text.Json;
using HtmlTinkerX;
using Microsoft.Playwright;

/// <summary>Opt-in qualification of the real HtmlForgeX-generated table, capture and classic worker delivery.</summary>
internal static class HfxOperationalInterop {
    /// <summary>Opens only the supplied generated fixture and writes opt-in, independently read evidence.</summary>
    internal static async Task RunAsync(string evidence, string[] args) {
        string directory = Path.GetFullPath(args.Single(value => value.StartsWith("--hfx-operational=", StringComparison.Ordinal))["--hfx-operational=".Length..]);
        string html = Path.Combine(directory, "operational-inventory.html"), expectedPath = Path.Combine(directory, "operational-inventory.expected.json");
        using var expected = JsonDocument.Parse(File.ReadAllText(expectedPath));
        var inputs = new[] { html, expectedPath }.ToDictionary(path => path, Hash);
        var reports = new List<object>();
        foreach (HtmlBrowserEngine engine in DataTablesInterop.Engines(args)) {
            string output = Path.Combine(evidence, engine.ToString().ToLowerInvariant());
            Directory.CreateDirectory(output);
            await using var session = await HtmlBrowser.OpenSessionAsync(new Uri(html).AbsoluteUri,
                new HtmlBrowserLaunchOptions { Browser = engine, Headless = true, Timeout = 120000, Timezone = "Europe/Warsaw" });
            var errors = new List<string>();
            session.Page.PageError += (_, message) => errors.Add(message);
            session.Page.Console += (_, message) => { if (message.Type == "error") errors.Add(message.Text); };
            var external = new List<string>();
            Func<IRoute, Task> route = async request => {
                if (request.Request.Url.StartsWith("http", StringComparison.OrdinalIgnoreCase)) { external.Add(request.Request.Url); await request.AbortAsync(); }
                else await request.ContinueAsync();
            };
            await session.Context.RouteAsync("**/*", route);
            // Re-enter after observers/routing are installed so initial page failures are covered too.
            await session.Page.GotoAsync(new Uri(html).AbsoluteUri, new() { WaitUntil = WaitUntilState.Load });
            await session.Page.WaitForFunctionAsync("!!globalThis.operationalInventory && document.querySelectorAll('#operational-inventory tbody tr').length === 100");
            await session.Page.SetViewportSizeAsync(1440, 1000);
            var cases = new List<object>();
            foreach (JsonElement contract in expected.RootElement.GetProperty("cases").EnumerateArray()) {
                string name = contract.GetProperty("name").GetString()!;
                JsonElement? compactNote = null;
                if (name == "selected") {
                    await session.Page.Locator("#inventory-filter").FillAsync("Warsaw");
                    await session.Page.Locator("#inventory-sort").ClickAsync();
                    await session.Page.Locator("input[data-select-service='OPS-0001']").UncheckAsync();
                    await session.Page.Locator("#inventory-hide-site").ClickAsync();
                    await session.Page.Locator("#inventory-move-latency").ClickAsync();
                } else if (name == "empty") await session.Page.Locator("#inventory-filter").FillAsync("no matching operational service");
                JsonElement captured = await session.Page.EvaluateAsync<JsonElement>("""
                    () => {
                      const capture = operationalInventory.capture();
                      return { keys:capture.columns.map(column=>column.key), headers:capture.columns.map(column=>column.header),
                        ids:capture.ids, rows:[...capture.rows()] };
                    }
                    """);
                HfxOperationalVerifier.VerifyCapture(expected.RootElement, contract, captured);
                var files = new List<object>();
                foreach (string format in new[] { "xlsx", "pdf" }) {
                    IDownload download;
                    try {
                        download = await session.Page.RunAndWaitForDownloadAsync(() => session.Page.Locator("#inventory-export-" + format).ClickAsync());
                    } catch (Exception error) {
                        await CaptureDownloadFailureAsync(session.Page, output, name, format, error, errors, external);
                        throw;
                    }
                    string path = Path.Combine(output, "hfx-operational-" + name + "." + format);
                    await download.SaveAsAsync(path);
                    await session.Page.WaitForFunctionAsync("!document.getElementById('inventory-export-pdf').disabled");
                    JsonElement completed = await session.Page.EvaluateAsync<JsonElement>("operationalInventory.lastExport");
                    HfxOperationalVerifier.VerifyCompletion(contract, completed, new FileInfo(path).Length);
                    File.WriteAllText(path + ".json", contract.GetRawText());
                    object readback = HfxOperationalVerifier.Verify(path, expected.RootElement, contract);
                    files.Add(new { path, sha256 = Hash(path), readback, completed });
                }
                await session.Page.ScreenshotAsync(new() { Path = Path.Combine(output, "hfx-operational-" + name + "-wide.png"), FullPage = true });
                if (name == "selected") {
                    await session.Page.SetViewportSizeAsync(390, 844);
                    WorkbookVerifier.Require(await session.Page.EvaluateAsync<bool>("document.documentElement.scrollWidth <= innerWidth"), "Actual HFX compact page overflows horizontally.");
                    await session.Page.ScreenshotAsync(new() { Path = Path.Combine(output, "hfx-operational-selected-compact.png"), FullPage = true });
                    await session.Page.Locator("#operational-inventory tr:not([hidden]) [data-field='supportWindow'][role='region']").FocusAsync();
                    await session.Page.Keyboard.PressAsync("PageDown");
                    await session.Page.WaitForFunctionAsync("document.activeElement?.dataset.field === 'supportWindow' && document.activeElement.scrollTop > 0");
                    compactNote = await session.Page.EvaluateAsync<JsonElement>("""
                        () => {
                          const note = document.activeElement;
                          return { label:note.getAttribute('aria-label'), text:note.textContent, raw:note.dataset.rawValue,
                            scrollTop:note.scrollTop, scrollHeight:note.scrollHeight, clientHeight:note.clientHeight };
                        }
                        """);
                    await session.Page.ScreenshotAsync(new() { Path = Path.Combine(output, "hfx-operational-note-compact.png"), FullPage = false });
                    await session.Page.Locator("#inventory-export-pdf").FocusAsync();
                    await session.Page.EvaluateAsync("document.getElementById('operational-inventory').parentElement.scrollLeft = 0");
                    await session.Page.SetViewportSizeAsync(1440, 1000);
                }
                cases.Add(new { name, captured, files, compactNote });
            }
            WorkbookVerifier.Require(errors.Count == 0 && external.Count == 0, "Actual offline HFX fixture had errors or external requests: " + string.Join("; ", errors.Concat(external)));
            var report = new { engine = engine.ToString(), browser = session.Browser?.Version, producer = expected.RootElement.GetProperty("producer").GetString(),
                protocol = "file:", worker = "classic Blob + public OfficeIMO execution", cases, errors, externalRequests = external.Count,
                limits = "Artifact/consumer proof only; no native Excel, assistive-technology or PDF/UA acceptance" };
            File.WriteAllText(Path.Combine(output, "hfx-operational-report.json"), JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
            reports.Add(report);
            await session.Context.UnrouteAsync("**/*", route);
        }
        WorkbookVerifier.Require(inputs.All(pair => Hash(pair.Key) == pair.Value), "Generated HFX inputs changed during qualification.");
        File.WriteAllText(Path.Combine(evidence, "hfx-operational-interop.json"), JsonSerializer.Serialize(new { inputs, reports }, new JsonSerializerOptions { WriteIndented = true }));
    }

    /// <summary>Preserves the fixture's actual failure state without replacing the original download error.</summary>
    private static async Task CaptureDownloadFailureAsync(IPage page, string output, string name, string format,
        Exception original, List<string> errors, List<string> external) {
        string prefix = Path.Combine(output, "hfx-operational-" + name + "-" + format + "-failure");
        JsonElement? state = null;
        string? stateFailure = null, screenshotFailure = null;
        try {
            state = await page.EvaluateAsync<JsonElement>("""
                () => ({
                  url: location.href,
                  readyState: document.readyState,
                  lastExport: globalThis.operationalInventory?.lastExport ?? null,
                  status: document.getElementById('inventory-status')?.textContent ?? null,
                  controls: [...document.querySelectorAll('button,input[id]')].map(control => ({
                    id: control.id, tag: control.tagName, disabled: control.disabled,
                    text: control.textContent, value: control.value
                  }))
                })
                """).WaitAsync(TimeSpan.FromSeconds(5));
        } catch (Exception error) { stateFailure = error.ToString(); }
        try {
            await page.ScreenshotAsync(new() { Path = prefix + ".png", FullPage = false, Timeout = 5000 });
        } catch (Exception error) { screenshotFailure = error.ToString(); }
        try {
            File.WriteAllText(prefix + ".json", JsonSerializer.Serialize(new {
                @case = name, format, observedUtc = DateTimeOffset.UtcNow, original = original.ToString(), state,
                pageErrors = errors.ToArray(), externalRequests = external.ToArray(), stateFailure,
                screenshot = screenshotFailure is null ? prefix + ".png" : null, screenshotFailure
            }, new JsonSerializerOptions { WriteIndented = true }));
        } catch (Exception error) {
            Console.Error.WriteLine("HFX failure evidence could not be saved: " + error);
        }
    }

    private static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();
}
