using System.Security.Cryptography;
using System.Text.Json;
using HtmlTinkerX;
using OfficeIMO.Browser;

/// <summary>Opt-in qualification against an explicit native CanopyX candidate, using the shared browser owner.</summary>
internal static class CanopyInterop {
    internal static async Task RunAsync(string repository, string evidence, string[] args) {
        string canopy = Path.GetFullPath(args.FirstOrDefault(a => a.StartsWith("--canopy=", StringComparison.Ordinal))?[9..]
            ?? throw new ArgumentException("Pass --canopy=<CanopyX-repository>."));
        string assets = Path.Combine(canopy, "CanopyX.Html", "Assets");
        string script = File.ReadAllText(Path.Combine(repository, "Build", "BrowserExports", "canopy-contracts.js"));
        var inputs = new Dictionary<string, string>();
        foreach (string name in new[] { "canopyx-reporting.js", "canopyx-reporting.css", "canopyx-grid.js", "canopyx-grid.css", "canopyx-filter-editor.js", "canopyx-filter-editor.css" })
            inputs.Add(name, File.ReadAllText(Path.Combine(assets, name)));
        var hashes = inputs.ToDictionary(pair => pair.Key, pair => Convert.ToHexString(SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(pair.Value))).ToLowerInvariant());
        string regular = Convert.ToBase64String(File.ReadAllBytes(Path.Combine(repository, "Website", "Apps", "OfficeIMO.Web.Converter", "Assets", "Fonts", "Carlito-Regular.ttf")));
        string bold = Convert.ToBase64String(File.ReadAllBytes(Path.Combine(repository, "Website", "Apps", "OfficeIMO.Web.Converter", "Assets", "Fonts", "Carlito-Bold.ttf")));
        var reports = new List<object>();
        foreach (var engine in DataTablesInterop.Engines(args)) {
            string output = Path.Combine(evidence, engine.ToString().ToLowerInvariant()); Directory.CreateDirectory(output);
            string host = Path.Combine(output, "host.html"); File.WriteAllText(host, "<!doctype html><meta charset=utf-8><meta name=viewport content=\"width=device-width,initial-scale=1\"><title>CanopyX exports</title><body class=cx-page><h1>CanopyX export qualification</h1><main id=fixture></main>");
            await using var session = await HtmlBrowser.OpenSessionAsync(new Uri(host).AbsoluteUri,
                new HtmlBrowserLaunchOptions { Browser = engine, Headless = true, Timeout = 120000 });
            var errors = new List<string>(); session.Page.PageError += (_, error) => errors.Add(error);
            var files = new List<object>();
            async Task CaptureAsync(string name, int width, int height) {
                await session.Page.SetViewportSizeAsync(width, height);
                await session.Page.EvaluateAsync("() => document.fonts.ready.then(() => true)");
                await session.Page.WaitForFunctionAsync("""
                    () => {
                        const host = document.querySelector('#fixture .cx-data-grid');
                        if (!host || host.classList.contains('cx-phone') !== (host.clientWidth <= 640)) return false;
                        const rows = [...host.querySelectorAll('.cx-row[data-cx-id]')];
                        return rows.length > 0 && rows.every(row =>
                            row.querySelectorAll('.cx-hidden-fields').length === (host.classList.contains('cx-phone') ? 1 : 0));
                    }
                    """);
                // ResizeObserver updates widths/phone chrome, then schedules the row paint on later frames.
                await session.Page.EvaluateAsync("() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(() => requestAnimationFrame(() => requestAnimationFrame(resolve)))))");
                JsonElement layout = await session.Page.EvaluateAsync<JsonElement>("""
                    () => {
                        const host = document.querySelector('#fixture .cx-data-grid'), grid = host.querySelector('.cx-grid');
                        return { viewport: { width: innerWidth, height: innerHeight }, hostWidth: host.clientWidth,
                            phone: host.classList.contains('cx-phone'), gridHeight: grid.clientHeight,
                            font: getComputedStyle(host).fontFamily, headingFont: getComputedStyle(document.querySelector('h1')).fontFamily,
                            columns: host.canopyx?.getColumnLayout?.(),
                            headings: [...host.querySelectorAll('.cx-hcell')].map(cell => cell.textContent),
                            rows: [...host.querySelectorAll('.cx-row[data-cx-id]')].map(row => ({ id: row.dataset.cxId,
                                hiddenFields: [...row.querySelectorAll('.cx-hidden-fields')].map(button => button.textContent) })) };
                    }
                    """);
                File.WriteAllText(Path.Combine(output, name + ".json"), JsonSerializer.Serialize(layout, new JsonSerializerOptions { WriteIndented = true }));
                await session.Page.ScreenshotAsync(new() { Path = Path.Combine(output, name + ".png"), FullPage = true });
            }
            await session.Page.ExposeFunctionAsync("writeCanopyFixture", (string name, string base64, string contract) => {
                if (Path.GetFileName(name) != name || !new[] { ".xlsx", ".pdf", ".csv" }.Contains(Path.GetExtension(name))) throw new InvalidDataException("Invalid Canopy fixture name.");
                string path = Path.Combine(output, name); File.WriteAllBytes(path, Convert.FromBase64String(base64)); File.WriteAllText(path + ".json", contract);
                files.Add(CanopyExportVerifier.Verify(path)); return true;
            });
            await session.Page.ExposeFunctionAsync("captureCanopyScreen", async (string name) => {
                if (name != "records") throw new InvalidDataException("Invalid screenshot stage.");
                await CaptureAsync("records-wide", 1280, 720);
                await CaptureAsync("records-compact", 420, 760);
                await session.Page.SetViewportSizeAsync(1280, 720);
                return true;
            });
            await session.Page.AddStyleTagAsync(new() { Content = inputs["canopyx-reporting.css"] + inputs["canopyx-filter-editor.css"] + "body{padding:16px;box-sizing:border-box}h1{font-size:22px;font-weight:400}#fixture{height:620px}.cx-host{height:100%}" });
            await session.Page.AddScriptTagAsync(new() { Content = inputs["canopyx-reporting.js"] + "\n" + inputs["canopyx-filter-editor.js"] });
            await session.Page.AddScriptTagAsync(new() { Content = BrowserAssets.CanopyXScript.Content });
            await session.Page.AddScriptTagAsync(new() { Content = script });
            JsonElement result = await session.Page.EvaluateAsync<JsonElement>("args => runCanopyContracts(args)",
                new { regular, bold, workerScript = BrowserAssets.CanopyXScript.Content });
            await session.Page.Locator("#canopy-host-export").ClickAsync();
            await session.Page.WaitForFunctionAsync("() => globalThis.canopyHostButtonComplete || globalThis.canopyHostButtonError");
            string? hostError = await session.Page.EvaluateAsync<string?>("globalThis.canopyHostButtonError || null");
            if (hostError is not null) throw new InvalidDataException(hostError);
            if (errors.Count != 0) throw new InvalidDataException(string.Join("; ", errors));
            await CaptureAsync("canopy-wide", 1280, 720);
            await CaptureAsync("canopy-compact", 420, 760);
            await session.Page.GotoAsync(new Uri(host).AbsoluteUri);
            await session.Page.AddScriptTagAsync(new() { Content = inputs["canopyx-grid.js"] });
            await session.Page.AddScriptTagAsync(new() { Content = BrowserAssets.CanopyXScript.Content });
            await session.Page.AddScriptTagAsync(new() { Content = script });
            JsonElement ordinary = await session.Page.EvaluateAsync<JsonElement>("args => runOrdinaryCanopyContracts(args)", new { regular, bold });
            if (args.Contains("--canopy-scale") || args.Contains("--canopy-pdf-scale")) {
                await session.Page.GotoAsync(new Uri(host).AbsoluteUri);
                await session.Page.AddScriptTagAsync(new() { Content = inputs["canopyx-reporting.js"] });
                await session.Page.AddScriptTagAsync(new() { Content = BrowserAssets.CanopyXScript.Content });
                await session.Page.AddScriptTagAsync(new() { Content = script });
                await CanopyScale.RunAsync(session, repository, output, args);
            }
            var report = new { engine = engine.ToString(), browserVersion = session.Browser?.Version, canopy, inputs = hashes, asset = BrowserAssets.CanopyXScript.HashedFileName, result, ordinary, files, errors };
            reports.Add(report);
            File.WriteAllText(Path.Combine(output, "report.json"), JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
            Console.WriteLine($"CanopyX {engine}: {files.Count} independently read exports; native capture/host/paging/cancellation contracts.");
        }
        foreach (var pair in inputs) if (File.ReadAllText(Path.Combine(assets, pair.Key)) != pair.Value)
            throw new InvalidDataException("The CanopyX candidate changed during qualification; rerun the affected boundary against a frozen candidate.");
        File.WriteAllText(Path.Combine(evidence, "canopy-interop.json"), JsonSerializer.Serialize(reports, new JsonSerializerOptions { WriteIndented = true }));
    }
}
