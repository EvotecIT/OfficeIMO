using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using HtmlTinkerX;
using OfficeIMO.Browser;

/// <summary>Optional third-party grid interoperability through HtmlTinkerX-owned browser sessions.</summary>
internal static class DataTablesInterop {
    internal static async Task RunAsync(string repository, string evidence, string[] args) {
        string assets = Path.GetFullPath(args.FirstOrDefault(a => a.StartsWith("--comparison-assets=", StringComparison.Ordinal))?.Substring(20)
            ?? throw new ArgumentException("Pass --comparison-assets=<verified-asset-directory>."));
        using JsonDocument manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(repository, "Build", "BrowserExports", "comparison-assets.json")));
        VerifyAssets(assets, manifest.RootElement);
        foreach (string stack in new[] { "bundled", "current" }) foreach (HtmlBrowserEngine engine in Engines(args)) {
            string output = Path.Combine(evidence, stack, engine.ToString().ToLowerInvariant()); Directory.CreateDirectory(output);
            Console.WriteLine($"DataTables interoperability {stack}/{engine}");
            await using HtmlBrowserSession session = await OpenAsync(assets, manifest.RootElement, stack, engine, output);
            var errors = new List<string>(); session.Page.PageError += (_, error) => errors.Add(error);
            await session.Page.ExposeFunctionAsync("writeFixture", (string name, string base64) => {
                if (Path.GetFileName(name) != name) throw new InvalidDataException("Invalid fixture name.");
                File.WriteAllBytes(Path.Combine(output, name), Convert.FromBase64String(base64)); return true;
            });
            await session.Page.AddScriptTagAsync(new() { Content = File.ReadAllText(Path.Combine(repository, "Build", "BrowserExports", "datatables-contracts.js")) });
            await session.Page.AddScriptTagAsync(new() { Content = File.ReadAllText(Path.Combine(repository, "Build", "BrowserExports", "datatables-worker.js")) });
            JsonElement result = await session.Page.EvaluateAsync<JsonElement>("script => runDataTablesContracts(script)", BrowserAssets.Script.Content);
            foreach (string workbook in Directory.GetFiles(output, "*.xlsx")) WorkbookVerifier.Verify(workbook);
            if (errors.Count != 0) throw new InvalidDataException(string.Join("; ", errors));
            await session.Page.ScreenshotAsync(new() { Path = Path.Combine(output, "table-wide.png"), FullPage = true });
            await session.Page.SetViewportSizeAsync(420, 760);
            await session.Page.ScreenshotAsync(new() { Path = Path.Combine(output, "table-compact.png"), FullPage = true });
            File.WriteAllText(Path.Combine(output, "report.json"), JsonSerializer.Serialize(new { engine = engine.ToString(), browser = session.Browser?.Version, result, errors }));
            Console.WriteLine($"Passed {stack}/{engine}: {Directory.GetFiles(output, "*.xlsx").Length} SDK-validated workbooks.");
        }
    }

    internal static HtmlBrowserEngine[] Engines(string[] args) {
        string? selected = args.FirstOrDefault(a => a.StartsWith("--engine=", StringComparison.Ordinal))?.Substring(9);
        var result = new[] { HtmlBrowserEngine.Chromium, HtmlBrowserEngine.Firefox, HtmlBrowserEngine.WebKit }
            .Where(engine => selected is null || string.Equals(selected, engine.ToString(), StringComparison.OrdinalIgnoreCase)).ToArray();
        return result.Length != 0 ? result : throw new ArgumentException("Unknown browser engine.");
    }

    internal static void VerifyAssets(string assets, JsonElement manifest) {
        foreach (JsonElement asset in manifest.GetProperty("assets").EnumerateArray()) {
            string name = asset.GetProperty("name").GetString()!;
            if (Path.GetFileName(name) != name) throw new InvalidDataException("Invalid comparison asset path.");
            using Stream source = File.OpenRead(Path.Combine(assets, name));
            string hash = Convert.ToHexString(SHA256.HashData(source)).ToLowerInvariant();
            if (hash != asset.GetProperty("sha256").GetString()) throw new InvalidDataException("Comparison asset hash differs: " + name);
        }
    }

    internal static async Task<HtmlBrowserSession> OpenAsync(string assets, JsonElement manifest, string stack, HtmlBrowserEngine engine, string output) {
        string host = Path.Combine(output, "host.html");
        File.WriteAllText(host, "<!doctype html><meta charset=\"utf-8\"><title>Table export qualification</title><style>body{font:16px system-ui;margin:24px}table{border-collapse:collapse}td,th{padding:8px;border:1px solid #ccc}button{padding:8px;margin:4px}.selected{background:#e2f0d9}</style>", new UTF8Encoding(false));
        var launch = new HtmlBrowserLaunchOptions { Browser = engine, Headless = true, Timeout = 120000, LoadState = HtmlBrowserLoadState.Load };
        if (engine == HtmlBrowserEngine.Chromium) launch.BrowserArguments.Add("--enable-precise-memory-info");
        HtmlBrowserSession session = await HtmlBrowser.OpenSessionAsync(new Uri(host).AbsoluteUri, launch);
        try {
            foreach (JsonElement script in manifest.GetProperty("stacks").GetProperty(stack).GetProperty("scripts").EnumerateArray())
                await session.Page.AddScriptTagAsync(new() { Path = Path.Combine(assets, script.GetString()!) });
            await session.Page.AddScriptTagAsync(new() { Content = BrowserAssets.DataTablesScript.Content });
            return session;
        } catch { await session.DisposeAsync(); throw; }
    }
}
