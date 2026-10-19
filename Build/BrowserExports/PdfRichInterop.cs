using System.Security.Cryptography;
using System.Text.Json;
using HtmlTinkerX;

/// <summary>Opt-in composed PDF contracts against current ES modules and classic bundles in HtmlTinkerX-owned browsers.</summary>
internal static class PdfRichInterop {
    internal static async Task RunAsync(string repository, string evidence, string[] args) {
        string moduleBase = args.FirstOrDefault(a => a.StartsWith("--module-url=", StringComparison.Ordinal))?.Substring(13)
            ?? throw new ArgumentException("--pdf-rich requires --module-url pointing to the current dist modules on a loopback HTTP server.");
        var moduleUri = new Uri(moduleBase.TrimEnd('/') + "/");
        if (!moduleUri.IsLoopback || moduleUri.Scheme != "http") throw new ArgumentException("Rich PDF module qualification requires loopback HTTP.");
        string script = File.ReadAllText(Path.Combine(repository, "Build", "BrowserExports", "pdf-rich-contracts.js"));
        string fixtureRoot = Path.Combine(repository, "OfficeIMO.JavaScript", "test", "fixtures", "pdf");
        var fixtures = new Dictionary<string, string>();
        foreach (string name in new[] { "rgb-pillow.png", "rgba-pillow.png", "rgba-filters.png" })
            fixtures[name] = Convert.ToBase64String(File.ReadAllBytes(Path.Combine(fixtureRoot, name)));
        string drawingFont = Convert.ToBase64String(File.ReadAllBytes(Path.Combine(repository, "Website", "Apps", "OfficeIMO.Web.Converter", "Assets", "Fonts", "Carlito-Regular.ttf")));
        string moduleRoot = Path.Combine(repository, "OfficeIMO.JavaScript", "dist");
        var hashes = ModuleHashes(moduleRoot);
        string bundlePath = Path.Combine(repository, "OfficeIMO.JavaScript", "bundles", "officeimo.js"), bundle = File.ReadAllText(bundlePath);
        string bundleHash = FileHash(bundlePath);
        var reports = new List<object>();
        foreach (HtmlBrowserEngine engine in DataTablesInterop.Engines(args)) {
            string output = Path.Combine(evidence, engine.ToString().ToLowerInvariant()); Directory.CreateDirectory(output);
            await using HtmlBrowserSession session = await HtmlBrowser.OpenSessionAsync(new Uri(moduleUri, "../host.html").AbsoluteUri,
                new HtmlBrowserLaunchOptions { Browser = engine, Headless = true, Timeout = 120000 });
            var errors = new List<string>(); session.Page.PageError += (_, error) => errors.Add(error);
            var files = new List<object>();
            await session.Page.ExposeFunctionAsync("writeRichPdfFixture", (string name, string base64, string contract) => {
                if (Path.GetFileName(name) != name || !name.EndsWith(".pdf", StringComparison.Ordinal)) throw new InvalidDataException("Invalid rich PDF fixture name.");
                string path = Path.Combine(output, name);
                File.WriteAllBytes(path, Convert.FromBase64String(base64)); File.WriteAllText(path + ".json", contract);
                files.Add(PdfExportVerifier.Verify(path)); return true;
            });
            await session.Page.AddScriptTagAsync(new() { Content = bundle });
            await session.Page.AddScriptTagAsync(new() { Content = script });
            var results = new Dictionary<string, JsonElement>();
            foreach (bool classic in new[] { false, true })
                results[classic ? "classic" : "esm"] = await session.Page.EvaluateAsync<JsonElement>("args => runRichPdfContracts(args)",
                    new { moduleBase = moduleUri.AbsoluteUri.TrimEnd('/'), fixtures, drawingFont, classic });
            if (errors.Count != 0) throw new InvalidDataException(string.Join("; ", errors));
            var report = new { engine = engine.ToString(), browser = session.Browser?.Version, results, files, errors };
            reports.Add(report);
            File.WriteAllText(Path.Combine(output, "report.json"), JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
            Console.WriteLine($"Rich PDF passed {engine}: {files.Count} independent-reader-validated documents.");
        }
        if (!hashes.OrderBy(pair => pair.Key).SequenceEqual(ModuleHashes(moduleRoot).OrderBy(pair => pair.Key)))
            throw new InvalidDataException("ES modules changed during browser qualification.");
        if (bundleHash != FileHash(bundlePath)) throw new InvalidDataException("Classic bundle changed during browser qualification.");
        File.WriteAllText(Path.Combine(evidence, "pdf-rich-interop.json"), JsonSerializer.Serialize(new { modules = hashes, bundles = new Dictionary<string, string> { ["officeimo.js"] = bundleHash }, reports }, new JsonSerializerOptions { WriteIndented = true }));
    }
    private static Dictionary<string, string> ModuleHashes(string root) => Directory.EnumerateFiles(root, "*.js", SearchOption.AllDirectories)
        .ToDictionary(path => Path.GetRelativePath(root, path).Replace('\\', '/'), FileHash);
    private static string FileHash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();
}
