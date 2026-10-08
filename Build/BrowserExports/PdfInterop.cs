using System.Text.Json;
using HtmlTinkerX;
using OfficeIMO.Browser;

/// <summary>Opt-in PDF qualification using HtmlTinkerX-owned browsers and the independent managed reader.</summary>
internal static class PdfInterop {
    internal static async Task RunAsync(string repository, string evidence, string[] args) {
        string script = File.ReadAllText(Path.Combine(repository, "Build", "BrowserExports", "pdf-contracts.js"));
        string regular = Convert.ToBase64String(File.ReadAllBytes(Path.Combine(repository, "Website", "Apps", "OfficeIMO.Web.Converter", "Assets", "Fonts", "Carlito-Regular.ttf")));
        string bold = Convert.ToBase64String(File.ReadAllBytes(Path.Combine(repository, "Website", "Apps", "OfficeIMO.Web.Converter", "Assets", "Fonts", "Carlito-Bold.ttf")));
        string symbols = Convert.ToBase64String(File.ReadAllBytes(Path.Combine(repository, "Website", "Apps", "OfficeIMO.Web.Converter", "Assets", "Fonts", "NotoSansSymbols2-Regular.ttf")));
        string japanese = Convert.ToBase64String(File.ReadAllBytes(Path.Combine(repository, "Website", "Apps", "OfficeIMO.Web.Converter", "Assets", "Fonts", "NotoSansJP-OfficeIMO-Common.ttf")));
        string? assets = args.FirstOrDefault(a => a.StartsWith("--comparison-assets=", StringComparison.Ordinal))?.Substring(20);
        using JsonDocument manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(repository, "Build", "BrowserExports", "comparison-assets.json")));
        if (assets is not null) DataTablesInterop.VerifyAssets(assets, manifest.RootElement);
        string stack = args.FirstOrDefault(a => a.StartsWith("--stack=", StringComparison.Ordinal))?.Substring(8) ?? "current";
        if (stack is not "current" and not "bundled") throw new ArgumentException("Unknown comparison stack.");
        var reports = new List<object>();
        foreach (HtmlBrowserEngine engine in DataTablesInterop.Engines(args)) {
            string output = Path.Combine(evidence, engine.ToString().ToLowerInvariant()); Directory.CreateDirectory(output);
            string host = Path.Combine(output, "host.html"); File.WriteAllText(host, "<!doctype html><meta charset=utf-8><title>PDF export qualification</title>");
            await using HtmlBrowserSession session = assets is not null
                ? await DataTablesInterop.OpenAsync(assets, manifest.RootElement, stack, engine, output)
                : await HtmlBrowser.OpenSessionAsync(new Uri(host).AbsoluteUri, new HtmlBrowserLaunchOptions { Browser = engine, Headless = true, Timeout = 120000 });
            var errors = new List<string>(); session.Page.PageError += (_, error) => errors.Add(error);
            var files = new List<object>(); long removedBytes = 0;
            await session.Page.AddScriptTagAsync(new() { Content = BrowserAssets.Script.Content });
            await session.Page.ExposeFunctionAsync("writePdfFixture", (string name, string base64, string contract) => {
                if (Path.GetFileName(name) != name || !name.EndsWith(".pdf", StringComparison.Ordinal)) throw new InvalidDataException("Invalid PDF fixture name.");
                File.WriteAllBytes(Path.Combine(output, name), Convert.FromBase64String(base64));
                File.WriteAllText(Path.Combine(output, name + ".json"), contract);
                files.Add(PdfExportVerifier.Verify(Path.Combine(output, name)));
                if (name.StartsWith("scale-", StringComparison.Ordinal)) {
                    removedBytes += new FileInfo(Path.Combine(output, name)).Length + new FileInfo(Path.Combine(output, name + ".json")).Length;
                    File.Delete(Path.Combine(output, name)); File.Delete(Path.Combine(output, name + ".json"));
                }
                return true;
            });
            await session.Page.AddScriptTagAsync(new() { Content = script });
            JsonElement result = await session.Page.EvaluateAsync<JsonElement>("args => runPdfContracts(args)",
                new { regular, bold, symbols, japanese, workerScript = BrowserAssets.Script.Content, scale = args.Contains("--scale") });
            if (errors.Count != 0) throw new InvalidDataException(string.Join("; ", errors));
            var report = new { engine = engine.ToString(), browser = session.Browser?.Version, stack, asset = BrowserAssets.Script.HashedFileName, result, files, removedBytes, errors };
            reports.Add(report);
            File.WriteAllText(Path.Combine(output, "report.json"), JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
            Console.WriteLine($"PDF passed {engine}: {files.Count} independent-reader-validated documents.");
        }
        File.WriteAllText(Path.Combine(evidence, "pdf-interop.json"), JsonSerializer.Serialize(reports, new JsonSerializerOptions { WriteIndented = true }));
    }
}
