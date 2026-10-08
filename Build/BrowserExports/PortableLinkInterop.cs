using System.Text.Json;
using HtmlTinkerX;
using OfficeIMO.Browser;

/// <summary>Cross-engine portable-link correctness using the shared browser owner.</summary>
internal static class PortableLinkInterop {
    internal static async Task RunAsync(string repository, string evidence, string[] args) {
        string script = File.ReadAllText(Path.Combine(repository, "Build", "BrowserExports", "portable-links.js"));
        var reports = new List<object>();
        foreach (var engine in DataTablesInterop.Engines(args)) {
            string output = Path.Combine(evidence, engine.ToString().ToLowerInvariant()); Directory.CreateDirectory(output);
            string host = Path.Combine(output, "host.html"); File.WriteAllText(host, "<!doctype html><meta charset=utf-8><title>Portable cell links</title>");
            await using var session = await HtmlBrowser.OpenSessionAsync(new Uri(host).AbsoluteUri,
                new HtmlBrowserLaunchOptions { Browser = engine, Headless = true, Timeout = 120000 });
            var errors = new List<string>(); session.Page.PageError += (_, error) => errors.Add(error);
            var files = new List<object>();
            await session.Page.ExposeFunctionAsync("writePortableLinkFixture", (string name, string base64) => {
                if (Path.GetFileName(name) != name) throw new InvalidDataException("Invalid portable-link fixture name.");
                string path = Path.Combine(output, name); File.WriteAllBytes(path, Convert.FromBase64String(base64));
                files.Add(PortableLinkVerifier.Verify(path)); return true;
            });
            await session.Page.AddScriptTagAsync(new() { Content = BrowserAssets.Script.Content });
            await session.Page.AddScriptTagAsync(new() { Content = script });
            var result = await session.Page.EvaluateAsync<JsonElement>("script => runPortableLinkContracts(script)", BrowserAssets.Script.Content);
            if (errors.Count != 0) throw new InvalidDataException(string.Join("; ", errors));
            reports.Add(new { engine = engine.ToString(), browserVersion = session.Browser?.Version, result, files, errors });
            Console.WriteLine($"Portable links {engine}: {files.Count} independently read documents, classic/fallback/worker CSV contracts.");
        }
        File.WriteAllText(Path.Combine(evidence, "portable-link-browser.json"), JsonSerializer.Serialize(reports, new JsonSerializerOptions { WriteIndented = true }));
    }
}
