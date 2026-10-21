using System.Security.Cryptography;
using System.Text.Json;
using HtmlTinkerX;

/// <summary>Real selected DataTables/Canopy sources exercise one shared professional report through Blob and awaited sink delivery.</summary>
internal static class ProfessionalPdfInterop {
    internal static async Task RunAsync(string repository, string evidence, string[] args) {
        string Option(string name) => args.FirstOrDefault(argument => argument.StartsWith(name + "=", StringComparison.Ordinal))?[(name.Length + 1)..]
            ?? throw new ArgumentException("Professional PDF qualification requires " + name + ".");
        var modules = new Uri(Option("--module-url").TrimEnd('/') + "/");
        var example = new Uri(Option("--example-url"));
        if (!modules.IsLoopback || modules.Scheme != "http" || !example.IsLoopback || example.Scheme != "http" || example.Authority != modules.Authority)
            throw new ArgumentException("Current modules and executable example require the same loopback HTTP server.");
        string assets = Path.GetFullPath(Option("--comparison-assets"));
        string canopy = Path.Combine(Path.GetFullPath(Option("--canopy")), "CanopyX.Html", "Assets");
        using var comparison = JsonDocument.Parse(File.ReadAllText(Path.Combine(repository, "Build", "BrowserExports", "comparison-assets.json")));
        DataTablesInterop.VerifyAssets(assets, comparison.RootElement);
        string scriptPath = Path.Combine(repository, "Build", "BrowserExports", "professional-pdf-contracts.js"), script = File.ReadAllText(scriptPath);
        string examplePath = Path.Combine(repository, "OfficeIMO.JavaScript.Examples", "report-engine-job.mjs");
        string bundlePath = Path.Combine(repository, "OfficeIMO.JavaScript", "bundles", "officeimo.js"), bundle = File.ReadAllText(bundlePath);
        string moduleRoot = Path.Combine(repository, "OfficeIMO.JavaScript", "dist");
        var moduleHashes = Directory.EnumerateFiles(moduleRoot, "*.js", SearchOption.AllDirectories).ToDictionary(path => Path.GetRelativePath(moduleRoot, path), Hash);
        string exampleHash = Hash(examplePath), bundleHash = Hash(bundlePath), scriptHash = Hash(scriptPath);
        var canopyInputs = new Dictionary<string, string>();
        foreach (string name in new[] { "canopyx-reporting.js", "canopyx-reporting.css", "canopyx-filter-editor.js", "canopyx-filter-editor.css", "canopyx-export.js" })
            canopyInputs.Add(name, File.ReadAllText(Path.Combine(canopy, name)));
        var canopyHashes = canopyInputs.ToDictionary(pair => pair.Key, pair => Hash(Path.Combine(canopy, pair.Key)));
        string Font(string name) => Convert.ToBase64String(File.ReadAllBytes(Path.Combine(repository, "Website", "Apps", "OfficeIMO.Web.Converter", "Assets", "Fonts", "Carlito-" + name + ".ttf")));
        string chartPath = Path.Combine(repository, "OfficeIMO.TestAssets", "JavaScript", "professional-service-status.png");
        string chartProvenancePath = Path.ChangeExtension(chartPath, ".json");
        using var chartProvenance = JsonDocument.Parse(File.ReadAllText(chartProvenancePath));
        string chartHash = Hash(chartPath), chartProvenanceHash = Hash(chartProvenancePath);
        if (chartHash != chartProvenance.RootElement.GetProperty("sha256").GetString())
            throw new InvalidDataException("The qualified service-status figure differs from its producer receipt.");
        string chart = Convert.ToBase64String(File.ReadAllBytes(chartPath));
        var chartSummary = chartProvenance.RootElement.GetProperty("summary").EnumerateObject().ToDictionary(item => item.Name, item => item.Value.GetInt32());
        var reports = new List<object>();
        foreach (HtmlBrowserEngine engine in DataTablesInterop.Engines(args)) foreach (string profile in new[] { "datatables", "canopyx" }) {
            string output = Path.Combine(evidence, engine.ToString().ToLowerInvariant(), profile); Directory.CreateDirectory(output);
            await using var session = await HtmlBrowser.OpenSessionAsync(new Uri(modules, "../").AbsoluteUri,
                new HtmlBrowserLaunchOptions { Browser = engine, Headless = true, Timeout = 120000, Timezone = "Europe/Warsaw" });
            var errors = new List<string>(); session.Page.PageError += (_, error) => errors.Add(error);
            var files = new List<object>();
            await session.Page.ExposeFunctionAsync("writeProfessionalPdf", (string name, string base64, string contract) => {
                if (Path.GetFileName(name) != name || !name.EndsWith(".pdf", StringComparison.Ordinal)) throw new InvalidDataException("Invalid professional PDF fixture name.");
                string path = Path.Combine(output, name); File.WriteAllBytes(path, Convert.FromBase64String(base64)); File.WriteAllText(path + ".json", contract);
                files.Add(ProfessionalPdfVerifier.Verify(path)); return true;
            });
            await session.Page.AddStyleTagAsync(new() { Content = "body{font:16px system-ui;margin:20px}table{border-collapse:collapse}th,td{padding:5px}#professional-grid{height:500px}" });
            if (profile == "datatables") {
                foreach (JsonElement dependency in comparison.RootElement.GetProperty("stacks").GetProperty("current").GetProperty("scripts").EnumerateArray())
                    await session.Page.AddScriptTagAsync(new() { Path = Path.Combine(assets, dependency.GetString()!) });
                await session.Page.AddScriptTagAsync(new() { Content = bundle });
            } else {
                await session.Page.AddStyleTagAsync(new() { Content = canopyInputs["canopyx-reporting.css"] + canopyInputs["canopyx-filter-editor.css"] });
                await session.Page.AddScriptTagAsync(new() { Content = canopyInputs["canopyx-reporting.js"] + "\n" + canopyInputs["canopyx-filter-editor.js"] + "\n" + canopyInputs["canopyx-export.js"] });
            }
            await session.Page.AddScriptTagAsync(new() { Content = script });
            JsonElement result = await session.Page.EvaluateAsync<JsonElement>("input => runProfessionalPdfContracts(input)", new {
                moduleBase = modules.AbsoluteUri.TrimEnd('/'), exampleUrl = example.AbsoluteUri, profile, regular = Font("Regular"), bold = Font("Bold"), chart, chartSummary
            });
            if (errors.Count != 0) throw new InvalidDataException("Professional PDF browser errors: " + string.Join("; ", errors));
            if (files.Count != 2) throw new InvalidDataException("Both real Blob and sink outputs are required.");
            var report = new { engine = engine.ToString(), browser = session.Browser?.Version, profile, result, files, errors };
            File.WriteAllText(Path.Combine(output, "report.json"), JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
            reports.Add(report); Console.WriteLine($"Professional PDF passed {engine}/{profile}: both independently read reports, slow-sink backpressure and cancellation.");
        }
        if (Hash(examplePath) != exampleHash || Hash(bundlePath) != bundleHash || Hash(scriptPath) != scriptHash || moduleHashes.Any(pair => Hash(Path.Combine(moduleRoot, pair.Key)) != pair.Value) ||
            canopyHashes.Any(pair => Hash(Path.Combine(canopy, pair.Key)) != pair.Value) || Hash(chartPath) != chartHash || Hash(chartProvenancePath) != chartProvenanceHash)
            throw new InvalidDataException("Professional PDF source/input graph changed during qualification.");
        File.WriteAllText(Path.Combine(evidence, "professional-pdf-interop.json"), JsonSerializer.Serialize(new {
            exampleSha256 = exampleHash, classicBundleSha256 = bundleHash, fixtureSha256 = scriptHash, modules = moduleHashes, canopy = canopyHashes,
            dataTablesStack = "current", chartSha256 = chartHash, chartProvenanceSha256 = chartProvenanceHash, chartSummary, reports
        }, new JsonSerializerOptions { WriteIndented = true }));
    }
    private static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();
}
