using System.Security.Cryptography;
using System.Text.Json;
using HtmlTinkerX;

/// <summary>Qualifies composed analytical reports from real selected grid captures, using independent artifact readers.</summary>
internal static class HierarchicalReportInterop {
    internal static async Task RunAsync(string repository, string evidence, string[] args) {
        string Option(string name) => args.FirstOrDefault(argument => argument.StartsWith(name + "=", StringComparison.Ordinal))?[(name.Length + 1)..]
            ?? throw new ArgumentException("Analytical report qualification requires " + name + ".");
        var modules = new Uri(Option("--module-url").TrimEnd('/') + "/");
        var example = new Uri(Option("--example-url"));
        if (!modules.IsLoopback || modules.Scheme != "http" || !example.IsLoopback || example.Scheme != "http" || example.Authority != modules.Authority)
            throw new ArgumentException("Modules and the analytical example require the same loopback HTTP server.");
        string assets = Path.GetFullPath(Option("--comparison-assets")), canopy = Path.GetFullPath(Option("--canopy-assets"));
        using var comparison = JsonDocument.Parse(File.ReadAllText(Path.Combine(repository, "Build", "BrowserExports", "comparison-assets.json")));
        DataTablesInterop.VerifyAssets(assets, comparison.RootElement);
        string fixture = Path.Combine(repository, "Build", "BrowserExports", "hierarchical-report-contracts.js");
        string bundle = Path.Combine(repository, "OfficeIMO.JavaScript", "bundles", "officeimo.js");
        string moduleRoot = Path.Combine(repository, "OfficeIMO.JavaScript", "dist");
        var hashes = new Dictionary<string, string>();
        void Capture(string path) => hashes.Add(path, Hash(path));
        foreach (string path in new[] { fixture, bundle, Path.Combine(repository, "OfficeIMO.JavaScript.Examples", "analytical-report-job.mjs"),
            Path.Combine(repository, "OfficeIMO.JavaScript.Examples", "report-engine-job.mjs") }) Capture(path);
        foreach (string path in Directory.EnumerateFiles(moduleRoot, "*.js", SearchOption.AllDirectories)) Capture(path);
        var canopyInputs = new Dictionary<string, string>();
        foreach (string name in new[] { "canopyx-reporting.js", "canopyx-reporting.css", "canopyx-filter-editor.js", "canopyx-filter-editor.css", "canopyx-export.js" }) {
            string path = Path.Combine(canopy, name); canopyInputs.Add(name, File.ReadAllText(path)); Capture(path);
        }
        string Font(string name) {
            string path = Path.Combine(repository, "Website", "Apps", "OfficeIMO.Web.Converter", "Assets", "Fonts", "Carlito-" + name + ".ttf");
            Capture(path); return Convert.ToBase64String(File.ReadAllBytes(path));
        }
        string regular = Font("Regular"), bold = Font("Bold");
        string chartPath = Path.Combine(repository, "OfficeIMO.TestAssets", "JavaScript", "professional-service-status.png");
        string provenancePath = Path.ChangeExtension(chartPath, ".json"); Capture(chartPath); Capture(provenancePath);
        using var provenance = JsonDocument.Parse(File.ReadAllText(provenancePath));
        if (Hash(chartPath) != provenance.RootElement.GetProperty("sha256").GetString()) throw new InvalidDataException("Canonical ChartForgeX figure differs from its producer receipt.");
        string chart = Convert.ToBase64String(File.ReadAllBytes(chartPath));
        var chartSummary = provenance.RootElement.GetProperty("summary").EnumerateObject().ToDictionary(item => item.Name, item => item.Value.GetInt32());
        string? selectedStack = args.FirstOrDefault(argument => argument.StartsWith("--stack=", StringComparison.Ordinal))?[8..];
        (string Profile, string Stack)[] profiles = [("datatables", "bundled"), ("datatables", "current"), ("canopyx", "native")];
        profiles = profiles.Where(profile => selectedStack is null || selectedStack == profile.Stack).ToArray();
        if (profiles.Length == 0) throw new ArgumentException("Unknown analytical-report stack.");
        var reports = new List<object>();
        foreach (HtmlBrowserEngine engine in DataTablesInterop.Engines(args)) foreach (var profile in profiles) {
            string output = Path.Combine(evidence, engine.ToString().ToLowerInvariant(), profile.Profile + "-" + profile.Stack);
            Directory.CreateDirectory(output);
            await using var session = await HtmlBrowser.OpenSessionAsync(new Uri(modules, "../").AbsoluteUri,
                new HtmlBrowserLaunchOptions { Browser = engine, Headless = true, Timeout = 120000, Timezone = "Europe/Warsaw" });
            var errors = new List<string>(); session.Page.PageError += (_, error) => errors.Add(error);
            session.Page.RequestFailed += (_, request) => errors.Add("Request failed: " + request.Url + " " + request.Failure);
            var files = new List<object>();
            await session.Page.ExposeFunctionAsync("writeHierarchicalArtifact", (string name, string base64, string contract) => {
                if (Path.GetFileName(name) != name || !(name.EndsWith(".xlsx", StringComparison.Ordinal) || name.EndsWith(".pdf", StringComparison.Ordinal)))
                    throw new InvalidDataException("Invalid analytical-report fixture name.");
                string path = Path.Combine(output, name); File.WriteAllBytes(path, Convert.FromBase64String(base64)); File.WriteAllText(path + ".json", contract);
                files.Add(HierarchicalReportVerifier.Verify(path)); return true;
            });
            if (profile.Profile == "datatables") {
                foreach (JsonElement dependency in comparison.RootElement.GetProperty("stacks").GetProperty(profile.Stack).GetProperty("scripts").EnumerateArray())
                    await session.Page.AddScriptTagAsync(new() { Path = Path.Combine(assets, dependency.GetString()!) });
                await session.Page.AddScriptTagAsync(new() { Path = bundle });
            } else {
                await session.Page.AddStyleTagAsync(new() { Content = canopyInputs["canopyx-reporting.css"] + canopyInputs["canopyx-filter-editor.css"] });
                await session.Page.AddScriptTagAsync(new() { Content = canopyInputs["canopyx-reporting.js"] + "\n" + canopyInputs["canopyx-filter-editor.js"] + "\n" + canopyInputs["canopyx-export.js"] });
            }
            await session.Page.AddScriptTagAsync(new() { Path = fixture });
            JsonElement result;
            try {
                result = await session.Page.EvaluateAsync<JsonElement>("input => runHierarchicalReportContracts(input)", new {
                    moduleBase = modules.AbsoluteUri.TrimEnd('/'), exampleUrl = example.AbsoluteUri, inventoryUrl = new Uri(example, "report-engine-job.mjs").AbsoluteUri,
                    profile = profile.Profile, stack = profile.Stack, regular, bold, chart, chartSummary
                });
            } catch (Exception error) {
                File.WriteAllText(Path.Combine(output, "failure.json"), JsonSerializer.Serialize(new { error = error.Message, errors }, new JsonSerializerOptions { WriteIndented = true }));
                throw;
            }
            if (errors.Count != 0) throw new InvalidDataException("Analytical-report browser errors: " + string.Join("; ", errors));
            if (files.Count != (profile.Stack == "current" ? 6 : 4)) throw new InvalidDataException("Analytical source did not produce every required delivery artifact.");
            var report = new { engine = engine.ToString(), browser = session.Browser?.Version, profile.Profile, profile.Stack, result, files, errors };
            File.WriteAllText(Path.Combine(output, "report.json"), JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
            reports.Add(report); Console.WriteLine($"Analytical reports passed {engine}/{profile.Profile}/{profile.Stack}: selected three-level capture, composition and independent numeric reads.");
        }
        if (hashes.Any(pair => Hash(pair.Key) != pair.Value)) throw new InvalidDataException("Analytical-report source/input graph changed during qualification.");
        File.WriteAllText(Path.Combine(evidence, "hierarchical-report-interop.json"), JsonSerializer.Serialize(new {
            sources = hashes.ToDictionary(pair => Path.GetRelativePath(repository, pair.Key), pair => pair.Value), chartSummary, reports
        }, new JsonSerializerOptions { WriteIndented = true }));
    }
    private static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();
}
