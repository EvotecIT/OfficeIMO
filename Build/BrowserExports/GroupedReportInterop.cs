using System.Security.Cryptography;
using System.Text.Json;
using HtmlTinkerX;

/// <summary>Qualifies grouped reports from real filtered, selected and reordered grid captures.</summary>
internal static class GroupedReportInterop {
    internal static async Task RunAsync(string repository, string evidence, string[] args) {
        string Option(string name) => args.FirstOrDefault(argument => argument.StartsWith(name + "=", StringComparison.Ordinal))?[(name.Length + 1)..]
            ?? throw new ArgumentException("Grouped report qualification requires " + name + ".");
        var modules = new Uri(Option("--module-url").TrimEnd('/') + "/");
        var example = new Uri(Option("--example-url"));
        if (!modules.IsLoopback || modules.Scheme != "http" || !example.IsLoopback || example.Scheme != "http" || example.Authority != modules.Authority)
            throw new ArgumentException("Modules and the executable example require the same loopback HTTP server.");
        string assets = Path.GetFullPath(Option("--comparison-assets"));
        string canopy = Path.Combine(Path.GetFullPath(Option("--canopy")), "CanopyX.Html", "Assets");
        using var comparison = JsonDocument.Parse(File.ReadAllText(Path.Combine(repository, "Build", "BrowserExports", "comparison-assets.json")));
        DataTablesInterop.VerifyAssets(assets, comparison.RootElement);
        string fixturePath = Path.Combine(repository, "Build", "BrowserExports", "grouped-report-contracts.js");
        string bundlePath = Path.Combine(repository, "OfficeIMO.JavaScript", "bundles", "officeimo.js");
        string examplePath = Path.Combine(repository, "OfficeIMO.JavaScript.Examples", "report-engine-job.mjs");
        string moduleRoot = Path.Combine(repository, "OfficeIMO.JavaScript", "dist");
        var hashes = new Dictionary<string, string> { [fixturePath] = Hash(fixturePath), [bundlePath] = Hash(bundlePath), [examplePath] = Hash(examplePath) };
        foreach (string file in Directory.EnumerateFiles(moduleRoot, "*.js", SearchOption.AllDirectories)) hashes.Add(file, Hash(file));
        var canopyInputs = new Dictionary<string, string>();
        foreach (string name in new[] { "canopyx-reporting.js", "canopyx-reporting.css", "canopyx-filter-editor.js", "canopyx-filter-editor.css", "canopyx-export.js" }) {
            string path = Path.Combine(canopy, name); canopyInputs.Add(name, File.ReadAllText(path)); hashes.Add(path, Hash(path));
        }
        string Font(string name) {
            string path = Path.Combine(repository, "Website", "Apps", "OfficeIMO.Web.Converter", "Assets", "Fonts", "Carlito-" + name + ".ttf");
            hashes[path] = Hash(path); return Convert.ToBase64String(File.ReadAllBytes(path));
        }
        string regular = Font("Regular"), bold = Font("Bold");
        string? selectedStack = args.FirstOrDefault(argument => argument.StartsWith("--stack=", StringComparison.Ordinal))?[8..];
        bool emptyOnly = args.Contains("--grouped-empty-only");
        (string Profile, string Stack)[] profiles = [("datatables", "bundled"), ("datatables", "current"), ("canopyx", "native")];
        profiles = profiles.Where(profile => (selectedStack is null || selectedStack == profile.Stack) && (!emptyOnly || profile.Stack == "current")).ToArray();
        if (profiles.Length == 0) throw new ArgumentException("Unknown grouped-report stack.");
        var reports = new List<object>();
        foreach (HtmlBrowserEngine engine in DataTablesInterop.Engines(args)) foreach (var profile in profiles) {
            string output = Path.Combine(evidence, engine.ToString().ToLowerInvariant(), profile.Profile + "-" + profile.Stack);
            Directory.CreateDirectory(output);
            await using var session = await HtmlBrowser.OpenSessionAsync(new Uri(modules, "../").AbsoluteUri,
                new HtmlBrowserLaunchOptions { Browser = engine, Headless = true, Timeout = 120000, Timezone = "Europe/Warsaw" });
            var errors = new List<string>(); session.Page.PageError += (_, error) => errors.Add(error);
            var files = new List<object>();
            await session.Page.ExposeFunctionAsync("writeGroupedArtifact", (string name, string base64, string contract) => {
                if (Path.GetFileName(name) != name || !(name.EndsWith(".xlsx", StringComparison.Ordinal) || name.EndsWith(".pdf", StringComparison.Ordinal)))
                    throw new InvalidDataException("Invalid grouped-report fixture name.");
                string path = Path.Combine(output, name); File.WriteAllBytes(path, Convert.FromBase64String(base64)); File.WriteAllText(path + ".json", contract);
                files.Add(GroupedReportVerifier.Verify(path)); return true;
            });
            if (profile.Profile == "datatables") {
                foreach (JsonElement dependency in comparison.RootElement.GetProperty("stacks").GetProperty(profile.Stack).GetProperty("scripts").EnumerateArray())
                    await session.Page.AddScriptTagAsync(new() { Path = Path.Combine(assets, dependency.GetString()!) });
                await session.Page.AddScriptTagAsync(new() { Path = bundlePath });
            } else {
                await session.Page.AddStyleTagAsync(new() { Content = canopyInputs["canopyx-reporting.css"] + canopyInputs["canopyx-filter-editor.css"] });
                await session.Page.AddScriptTagAsync(new() { Content = canopyInputs["canopyx-reporting.js"] + "\n" + canopyInputs["canopyx-filter-editor.js"] + "\n" + canopyInputs["canopyx-export.js"] });
            }
            await session.Page.AddScriptTagAsync(new() { Path = fixturePath });
            JsonElement result = await session.Page.EvaluateAsync<JsonElement>("input => runGroupedReportContracts(input)", new {
                moduleBase = modules.AbsoluteUri.TrimEnd('/'), exampleUrl = example.AbsoluteUri,
                profile = profile.Profile, stack = profile.Stack, regular, bold, emptyOnly
            });
            if (errors.Count != 0) throw new InvalidDataException("Grouped-report browser errors: " + string.Join("; ", errors));
            int expectedFiles = emptyOnly ? 3 : profile.Stack == "current" ? 8 : 3;
            if (files.Count != expectedFiles) throw new InvalidDataException("Grouped source did not produce all required delivery artifacts.");
            var report = new { engine = engine.ToString(), browser = session.Browser?.Version, profile.Profile, profile.Stack, result, files, errors };
            File.WriteAllText(Path.Combine(output, "report.json"), JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
            reports.Add(report); Console.WriteLine($"Grouped reports passed {engine}/{profile.Profile}/{profile.Stack}: selected raw capture, summaries and independent artifact reads.");
        }
        if (hashes.Any(pair => Hash(pair.Key) != pair.Value)) throw new InvalidDataException("Grouped-report source/input graph changed during qualification.");
        File.WriteAllText(Path.Combine(evidence, "grouped-report-interop.json"), JsonSerializer.Serialize(new {
            sources = hashes.ToDictionary(pair => Path.GetRelativePath(repository, pair.Key), pair => pair.Value), reports
        }, new JsonSerializerOptions { WriteIndented = true }));
    }
    private static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();
}
