using System.Security.Cryptography;
using System.Text.Json;
using HtmlTinkerX;

/// <summary>Opt-in real grid/worker producers for independently verified semantic PDF reports.</summary>
internal static class SemanticReportInterop {
    internal static async Task RunAsync(string repository, string evidence, string[] args) {
        string Option(string name) => args.FirstOrDefault(value => value.StartsWith(name + "=", StringComparison.Ordinal))?[(name.Length + 1)..]
            ?? throw new ArgumentException("Semantic report qualification requires " + name + ".");
        var modules = new Uri(Option("--module-url").TrimEnd('/') + "/");
        if (!modules.IsLoopback || modules.Scheme != "http") throw new ArgumentException("Semantic modules require loopback HTTP.");
        string comparisonAssets = Path.GetFullPath(Option("--comparison-assets")), canopyAssets = Path.GetFullPath(Option("--canopy-assets"));
        using var comparison = JsonDocument.Parse(File.ReadAllText(Path.Combine(repository, "Build", "BrowserExports", "comparison-assets.json")));
        DataTablesInterop.VerifyAssets(comparisonAssets, comparison.RootElement);
        string fixture = Path.Combine(repository, "Build", "BrowserExports", "semantic-report-contracts.js");
        string bundle = Path.Combine(repository, "OfficeIMO.JavaScript", "bundles", "officeimo.js");
        var hashes = new Dictionary<string, string>();
        void Capture(string path) => hashes.TryAdd(path, Hash(path));
        Capture(fixture); Capture(bundle);
        foreach (string path in Directory.EnumerateFiles(Path.Combine(repository, "OfficeIMO.JavaScript", "dist"), "*.js", SearchOption.AllDirectories)) Capture(path);
        var canopyInputs = new Dictionary<string, string>();
        foreach (string name in new[] { "canopyx-reporting.js", "canopyx-reporting.css", "canopyx-filter-editor.js", "canopyx-filter-editor.css", "canopyx-export.js" }) {
            string path = Path.Combine(canopyAssets, name); Capture(path); canopyInputs.Add(name, File.ReadAllText(path));
        }
        string Font(string name) {
            string path = Path.Combine(repository, "Website", "Apps", "OfficeIMO.Web.Converter", "Assets", "Fonts", "Carlito-" + name + ".ttf");
            Capture(path); return Convert.ToBase64String(File.ReadAllBytes(path));
        }
        string regular = Font("Regular"), bold = Font("Bold");
        (string Profile, string Stack)[] profiles = [("hfx", "worker"), ("datatables", "bundled"), ("datatables", "current"), ("canopyx", "offset"), ("canopyx", "cursor")];
        string? selected = args.FirstOrDefault(value => value.StartsWith("--semantic-profile=", StringComparison.Ordinal))?[19..];
        string language = args.FirstOrDefault(value => value.StartsWith("--semantic-language=", StringComparison.Ordinal))?[20..] ?? "en";
        if (language is not ("en" or "pl")) throw new ArgumentException("Semantic report language must be en or pl.");
        if (language == "pl" && selected != "hfx-worker") throw new ArgumentException("The Polish native fixture requires --semantic-profile=hfx-worker; the normal fifteen-profile matrix remains English.");
        profiles = profiles.Where(profile => selected is null || selected == profile.Profile + "-" + profile.Stack).ToArray();
        if (profiles.Length == 0) throw new ArgumentException("Unknown semantic report profile.");
        var reports = new List<object>();
        foreach (HtmlBrowserEngine engine in DataTablesInterop.Engines(args)) foreach (var profile in profiles) {
            string output = Path.Combine(evidence, engine.ToString().ToLowerInvariant(), profile.Profile + "-" + profile.Stack + (language == "en" ? "" : "-" + language));
            Directory.CreateDirectory(output);
            await using var session = await HtmlBrowser.OpenSessionAsync(new Uri(modules, "../").AbsoluteUri,
                new HtmlBrowserLaunchOptions { Browser = engine, Headless = true, Timeout = 120000, Timezone = "Europe/Warsaw" });
            var errors = new List<string>(); session.Page.PageError += (_, error) => errors.Add(error);
            session.Page.RequestFailed += (_, request) => errors.Add("Request failed: " + request.Url + " " + request.Failure);
            var files = new List<object>();
            await session.Page.ExposeFunctionAsync("writeSemanticArtifact", (string name, string base64, string contract) => {
                if (Path.GetFileName(name) != name || !name.EndsWith(".pdf", StringComparison.Ordinal)) throw new InvalidDataException("Invalid semantic fixture name.");
                string path = Path.Combine(output, name);
                File.WriteAllBytes(path, Convert.FromBase64String(base64)); File.WriteAllText(path + ".json", contract);
                files.Add(new { name, bytes = new FileInfo(path).Length, sha256 = Hash(path) });
                return true;
            });
            if (profile.Profile == "datatables") {
                foreach (JsonElement dependency in comparison.RootElement.GetProperty("stacks").GetProperty(profile.Stack).GetProperty("scripts").EnumerateArray()) {
                    string path = Path.Combine(comparisonAssets, dependency.GetString()!); Capture(path);
                    await session.Page.AddScriptTagAsync(new() { Path = path });
                }
                await session.Page.AddScriptTagAsync(new() { Path = bundle });
            } else if (profile.Profile == "canopyx") {
                await session.Page.AddStyleTagAsync(new() { Content = canopyInputs["canopyx-reporting.css"] + canopyInputs["canopyx-filter-editor.css"] });
                await session.Page.AddScriptTagAsync(new() { Content = canopyInputs["canopyx-reporting.js"] + "\n" + canopyInputs["canopyx-filter-editor.js"] + "\n" + canopyInputs["canopyx-export.js"] });
            }
            await session.Page.AddScriptTagAsync(new() { Path = fixture });
            JsonElement result;
            try {
                result = await session.Page.EvaluateAsync<JsonElement>("input => runSemanticReportContracts(input)",
                    new { moduleBase = modules.AbsoluteUri.TrimEnd('/'), profile = profile.Profile, stack = profile.Stack, regular, bold, language });
            } catch (Exception error) {
                File.WriteAllText(Path.Combine(output, "failure.json"), JsonSerializer.Serialize(new { error = error.Message, errors }, new JsonSerializerOptions { WriteIndented = true })); throw;
            }
            if (errors.Count != 0) throw new InvalidDataException("Semantic browser errors: " + string.Join("; ", errors));
            if (files.Count != 1) throw new InvalidDataException("Semantic source did not deliver exactly one report.");
            var report = new { engine = engine.ToString(), browser = session.Browser?.Version, profile.Profile, profile.Stack, language, result, files, errors };
            File.WriteAllText(Path.Combine(output, "report.json"), JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true })); reports.Add(report);
            Console.WriteLine($"Semantic producer passed {engine}/{profile.Profile}/{profile.Stack}; independent PDF readback remains separate.");
        }
        if (hashes.Any(pair => Hash(pair.Key) != pair.Value)) throw new InvalidDataException("Semantic report source/input graph changed during qualification.");
        File.WriteAllText(Path.Combine(evidence, "semantic-report-interop.json"), JsonSerializer.Serialize(new {
            sources = hashes.ToDictionary(pair => Path.GetRelativePath(repository, pair.Key), pair => pair.Value), reports
        }, new JsonSerializerOptions { WriteIndented = true }));
    }
    private static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();
}
