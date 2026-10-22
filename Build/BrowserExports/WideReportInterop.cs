using System.Security.Cryptography;
using System.Text.Json;
using HtmlTinkerX;

/// <summary>Opt-in wide PDF proof over real DataTables and CanopyX captures, with a representative operational inventory.</summary>
internal static class WideReportInterop {
    internal static async Task RunAsync(string repository, string evidence, string[] args) {
        string Option(string name) => args.FirstOrDefault(value => value.StartsWith(name + "=", StringComparison.Ordinal))?[(name.Length + 1)..]
            ?? throw new ArgumentException("Wide report qualification requires " + name + ".");
        var modules = new Uri(Option("--module-url").TrimEnd('/') + "/");
        if (!modules.IsLoopback || modules.Scheme != "http") throw new ArgumentException("Wide report modules require a loopback HTTP server.");
        string assets = Path.GetFullPath(Option("--comparison-assets")), canopy = Path.GetFullPath(Option("--canopy-assets"));
        using var comparison = JsonDocument.Parse(File.ReadAllText(Path.Combine(repository, "Build", "BrowserExports", "comparison-assets.json")));
        DataTablesInterop.VerifyAssets(assets, comparison.RootElement);
        string fixture = Path.Combine(repository, "Build", "BrowserExports", "wide-report-contracts.js");
        string bundle = Path.Combine(repository, "OfficeIMO.JavaScript", "bundles", "officeimo.js");
        string moduleRoot = Path.Combine(repository, "OfficeIMO.JavaScript", "dist");
        var hashes = new Dictionary<string, string>();
        void Capture(string path) => hashes.Add(path, Hash(path));
        Capture(fixture); Capture(bundle);
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
        (string Profile, string Stack)[] profiles = [("hfx", "worker"), ("datatables", "bundled"), ("datatables", "current"), ("canopyx", "offset"), ("canopyx", "cursor")];
        string? selected = args.FirstOrDefault(value => value.StartsWith("--wide-profile=", StringComparison.Ordinal))?[15..];
        profiles = profiles.Where(profile => selected is null || selected == profile.Profile + "-" + profile.Stack).ToArray();
        if (profiles.Length == 0) throw new ArgumentException("Unknown wide report profile.");
        var reports = new List<object>();
        foreach (HtmlBrowserEngine engine in DataTablesInterop.Engines(args)) foreach (var profile in profiles) {
            string output = Path.Combine(evidence, engine.ToString().ToLowerInvariant(), profile.Profile + "-" + profile.Stack); Directory.CreateDirectory(output);
            await using var session = await HtmlBrowser.OpenSessionAsync(new Uri(modules, "../").AbsoluteUri,
                new HtmlBrowserLaunchOptions { Browser = engine, Headless = true, Timeout = 120000, Timezone = "Europe/Warsaw" });
            var errors = new List<string>(); session.Page.PageError += (_, error) => errors.Add(error);
            session.Page.RequestFailed += (_, request) => errors.Add("Request failed: " + request.Url + " " + request.Failure);
            var files = new List<object>();
            await session.Page.ExposeFunctionAsync("writeWideArtifact", (string name, string base64, string contract) => {
                if (Path.GetFileName(name) != name || !name.EndsWith(".pdf", StringComparison.Ordinal)) throw new InvalidDataException("Invalid wide PDF fixture name.");
                string path = Path.Combine(output, name); File.WriteAllBytes(path, Convert.FromBase64String(base64)); File.WriteAllText(path + ".json", contract);
                files.Add(WideReportVerifier.Verify(path)); return true;
            });
            if (profile.Profile == "datatables") {
                foreach (JsonElement dependency in comparison.RootElement.GetProperty("stacks").GetProperty(profile.Stack).GetProperty("scripts").EnumerateArray())
                    await session.Page.AddScriptTagAsync(new() { Path = Path.Combine(assets, dependency.GetString()!) });
                await session.Page.AddScriptTagAsync(new() { Path = bundle });
            } else if (profile.Profile == "canopyx") {
                await session.Page.AddStyleTagAsync(new() { Content = canopyInputs["canopyx-reporting.css"] + canopyInputs["canopyx-filter-editor.css"] });
                await session.Page.AddScriptTagAsync(new() { Content = canopyInputs["canopyx-reporting.js"] + "\n" + canopyInputs["canopyx-filter-editor.js"] + "\n" + canopyInputs["canopyx-export.js"] });
            }
            await session.Page.AddScriptTagAsync(new() { Path = fixture });
            JsonElement result;
            try { result = await session.Page.EvaluateAsync<JsonElement>("input => runWideReportContracts(input)", new {
                moduleBase = modules.AbsoluteUri.TrimEnd('/'), profile = profile.Profile, stack = profile.Stack, regular, bold
            }); } catch (Exception error) {
                File.WriteAllText(Path.Combine(output, "failure.json"), JsonSerializer.Serialize(new { error = error.Message, errors }, new JsonSerializerOptions { WriteIndented = true })); throw;
            }
            if (errors.Count != 0) throw new InvalidDataException("Wide report browser errors: " + string.Join("; ", errors));
            if (files.Count != 1) throw new InvalidDataException("Wide source did not produce its required delivery artifact.");
            var report = new { engine = engine.ToString(), browser = session.Browser?.Version, profile.Profile, profile.Stack, result, files, errors };
            File.WriteAllText(Path.Combine(output, "report.json"), JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true })); reports.Add(report);
            Console.WriteLine($"Wide reports passed {engine}/{profile.Profile}/{profile.Stack}: selected values, aligned bands and independent text geometry.");
        }
        if (hashes.Any(pair => Hash(pair.Key) != pair.Value)) throw new InvalidDataException("Wide report source/input graph changed during qualification.");
        File.WriteAllText(Path.Combine(evidence, "wide-report-interop.json"), JsonSerializer.Serialize(new {
            sources = hashes.ToDictionary(pair => Path.GetRelativePath(repository, pair.Key), pair => pair.Value), reports
        }, new JsonSerializerOptions { WriteIndented = true }));
    }
    private static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();
}
