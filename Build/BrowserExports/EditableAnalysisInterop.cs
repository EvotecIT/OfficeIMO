using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using HtmlTinkerX;

/// <summary>Qualify one editable workbook through real grid captures and offline classic assets, using HtmlTinkerX sessions.</summary>
internal static class EditableAnalysisInterop {
    internal static async Task RunAsync(string repository, string evidence, string[] args) {
        string Option(string name) => args.FirstOrDefault(argument => argument.StartsWith(name + "=", StringComparison.Ordinal))?[(name.Length + 1)..]
            ?? throw new ArgumentException("Editable analysis qualification requires " + name + ".");
        var modules = new Uri(Option("--module-url").TrimEnd('/') + "/");
        var example = new Uri(Option("--example-url"));
        if (!modules.IsLoopback || modules.Scheme != "http" || !example.IsLoopback || example.Scheme != "http" || example.Authority != modules.Authority)
            throw new ArgumentException("Current modules and the executable example require the same loopback HTTP server.");
        string assets = Path.GetFullPath(Option("--comparison-assets"));
        string canopy = Path.Combine(Path.GetFullPath(Option("--canopy")), "CanopyX.Html", "Assets");
        using var comparison = JsonDocument.Parse(File.ReadAllText(Path.Combine(repository, "Build", "BrowserExports", "comparison-assets.json")));
        DataTablesInterop.VerifyAssets(assets, comparison.RootElement);
        string scriptPath = Path.Combine(repository, "Build", "BrowserExports", "editable-analysis-contracts.js"), script = File.ReadAllText(scriptPath);
        string examplePath = Path.Combine(repository, "OfficeIMO.JavaScript.Examples", "report-engine-job.mjs"), exampleSource = File.ReadAllText(examplePath);
        string bundlePath = Path.Combine(repository, "OfficeIMO.JavaScript", "bundles", "officeimo.js"), bundle = File.ReadAllText(bundlePath);
        string executionPath = Path.Combine(repository, "OfficeIMO.JavaScript", "bundles", "officeimo-execution.js"), execution = File.ReadAllText(executionPath);
        string moduleRoot = Path.Combine(repository, "OfficeIMO.JavaScript", "dist");
        var hashes = new Dictionary<string, string> { [examplePath] = Hash(examplePath), [bundlePath] = Hash(bundlePath),
            [executionPath] = Hash(executionPath), [scriptPath] = Hash(scriptPath) };
        foreach (string file in Directory.EnumerateFiles(moduleRoot, "*.js", SearchOption.AllDirectories)) hashes.Add(file, Hash(file));
        var canopyInputs = new Dictionary<string, string>();
        foreach (string name in new[] { "canopyx-reporting.js", "canopyx-reporting.css", "canopyx-filter-editor.js", "canopyx-filter-editor.css", "canopyx-export.js" }) {
            string path = Path.Combine(canopy, name); canopyInputs.Add(name, File.ReadAllText(path)); hashes.Add(path, Hash(path));
        }
        string? selectedProfile = args.FirstOrDefault(argument => argument.StartsWith("--analysis-profile=", StringComparison.Ordinal))?[19..];
        string? selectedStack = args.FirstOrDefault(argument => argument.StartsWith("--stack=", StringComparison.Ordinal))?[8..];
        (string Profile, string Stack)[] profiles = [("datatables", "bundled"), ("datatables", "current"), ("canopyx", "native"), ("offline-currentview", "classic")];
        profiles = profiles.Where(item => (selectedProfile is null || selectedProfile == item.Profile) &&
            (selectedStack is null || item.Profile != "datatables" || item.Stack == selectedStack)).ToArray();
        if (profiles.Length == 0) throw new ArgumentException("Unknown editable-analysis profile or DataTables stack.");
        var reports = new List<object>();
        foreach (HtmlBrowserEngine engine in DataTablesInterop.Engines(args)) foreach (var profile in profiles) {
            string output = Path.Combine(evidence, engine.ToString().ToLowerInvariant(), profile.Profile + "-" + profile.Stack); Directory.CreateDirectory(output);
            bool offline = profile.Profile == "offline-currentview";
            string url = new Uri(modules, "../").AbsoluteUri;
            if (offline) {
                if (exampleSource.Contains("</script", StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("The inline executable example contains an HTML script terminator.");
                File.WriteAllText(Path.Combine(output, "officeimo.js"), bundle, new UTF8Encoding(false));
                File.WriteAllText(Path.Combine(output, "officeimo-execution.js"), execution, new UTF8Encoding(false));
                string host = Path.Combine(output, "current-view.html");
                File.WriteAllText(host, "<!doctype html><meta charset=utf-8><title>Offline editable analysis</title>" +
                    "<script src=officeimo.js></script><script src=officeimo-execution.js></script><script type=module>\n" + exampleSource +
                    "\nglobalThis.serviceAnalysisExample={reportColumns,operationalRows,writeServiceAnalysis};</script>", new UTF8Encoding(false));
                url = new Uri(host).AbsoluteUri;
            }
            await using var session = await HtmlBrowser.OpenSessionAsync(url,
                new HtmlBrowserLaunchOptions { Browser = engine, Headless = true, Timeout = 120000, Timezone = "Europe/Warsaw" });
            var errors = new List<string>(); session.Page.PageError += (_, error) => errors.Add(error);
            var files = new List<object>();
            await session.Page.ExposeFunctionAsync("writeEditableAnalysis", (string name, string base64, string contract) => {
                if (Path.GetFileName(name) != name || !name.EndsWith(".xlsx", StringComparison.Ordinal)) throw new InvalidDataException("Invalid editable analysis fixture name.");
                string path = Path.Combine(output, name); File.WriteAllBytes(path, Convert.FromBase64String(base64)); File.WriteAllText(path + ".json", contract);
                files.Add(EditableAnalysisVerifier.Verify(path)); return true;
            });
            if (profile.Profile == "datatables") {
                foreach (JsonElement dependency in comparison.RootElement.GetProperty("stacks").GetProperty(profile.Stack).GetProperty("scripts").EnumerateArray())
                    await session.Page.AddScriptTagAsync(new() { Path = Path.Combine(assets, dependency.GetString()!) });
                await session.Page.AddScriptTagAsync(new() { Content = bundle });
            } else if (profile.Profile == "canopyx") {
                await session.Page.AddStyleTagAsync(new() { Content = canopyInputs["canopyx-reporting.css"] + canopyInputs["canopyx-filter-editor.css"] });
                await session.Page.AddScriptTagAsync(new() { Content = canopyInputs["canopyx-reporting.js"] + "\n" + canopyInputs["canopyx-filter-editor.js"] + "\n" + canopyInputs["canopyx-export.js"] });
            } else await session.Page.WaitForFunctionAsync("typeof globalThis.serviceAnalysisExample?.writeServiceAnalysis === 'function'");
            await session.Page.AddScriptTagAsync(new() { Content = script });
            JsonElement result = await session.Page.EvaluateAsync<JsonElement>("input => runEditableAnalysisContracts(input)", new {
                moduleBase = modules.AbsoluteUri.TrimEnd('/'), exampleUrl = example.AbsoluteUri, profile = profile.Profile,
                classicScript = offline ? bundle : "", executionScript = offline ? execution : ""
            });
            if (errors.Count != 0) throw new InvalidDataException("Editable analysis browser errors: " + string.Join("; ", errors));
            if (files.Count != 4) throw new InvalidDataException("Blob, sink, worker and empty artifacts are required for each source profile.");
            var report = new { engine = engine.ToString(), browser = session.Browser?.Version, profile = profile.Profile, stack = profile.Stack,
                delivery = offline ? "local file with exact classic assets and inline executable example; no server" : "current source modules over loopback HTTP",
                result, files, errors, nativeEditAcceptance = "separate Excel edit, recalculate, save and reopen required" };
            File.WriteAllText(Path.Combine(output, "report.json"), JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
            reports.Add(report); Console.WriteLine($"Editable analysis passed {engine}/{profile.Profile}/{profile.Stack}: real selected source, Blob/sink/worker/empty, cancellation and independent workbook readback.");
        }
        if (hashes.Any(pair => Hash(pair.Key) != pair.Value)) throw new InvalidDataException("Editable analysis source or input graph changed during qualification.");
        File.WriteAllText(Path.Combine(evidence, "editable-analysis-interop.json"), JsonSerializer.Serialize(new {
            sources = hashes.ToDictionary(pair => Path.GetRelativePath(repository, pair.Key), pair => pair.Value), reports
        }, new JsonSerializerOptions { WriteIndented = true }));
    }

    private static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();
}
