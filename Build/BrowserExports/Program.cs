using System.Text.Json;
using HtmlTinkerX;
using OfficeIMO.JavaScript.Assets;
using OfficeIMO.TestAssets;

if (args.Length > 0 && args[0] == "--validate-table-engine") {
    Console.WriteLine(JsonSerializer.Serialize(TableEngineComparisonVerifier.Verify(args)));
    return;
}
if (args.Length == 2 && args[0] == "--validate-wide") {
    string[] files = Directory.GetFiles(Path.GetFullPath(args[1]), "*.pdf", SearchOption.AllDirectories);
    if (files.Length == 0) throw new InvalidDataException("No wide report fixtures to validate.");
    foreach (string file in files) Console.WriteLine(JsonSerializer.Serialize(WideReportVerifier.Verify(file)));
    return;
}
if (args.Length == 2 && args[0] == "--validate-analytical") {
    string[] files = Directory.GetFiles(Path.GetFullPath(args[1]), "*", SearchOption.AllDirectories)
        .Where(file => file.EndsWith(".xlsx", StringComparison.Ordinal) || file.EndsWith(".pdf", StringComparison.Ordinal)).ToArray();
    if (files.Length == 0) throw new InvalidDataException("No analytical report fixtures to validate.");
    foreach (string file in files) Console.WriteLine(JsonSerializer.Serialize(HierarchicalReportVerifier.Verify(file)));
    return;
}
if (args.Length == 2 && args[0] == "--validate-grouped") {
    string[] files = Directory.GetFiles(Path.GetFullPath(args[1]), "*", SearchOption.AllDirectories)
        .Where(file => file.EndsWith(".xlsx", StringComparison.Ordinal) || file.EndsWith(".pdf", StringComparison.Ordinal)).ToArray();
    if (files.Length == 0) throw new InvalidDataException("No grouped report fixtures to validate.");
    foreach (string file in files)
        Console.WriteLine(JsonSerializer.Serialize(GroupedReportVerifier.Verify(file)));
    return;
}
if (args.Length == 2 && args[0] == "--validate-editable-analysis") {
    string[] files = Directory.GetFiles(Path.GetFullPath(args[1]), "*.xlsx", SearchOption.AllDirectories);
    if (files.Length == 0) throw new InvalidDataException("No editable analysis fixtures to validate.");
    foreach (string file in files)
        Console.WriteLine(JsonSerializer.Serialize(EditableAnalysisVerifier.Verify(file)));
    return;
}
if (args.Length == 3 && args[0] == "--validate-editable-native") {
    Console.WriteLine(JsonSerializer.Serialize(EditableAnalysisVerifier.VerifyNativeEdit(Path.GetFullPath(args[1]), Path.GetFullPath(args[2]))));
    return;
}
if (args.Length == 2 && args[0] == "--validate-links") {
    var files = Directory.GetFiles(Path.GetFullPath(args[1])).Where(f => f.EndsWith(".xlsx", StringComparison.Ordinal) || f.EndsWith(".pdf", StringComparison.Ordinal)).ToArray();
    if (files.Length == 0) throw new InvalidDataException("No portable-link fixtures to validate.");
    foreach (string file in files) Console.WriteLine(JsonSerializer.Serialize(PortableLinkVerifier.Verify(file)));
    return;
}
if (args.Length == 2 && args[0] == "--validate-professional") {
    string[] files = Directory.GetFiles(Path.GetFullPath(args[1]), "*.pdf", SearchOption.AllDirectories);
    if (files.Length == 0) throw new InvalidDataException("No professional PDF fixtures to validate.");
    foreach (string file in files) Console.WriteLine(JsonSerializer.Serialize(ProfessionalPdfVerifier.Verify(file)));
    return;
}
if (args.Length == 2 && args[0] == "--validate") {
    WorkbookVerifier.Verify(Path.GetFullPath(args[1]));
    Console.WriteLine("Workbook passed OfficeIMO loading and Open XML validation.");
    return;
}
if (args.Length == 2 && args[0] == "--validate-directory") {
    string[] files = Directory.GetFiles(Path.GetFullPath(args[1]), "*.xlsx");
    if (files.Length == 0) throw new InvalidDataException("No XLSX files to validate.");
    foreach (string file in files) WorkbookVerifier.Verify(file);
    Console.WriteLine($"Validated {files.Length} XLSX files with both OfficeIMO readers and the Open XML SDK.");
    return;
}
if (args.Length < 2) throw new ArgumentException("Usage: <repository> <evidence-directory> [--scale] [--limits] [--example=<directory>]");
string repository = Path.GetFullPath(args[0]), evidence = Path.GetFullPath(args[1]);
Directory.CreateDirectory(evidence);
string? exampleOnly = args.FirstOrDefault(argument => argument.StartsWith("--example-only=", StringComparison.Ordinal));
if (exampleOnly is not null) {
    string exampleDirectory = Path.GetFullPath(exampleOnly["--example-only=".Length..]);
    var exampleReports = new List<object>();
    string exampleHost = Path.Combine(evidence, "example-host.html");
    File.WriteAllText(exampleHost, "<!doctype html><meta charset=utf-8><title>Offline example qualification</title>");
    foreach (HtmlBrowserEngine engine in DataTablesInterop.Engines(args)) {
        await using var session = await HtmlBrowser.OpenSessionAsync(new Uri(exampleHost).AbsoluteUri,
            new HtmlBrowserLaunchOptions { Browser = engine, Headless = true, Timeout = 120000, Timezone = "Europe/Warsaw" });
        foreach (bool translated in args.Contains("--example-translated") ? new[] { false, true } : new[] { false }) {
            string output = Path.Combine(evidence, engine.ToString().ToLowerInvariant(), translated ? "translated" : "default");
            Directory.CreateDirectory(output);
            await ExampleVerifier.RunAsync(session, exampleDirectory, output, translated);
            using var result = JsonDocument.Parse(File.ReadAllText(Path.Combine(output, "example-report.json")));
            exampleReports.Add(new { engine = engine.ToString(), browser = session.Browser?.Version, result = result.RootElement.Clone() });
        }
    }
    File.WriteAllText(Path.Combine(evidence, "example-only-report.json"), JsonSerializer.Serialize(new {
        sources = Directory.GetFiles(exampleDirectory).Where(file => file.EndsWith(".html", StringComparison.Ordinal) || file.EndsWith(".js", StringComparison.Ordinal))
            .ToDictionary(file => Path.GetFileName(file)!, file => Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(file))).ToLowerInvariant()),
        reports = exampleReports
    }, new JsonSerializerOptions { WriteIndented = true }));
    return;
}
if (args.Contains("--execution")) { await ExecutionInterop.RunAsync(repository, evidence, args); return; }
if (args.Any(value => value.StartsWith("--hfx-operational=", StringComparison.Ordinal))) { await HfxOperationalInterop.RunAsync(evidence, args); return; }
if (args.Contains("--shaping")) { await ShapingInterop.RunAsync(repository, evidence, args); return; }
if (args.Contains("--editable-analysis")) { await EditableAnalysisInterop.RunAsync(repository, evidence, args); return; }
if (args.Contains("--grouped-report")) { await GroupedReportInterop.RunAsync(repository, evidence, args); return; }
if (args.Contains("--analytical-report")) { await HierarchicalReportInterop.RunAsync(repository, evidence, args); return; }
if (args.Contains("--wide-report")) { await WideReportInterop.RunAsync(repository, evidence, args); return; }
if (args.Contains("--semantic-report")) { await SemanticReportInterop.RunAsync(repository, evidence, args); return; }
if (args.Contains("--pdf-professional")) { await ProfessionalPdfInterop.RunAsync(repository, evidence, args); return; }
if (args.Contains("--report-engine")) { await ReportEngineInterop.RunAsync(repository, evidence, args); return; }
if (args.Any(a => a.StartsWith("--canopy=", StringComparison.Ordinal))) { await CanopyInterop.RunAsync(repository, evidence, args); return; }
if (args.Contains("--links")) { await PortableLinkInterop.RunAsync(repository, evidence, args); return; }
if (args.Contains("--pdf")) { await PdfInterop.RunAsync(repository, evidence, args); return; }
if (args.Contains("--pdf-rich")) { await PdfRichInterop.RunAsync(repository, evidence, args); return; }
if (args.Contains("--datatables-session")) { await DataTablesComparisonSession.RunAsync(repository, evidence, args); return; }
if (args.Contains("--datatables")) { await DataTablesInterop.RunAsync(repository, evidence, args); return; }
if (args.Contains("--qualify")) { await ExportQualification.RunAsync(repository, evidence, args); return; }
string vectorJson = File.ReadAllText(Path.Combine(repository, "OfficeIMO.TestAssets", "CSV", "browser-exports.json"));
string scenarios = File.ReadAllText(Path.Combine(repository, "Build", "BrowserExports", "scenarios.js"));
string layers = File.ReadAllText(Path.Combine(repository, "Build", "BrowserExports", "layers.js"));
string reportContracts = File.ReadAllText(Path.Combine(repository, "Build", "BrowserExports", "reports.js"));
string fixtureJson = File.ReadAllText(Path.Combine(repository, "OfficeIMO.TestAssets", "JavaScript", "xlsx-writer.json"));
string? moduleUrl = args.FirstOrDefault(a => a.StartsWith("--module-url=", StringComparison.Ordinal))?.Substring("--module-url=".Length);
string host = Path.Combine(evidence, "fixture-host.html");
File.WriteAllText(Path.Combine(evidence, JavaScriptAssets.Script.HashedFileName), JavaScriptAssets.Script.Content, new System.Text.UTF8Encoding(false));
File.WriteAllText(host, "<!doctype html><meta charset=\"utf-8\"><title>Browser export fixture host</title><script src=\"" + JavaScriptAssets.Script.HashedFileName + "\"></script>");
var reports = new List<object>();
var engines = new[] { HtmlBrowserEngine.Chromium, HtmlBrowserEngine.Firefox, HtmlBrowserEngine.WebKit };
string? selectedEngine = args.FirstOrDefault(a => a.StartsWith("--engine=", StringComparison.Ordinal))?.Substring("--engine=".Length);
if (selectedEngine is not null) {
    engines = engines.Where(engine => string.Equals(engine.ToString(), selectedEngine, StringComparison.OrdinalIgnoreCase)).ToArray();
    if (engines.Length == 0) throw new ArgumentException("Unknown browser engine: " + selectedEngine);
}
foreach (HtmlBrowserEngine engine in engines) {
    string output = Path.Combine(evidence, engine.ToString().ToLowerInvariant());
    Directory.CreateDirectory(output);
    Console.WriteLine($"Starting {engine} through HtmlTinkerX.");
    var launch = new HtmlBrowserLaunchOptions {
        Browser = engine, Headless = true, Timezone = "Europe/Warsaw", Timeout = 120000, LoadState = HtmlBrowserLoadState.Load
    };
    if (engine == HtmlBrowserEngine.Chromium && args.Contains("--scale")) launch.BrowserArguments.Add("--enable-precise-memory-info");
    await using HtmlBrowserSession session = await HtmlBrowser.OpenSessionAsync(new Uri(host).AbsoluteUri, launch);
    Console.WriteLine($"{engine}: fixture page ready.");
    session.Page.Console += (_, message) => {
        if (message.Text.StartsWith("OfficeIMO report:", StringComparison.Ordinal)) Console.WriteLine($"{engine}: {message.Text}");
    };
    var errors = new List<string>();
    var produced = new List<string>();
    session.Page.PageError += (_, error) => errors.Add(error);
    await session.Page.ExposeFunctionAsync("writeFixture", (string fileName, string base64) => {
        if (Path.GetFileName(fileName) != fileName) throw new InvalidDataException("Invalid fixture file name.");
        File.WriteAllBytes(Path.Combine(output, fileName), Convert.FromBase64String(base64));
        produced.Add(fileName);
        return true;
    });
    await session.Page.AddScriptTagAsync(new() { Content = scenarios });
    Console.WriteLine($"{engine}: workbook and CSV scenarios.");
    JsonElement result = await session.Page.EvaluateAsync<JsonElement>("async args => runBrowserScenarios(args)",
        new { vectorJson, workerScript = JavaScriptAssets.Script.Content, limits = args.Contains("--limits") }).WaitAsync(TimeSpan.FromMinutes(3));
    await session.Page.AddScriptTagAsync(new() { Content = layers });
    Console.WriteLine($"{engine}: module and corpus scenarios.");
    JsonElement layerResult = await session.Page.EvaluateAsync<JsonElement>("args => runLayerScenarios(args)", new { fixtureJson, moduleBase = moduleUrl }).WaitAsync(TimeSpan.FromMinutes(3));
    await session.Page.AddScriptTagAsync(new() { Content = reportContracts });
    Console.WriteLine($"{engine}: report scenarios.");
    JsonElement reportResult = await session.Page.EvaluateAsync<JsonElement>("fixture => runReportContracts(fixture)", fixtureJson).WaitAsync(TimeSpan.FromMinutes(3));
    foreach (string name in produced.Where(name => name.EndsWith(".xlsx", StringComparison.Ordinal))) WorkbookVerifier.Verify(Path.Combine(output, name));
    foreach (string name in produced.Where(name => name.StartsWith("report-", StringComparison.Ordinal))) ReportVerifier.Verify(Path.Combine(output, name), fixtureJson);
    using (JsonDocument corpus = JsonDocument.Parse(fixtureJson)) {
        foreach (JsonElement spec in corpus.RootElement.GetProperty("cases").EnumerateArray())
            foreach (string kind in moduleUrl is null ? new[] { "classic" } : new[] { "classic", "esm" })
                foreach (string compression in new[] { "auto", "store" })
                    JavaScriptWorkbookContract.Verify(Path.Combine(output, "corpus-" + kind + "-" + spec.GetProperty("name").GetString() + "-" + compression + ".xlsx"), spec);
    }
    using JsonDocument vectors = JsonDocument.Parse(vectorJson);
    foreach (JsonElement vector in vectors.RootElement.GetProperty("cases").EnumerateArray()) {
        byte[] expected = BrowserCsvVectorContract.Write(vector);
        if (!expected.SequenceEqual(BrowserCsvVectorContract.Expected(vector))) throw new InvalidDataException(".NET CSV vector mismatch.");
        string name = vector.GetProperty("name").GetString()!;
        if (!expected.SequenceEqual(File.ReadAllBytes(Path.Combine(output, name + ".csv")))) throw new InvalidDataException("Browser CSV byte mismatch: " + name);
    }
    if (errors.Count != 0) throw new InvalidDataException("Browser errors: " + string.Join("; ", errors));
    var report = new { engine = engine.ToString(), browserVersion = session.Browser?.Version, result,
        layerResult, reportResult, workbooks = produced.Count(name => name.EndsWith(".xlsx", StringComparison.Ordinal)), csvVectors = vectors.RootElement.GetProperty("cases").GetArrayLength(), errors };
    reports.Add(report);
    File.WriteAllText(Path.Combine(output, "report.json"), JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
    Console.WriteLine($"Passed {engine}: {report.workbooks} workbooks, {report.csvVectors} byte-identical CSV vectors.");
    string? example = args.FirstOrDefault(a => a.StartsWith("--example=", StringComparison.Ordinal));
    if (example is not null) await ExampleVerifier.RunAsync(session, Path.GetFullPath(example.Substring("--example=".Length)), output);
    if (args.Contains("--scale")) {
        await session.Page.AddScriptTagAsync(new() { Content = scenarios });
        foreach (string format in new[] { "xlsx", "csv" }) {
            var download = await session.Page.RunAndWaitForDownloadAsync(() => session.Page.EvaluateAsync("format => runScale(format)", format),
                new() { Timeout = 600000 });
            string path = Path.Combine(output, "scale." + format);
            await download.SaveAsAsync(path);
            JsonElement metrics = await session.Page.EvaluateAsync<JsonElement>("globalThis.scaleMetrics");
            File.WriteAllText(Path.Combine(output, "scale-" + format + ".json"), metrics.GetRawText());
            if (format == "xlsx") WorkbookVerifier.VerifyScale(path);
            Console.WriteLine($"Scale {engine}/{format}: {metrics.GetRawText()}");
        }
    }
    if (errors.Count != 0) throw new InvalidDataException("Browser errors: " + string.Join("; ", errors));
}
File.WriteAllText(Path.Combine(evidence, "browser-interop.json"), JsonSerializer.Serialize(reports, new JsonSerializerOptions { WriteIndented = true }));
