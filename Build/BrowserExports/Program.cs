using System.Text.Json;
using HtmlTinkerX;
using OfficeIMO.Browser;
using OfficeIMO.TestAssets;

if (args.Length == 2 && args[0] == "--validate") {
    WorkbookVerifier.Verify(Path.GetFullPath(args[1]));
    Console.WriteLine("Workbook passed OfficeIMO loading and Open XML validation.");
    return;
}
if (args.Length < 2) throw new ArgumentException("Usage: <repository> <evidence-directory> [--scale] [--limits] [--example=<directory>]");
string repository = Path.GetFullPath(args[0]), evidence = Path.GetFullPath(args[1]);
Directory.CreateDirectory(evidence);
string vectorJson = File.ReadAllText(Path.Combine(repository, "OfficeIMO.TestAssets", "CSV", "browser-exports.json"));
string scenarios = File.ReadAllText(Path.Combine(repository, "Build", "BrowserExports", "scenarios.js"));
string host = Path.Combine(evidence, "fixture-host.html");
File.WriteAllText(host, "<!doctype html><meta charset=\"utf-8\"><title>Browser export fixture host</title>");
var reports = new List<object>();
foreach (HtmlBrowserEngine engine in new[] { HtmlBrowserEngine.Chromium, HtmlBrowserEngine.Firefox, HtmlBrowserEngine.WebKit }) {
    string output = Path.Combine(evidence, engine.ToString().ToLowerInvariant());
    Directory.CreateDirectory(output);
    Console.WriteLine($"Starting {engine} through HtmlTinkerX.");
    var launch = new HtmlBrowserLaunchOptions {
        Browser = engine, Headless = true, Timezone = "Europe/Warsaw", Timeout = 120000, LoadState = HtmlBrowserLoadState.Load
    };
    if (engine == HtmlBrowserEngine.Chromium && args.Contains("--scale")) launch.BrowserArguments.Add("--enable-precise-memory-info");
    await using HtmlBrowserSession session = await HtmlBrowser.OpenSessionAsync(new Uri(host).AbsoluteUri, launch);
    var errors = new List<string>();
    var produced = new List<string>();
    session.Page.PageError += (_, error) => errors.Add(error);
    await session.Page.ExposeFunctionAsync("writeFixture", (string fileName, string base64) => {
        if (Path.GetFileName(fileName) != fileName) throw new InvalidDataException("Invalid fixture file name.");
        File.WriteAllBytes(Path.Combine(output, fileName), Convert.FromBase64String(base64));
        produced.Add(fileName);
        return true;
    });
    await session.Page.AddScriptTagAsync(new() { Content = BrowserAssets.Script.Content });
    await session.Page.AddScriptTagAsync(new() { Content = scenarios });
    JsonElement result = await session.Page.EvaluateAsync<JsonElement>("async args => runBrowserScenarios(args)",
        new { vectorJson, workerScript = BrowserAssets.Script.Content, limits = args.Contains("--limits") });
    foreach (string name in produced.Where(name => name.EndsWith(".xlsx", StringComparison.Ordinal))) WorkbookVerifier.Verify(Path.Combine(output, name));
    using JsonDocument vectors = JsonDocument.Parse(vectorJson);
    foreach (JsonElement vector in vectors.RootElement.GetProperty("cases").EnumerateArray()) {
        byte[] expected = BrowserCsvVectorContract.Write(vector);
        if (!expected.SequenceEqual(BrowserCsvVectorContract.Expected(vector))) throw new InvalidDataException(".NET CSV vector mismatch.");
        string name = vector.GetProperty("name").GetString()!;
        if (!expected.SequenceEqual(File.ReadAllBytes(Path.Combine(output, name + ".csv")))) throw new InvalidDataException("Browser CSV byte mismatch: " + name);
    }
    if (errors.Count != 0) throw new InvalidDataException("Browser errors: " + string.Join("; ", errors));
    var report = new { engine = engine.ToString(), browserVersion = session.Browser?.Version, result,
        workbooks = produced.Count(name => name.EndsWith(".xlsx", StringComparison.Ordinal)), csvVectors = vectors.RootElement.GetProperty("cases").GetArrayLength(), errors };
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
