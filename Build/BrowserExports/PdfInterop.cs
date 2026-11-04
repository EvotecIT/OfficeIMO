using System.Text.Json;
using HtmlTinkerX;
using OfficeIMO.JavaScript.Assets;

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
        const int ProfileTimeoutMilliseconds = 120000;
        bool scale = args.Contains("--scale");
        var reports = new List<object>();
        foreach (HtmlBrowserEngine engine in DataTablesInterop.Engines(args)) {
            string output = Path.Combine(evidence, engine.ToString().ToLowerInvariant()); Directory.CreateDirectory(output);
            const int MaxProgressEvents = 256;
            var progressGate = new object();
            var progress = new Queue<object>();
            int droppedEvents = 0;
            string status = "running";
            void Record(string phase, string? detail = null) {
                lock (progressGate) {
                    if (progress.Count == MaxProgressEvents) { progress.Dequeue(); droppedEvents++; }
                    progress.Enqueue(new { utc = DateTimeOffset.UtcNow, phase, detail = detail is { Length: > 2048 } ? detail[..2048] : detail });
                }
                if (phase != "browser-console") {
                    try { Console.WriteLine($"PDF progress {stack}/{engine}: {phase} {detail}"); }
                    catch { /* Diagnostics must not replace the qualification failure. */ }
                }
            }
            void SaveProgress() {
                try {
                    object snapshot;
                    lock (progressGate) snapshot = new { engine = engine.ToString(), stack, status, droppedEvents, events = progress.ToArray() };
                    File.WriteAllText(Path.Combine(output, "progress.json"), JsonSerializer.Serialize(snapshot, new JsonSerializerOptions { WriteIndented = true }));
                } catch (Exception diagnosticError) {
                    try { Console.Error.WriteLine($"Could not write PDF progress {stack}/{engine}: {diagnosticError.Message}"); }
                    catch { /* A diagnostic write failure must not mask the original exception. */ }
                }
            }
            Record("profile-start");
            try {
                string host = Path.Combine(output, "host.html"); File.WriteAllText(host, "<!doctype html><meta charset=utf-8><title>PDF export qualification</title>");
                await using HtmlBrowserSession session = assets is not null
                    ? await DataTablesInterop.OpenAsync(assets, manifest.RootElement, stack, engine, output)
                    : await HtmlBrowser.OpenSessionAsync(new Uri(host).AbsoluteUri, new HtmlBrowserLaunchOptions { Browser = engine, Headless = true, Timeout = ProfileTimeoutMilliseconds });
                try {
                    Record("browser-ready", session.Browser?.Version);
                    var errors = new List<string>();
                    session.Page.PageError += (_, error) => { errors.Add(error); Record("page-error", error); };
                    session.Page.Console += (_, message) => Record("browser-console", $"{message.Type}: {message.Text}");
                    var files = new List<object>(); long removedBytes = 0;
                    await session.Page.AddScriptTagAsync(new() { Content = JavaScriptAssets.Script.Content });
                    await session.Page.ExposeFunctionAsync("writePdfFixture", (string name, string base64, string contract) => {
                        Record("callback-received", name);
                        if (Path.GetFileName(name) != name || !name.EndsWith(".pdf", StringComparison.Ordinal)) throw new InvalidDataException("Invalid PDF fixture name.");
                        File.WriteAllBytes(Path.Combine(output, name), Convert.FromBase64String(base64));
                        File.WriteAllText(Path.Combine(output, name + ".json"), contract);
                        Record("callback-saved", name);
                        Record("verify-begin", name);
                        files.Add(PdfExportVerifier.Verify(Path.Combine(output, name)));
                        Record("verified", name);
                        if (name.StartsWith("scale-", StringComparison.Ordinal)) {
                            removedBytes += new FileInfo(Path.Combine(output, name)).Length + new FileInfo(Path.Combine(output, name + ".json")).Length;
                            File.Delete(Path.Combine(output, name)); File.Delete(Path.Combine(output, name + ".json"));
                        }
                        Record("callback-return", name);
                        return true;
                    });
                    await session.Page.AddScriptTagAsync(new() { Content = script });
                    Record("evaluation-begin");
                    Task<JsonElement> evaluation = session.Page.EvaluateAsync<JsonElement>("args => runPdfContracts(args)",
                        new { regular, bold, symbols, japanese, workerScript = JavaScriptAssets.Script.Content, scale });
                    // Opt-in scale measurements retain their duration; ordinary correctness contracts are bounded.
                    JsonElement result = scale ? await evaluation
                        : await evaluation.WaitAsync(TimeSpan.FromMilliseconds(ProfileTimeoutMilliseconds));
                    Record("evaluation-complete");
                    if (errors.Count != 0) throw new InvalidDataException(string.Join("; ", errors));
                    var report = new { engine = engine.ToString(), browser = session.Browser?.Version, stack, asset = JavaScriptAssets.Script.HashedFileName, result, files, removedBytes, errors };
                    reports.Add(report);
                    File.WriteAllText(Path.Combine(output, "report.json"), JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
                    Console.WriteLine($"PDF passed {engine}: {files.Count} independent-reader-validated documents.");
                    status = "passed";
                    Record("profile-passed");
                } catch (Exception error) {
                    status = "failed";
                    Record("profile-failed", error.ToString());
                    throw;
                } finally {
                    Record("browser-dispose-begin");
                    SaveProgress();
                }
            } catch (Exception error) {
                if (status != "failed") { status = "failed"; Record("profile-failed", error.ToString()); }
                SaveProgress();
                throw;
            }
            Record("browser-disposed");
            SaveProgress();
        }
        File.WriteAllText(Path.Combine(evidence, "pdf-interop.json"), JsonSerializer.Serialize(reports, new JsonSerializerOptions { WriteIndented = true }));
    }
}
