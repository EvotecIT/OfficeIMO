using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using HtmlTinkerX;
using OfficeIMO.Browser;

/// <summary>Opt-in browser/output qualification, with bounded file delivery and per-artifact validation.</summary>
internal static class ExportQualification {
    internal sealed record Case(string Format, int Rows, int Columns, bool Styled = false, bool Unique = false,
        bool LongText = false, bool Worker = false, bool Fallback = false, bool SlowSink = false,
        int? CancelAfterRows = null, bool HangSink = false, bool ResourceLimit = false, bool DelayedPages = false, bool PendingPage = false, bool Conditional = false) {
        internal string Name => $"{Format}-{Rows}-{Columns}-{(Styled ? "styled" : "plain")}" +
            (Unique ? "-unique" : "-repeated") + (LongText ? "-unicode" : "") + (Worker ? "-worker" : "") +
            (Fallback ? "-fallback" : "") + (SlowSink ? "-slow" : "") + (CancelAfterRows.HasValue ? "-cancel" : "") +
            (HangSink ? "-hung" : "") + (ResourceLimit ? "-limit" : "") + (DelayedPages ? "-paged" : "") + (PendingPage ? "-page-cancel" : "") + (Conditional ? "-conditional" : "");
    }
    internal static async Task RunAsync(string repository, string evidence, string[] args) {
        int[] sizes = args.Contains("--full") ? new[] { 10000, 100000, 250000, 1000000 } : new[] { 10000 };
        string? rows = args.FirstOrDefault(a => a.StartsWith("--rows=", StringComparison.Ordinal));
        if (rows is not null) sizes = rows.Substring(7).Split(',').Select(int.Parse).ToArray();
        if (sizes.Any(n => n < 1000 || n > 1048570)) throw new ArgumentException("Qualification sizes must be between 1,000 and 1,048,570 rows.");
        var cases = new List<Case>();
        foreach (int size in sizes) foreach (int width in new[] { 4, 20 }) foreach (bool styled in new[] { false, true }) foreach (bool unique in new[] { false, true })
            foreach (string format in new[] { "xlsx", "csv" }) cases.Add(new(format, size, width, styled, Unique: unique, LongText: size == 10000));
        if (!args.Contains("--matrix-only")) foreach (string format in new[] { "xlsx", "csv" }) {
            cases.Add(new(format, 10000, 20, Styled: true, Unique: true, LongText: true, Worker: true));
            cases.Add(new(format, 10000, 4, LongText: true, Worker: true, Fallback: true));
            cases.Add(new(format, 10000, 20, Styled: true, SlowSink: true, Fallback: true));
            cases.Add(new(format, 10000, 4, CancelAfterRows: 500));
            cases.Add(new(format, 10000, 4, Worker: true, CancelAfterRows: 500));
            cases.Add(new(format, 10000, 4, HangSink: true));
            cases.Add(new(format, 10000, 4, ResourceLimit: true));
            cases.Add(new(format, 10000, 20, Styled: true, Unique: true, DelayedPages: true));
            cases.Add(new(format, 10000, 4, PendingPage: true));
            cases.Add(new(format, 10000, 4, Worker: true, PendingPage: true));
            cases.Add(new(format, 100000, 20, Styled: true, Unique: true, Worker: true));
            cases.Add(new(format, 100000, 20, Styled: true, Worker: true, Fallback: true));
        }
        if (args.Contains("--conditional")) cases = cases.Where(spec => spec.Format == "xlsx" && (!spec.Unique || spec.Worker || spec.DelayedPages))
            .Select(spec => spec with { Conditional = true }).ToList();
        string script = BrowserAssets.Script.Content, qualification = File.ReadAllText(Path.Combine(repository, "Build", "BrowserExports", "qualification.js"));
        string host = Path.Combine(evidence, "qualification.html");
        File.WriteAllText(Path.Combine(evidence, BrowserAssets.Script.HashedFileName), script, new UTF8Encoding(false));
        File.WriteAllText(host, "<!doctype html><meta charset=\"utf-8\"><title>OfficeIMO export qualification</title><script src=\"" + BrowserAssets.Script.HashedFileName + "\"></script>");
        var engines = new[] { HtmlBrowserEngine.Chromium, HtmlBrowserEngine.Firefox, HtmlBrowserEngine.WebKit };
        string? selected = args.FirstOrDefault(a => a.StartsWith("--engine=", StringComparison.Ordinal))?.Substring(9);
        if (selected is not null) engines = engines.Where(e => string.Equals(e.ToString(), selected, StringComparison.OrdinalIgnoreCase)).ToArray();
        if (engines.Length == 0) throw new ArgumentException("Unknown browser engine.");
        foreach (var engine in engines) {
            string output = Path.Combine(evidence, engine.ToString().ToLowerInvariant()); Directory.CreateDirectory(output);
            var launch = new HtmlBrowserLaunchOptions { Browser = engine, Headless = true, Timeout = 120000, LoadState = HtmlBrowserLoadState.Load };
            if (engine == HtmlBrowserEngine.Chromium) launch.BrowserArguments.Add("--enable-precise-memory-info");
            await using var session = await HtmlBrowser.OpenSessionAsync(new Uri(host).AbsoluteUri, launch);
            var errors = new List<string>(); session.Page.PageError += (_, error) => errors.Add(error);
            FileStream? destination = null;
            await session.Page.ExposeFunctionAsync("acceptQualificationChunk", async (string base64) => {
                byte[] bytes = Convert.FromBase64String(base64);
                if (bytes.Length > 65536 || destination is null) throw new InvalidDataException("Invalid qualification output chunk.");
                await destination.WriteAsync(bytes); return true;
            });
            await session.Page.AddScriptTagAsync(new() { Content = qualification });
            foreach (Case spec in cases) {
                string path = Path.Combine(output, spec.Name + "." + spec.Format);
                Console.WriteLine($"Qualifying {engine}/{spec.Name}");
                JsonElement metrics;
                await using (destination = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 65536, FileOptions.Asynchronous)) {
                    metrics = await session.Page.EvaluateAsync<JsonElement>("args => runQualificationCase(args)", new {
                        format = spec.Format, rows = spec.Rows, columns = spec.Columns, styled = spec.Styled, unique = spec.Unique, longText = spec.LongText,
                        worker = spec.Worker, fallback = spec.Fallback, slowSink = spec.SlowSink, cancelAfterRows = spec.CancelAfterRows, conditional = spec.Conditional,
                        hangSink = spec.HangSink, resourceLimit = spec.ResourceLimit, delayedPages = spec.DelayedPages, pendingPage = spec.PendingPage, workerScript = script, qualificationScript = qualification
                    });
                }
                destination = null;
                bool rejected = metrics.GetProperty("rejected").ValueKind != JsonValueKind.Null;
                // Null is omitted at the JS boundary; null cancelAfterRows is normalized by the scenario.
                object proof = rejected ? QualificationVerifier.VerifyFailure(path, spec) : QualificationVerifier.Verify(path, spec);
                if (errors.Count != 0) throw new InvalidDataException("Browser errors: " + string.Join("; ", errors));
                string sha;
                using (Stream hashSource = File.OpenRead(path)) sha = Convert.ToHexString(SHA256.HashData(hashSource)).ToLowerInvariant();
                var report = new { engine = engine.ToString(), version = session.Browser?.Version, scriptHash = BrowserAssets.Script.HashedFileName,
                    utc = DateTime.UtcNow, metrics, proof, sha256 = sha, retained = args.Contains("--keep") };
                File.WriteAllText(Path.Combine(output, spec.Name + ".json"), JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
                Console.WriteLine($"Passed {engine}/{spec.Name}: {metrics.GetProperty("outputBytes")} bytes, {metrics.GetProperty("elapsedMs")} ms; every value validated.");
                if (!args.Contains("--keep")) File.Delete(path);
            }
        }
    }
}
