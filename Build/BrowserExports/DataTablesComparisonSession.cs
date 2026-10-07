using System.Security.Cryptography;
using System.Text.Json;
using HtmlTinkerX;
using OfficeIMO.Browser;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

/// <summary>JSON-line control of browser setup, one export operation, and independent validation.</summary>
internal static class DataTablesComparisonSession {
    internal static async Task RunAsync(string repository, string evidence, string[] args) {
        string assets = Path.GetFullPath(args.FirstOrDefault(a => a.StartsWith("--comparison-assets=", StringComparison.Ordinal))?.Substring(20)
            ?? throw new ArgumentException("Comparison assets are required."));
        using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(repository, "Build", "BrowserExports", "comparison-assets.json")));
        DataTablesInterop.VerifyAssets(assets, manifest.RootElement);
        HtmlBrowserSession? session = null;
        string? identity = null;
        JsonElement spec = default, measurement = default;
        FileStream? destination = null;
        var errors = new List<string>();
        Console.WriteLine("{\"ready\":true}");
        try {
            while (await Console.In.ReadLineAsync() is { } line) {
                try {
                    using var request = JsonDocument.Parse(line); JsonElement command = request.RootElement;
                    string operation = command.GetProperty("command").GetString()!;
                    if (operation == "close") { Console.WriteLine("{\"ok\":true}"); break; }
                    object result;
                    if (operation == "prepare") {
                        string stack = command.GetProperty("stack").GetString()!, browser = command.GetProperty("browser").GetString()!;
                        spec = command.Clone(); errors.Clear();
                        int rows = spec.GetProperty("rows").GetInt32(), columns = spec.GetProperty("columns").GetInt32();
                        if (rows < 1 || rows > 1000000 || columns < 1 || columns > 100 || (long)rows * columns > 20000000)
                            throw new ArgumentException("Comparison shape exceeds the bounded 20-million-cell matrix.");
                        if (identity != stack + "/" + browser || session is null) {
                            if (session is not null) await session.DisposeAsync(); session = null;
                            var engine = DataTablesInterop.Engines(new[] { "--engine=" + browser }).Single();
                            string output = Path.Combine(evidence, stack, browser); Directory.CreateDirectory(output);
                            session = await DataTablesInterop.OpenAsync(assets, manifest.RootElement, stack, engine, output);
                            session.Page.PageError += (_, error) => errors.Add(error);
                            session.Page.Download += (_, download) => { _ = download.CancelAsync().ContinueWith(task => { _ = task.Exception; }, TaskContinuationOptions.OnlyOnFaulted); };
                            await session.Page.ExposeFunctionAsync("acceptComparisonChunk", async (string base64) => {
                                byte[] bytes = Convert.FromBase64String(base64);
                                if (bytes.Length > 65536 || destination is null) throw new InvalidDataException("Invalid comparison output chunk.");
                                await destination.WriteAsync(bytes); return true;
                            });
                            await session.Page.AddScriptTagAsync(new() { Content = File.ReadAllText(Path.Combine(repository, "Build", "BrowserExports", "datatables-measure.js")) });
                            identity = stack + "/" + browser;
                        }
                        result = await session.Page.EvaluateAsync<JsonElement>("spec => prepareDataTablesMeasurement(spec)", new { rows, columns, unique = spec.GetProperty("unique").GetBoolean() }).WaitAsync(TimeSpan.FromMinutes(5));
                    } else if (operation == "run") {
                        if (session is null) throw new InvalidOperationException("Prepare the browser first.");
                        bool profile = command.TryGetProperty("profile", out var requestedProfile) && requestedProfile.GetBoolean();
                        if (profile && !string.Equals(spec.GetProperty("browser").GetString(), "Chromium", StringComparison.OrdinalIgnoreCase)) throw new ArgumentException("CPU profiling requires Chromium.");
                        var profiler = profile ? await session.Page.Context.NewCDPSessionAsync(session.Page) : null;
                        try {
                            if (profiler is not null) { await profiler.SendAsync("Profiler.enable"); await profiler.SendAsync("Profiler.start"); }
                            measurement = await session.Page.EvaluateAsync<JsonElement>("args => runDataTablesMeasurement(args.lane, args.format)",
                                new { lane = command.GetProperty("lane").GetString(), format = spec.GetProperty("format").GetString() }).WaitAsync(TimeSpan.FromMinutes(10));
                        } finally {
                            if (profiler is not null) {
                                try {
                                    var trace = await profiler.SendAsync("Profiler.stop");
                                    File.WriteAllText(Path.Combine(evidence, "export.cpuprofile"), trace!.Value.GetProperty("profile").GetRawText());
                                } finally { await profiler.DetachAsync(); }
                            }
                        }
                        result = measurement;
                    } else if (operation == "validate") {
                        if (session is null) throw new InvalidOperationException("No browser measurement to validate.");
                        string format = spec.GetProperty("format").GetString()!, path = Path.Combine(evidence, "current-output." + format);
                        await using (destination = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 65536, FileOptions.Asynchronous))
                            await session.Page.EvaluateAsync("() => deliverDataTablesMeasurement()");
                        destination = null;
                        object proof;
                        try { proof = DataTablesComparisonVerifier.Verify(path, format, spec.GetProperty("rows").GetInt32(), spec.GetProperty("columns").GetInt32(), spec.GetProperty("unique").GetBoolean()); }
                        catch {
                            string failed = Path.Combine(evidence, "first-failed-output." + format);
                            if (!File.Exists(failed) && new FileInfo(path).Length <= 64L * 1024 * 1024) File.Move(path, failed);
                            else File.Delete(path);
                            throw;
                        }
                        if (errors.Count != 0) throw new InvalidDataException(string.Join("; ", errors));
                        if (new FileInfo(path).Length != measurement.GetProperty("outputBytes").GetInt64()) throw new InvalidDataException("Transferred file size differs.");
                        string[] schemaErrors = [];
                        if (format == "xlsx") {
                            using var document = SpreadsheetDocument.Open(path, false);
                            schemaErrors = new OpenXmlValidator().Validate(document.WorkbookPart!.WorkbookStylesPart!).Take(20).Select(error => error.Description).ToArray();
                        }
                        if (format == "xlsx" && schemaErrors.Length == 0 && spec.GetProperty("rows").GetInt32() <= 10000)
                            WorkbookVerifier.Verify(path, 128L * 1024 * 1024);
                        string hash; using (Stream file = File.OpenRead(path)) hash = Convert.ToHexString(SHA256.HashData(file)).ToLowerInvariant();
                        result = new { proof, conformance = new { passed = schemaErrors.Length == 0, stylesErrors = schemaErrors }, sha256 = hash, browserVersion = session.Browser?.Version, asset = BrowserAssets.DataTablesScript.HashedFileName };
                        File.Delete(path);
                    } else throw new ArgumentException("Unknown comparison command.");
                    Console.WriteLine(JsonSerializer.Serialize(new { ok = true, result }));
                } catch (Exception error) {
                    if (session is not null) {
                        try { await session.DisposeAsync().AsTask().WaitAsync(TimeSpan.FromSeconds(15)); } catch { /* Caller also owns a bounded process lifetime. */ }
                        session = null; identity = null;
                    }
                    Console.WriteLine(JsonSerializer.Serialize(new { ok = false, error = error.ToString() }));
                }
            }
        } finally { if (session is not null) await session.DisposeAsync(); }
    }
}
