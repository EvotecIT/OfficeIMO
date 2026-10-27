using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using HtmlTinkerX;
using OfficeIMO.CSV;
using OfficeIMO.Excel;

/// <summary>Real worker protocol qualification through the canonical browser/session owner.</summary>
internal static class ExecutionInterop {
    internal static async Task RunAsync(string repository, string evidence, string[] args) {
        string moduleBase = args.FirstOrDefault(a => a.StartsWith("--module-url=", StringComparison.Ordinal))?.Substring(13)
            ?? throw new ArgumentException("--execution requires --module-url pointing to current modules on loopback HTTP.");
        var moduleUri = new Uri(moduleBase.TrimEnd('/') + "/");
        if (!moduleUri.IsLoopback || moduleUri.Scheme != "http") throw new ArgumentException("Worker qualification requires loopback HTTP.");
        string script = File.ReadAllText(Path.Combine(repository, "Build", "BrowserExports", "execution-contracts.js"));
        string classicScript = File.ReadAllText(Path.Combine(repository, "OfficeIMO.JavaScript", "bundles", "officeimo.js"));
        string executionScript = File.ReadAllText(Path.Combine(repository, "OfficeIMO.JavaScript", "bundles", "officeimo-execution.js"));
        string moduleScript = File.ReadAllText(Path.Combine(repository, "OfficeIMO.JavaScript", "bundles", "officeimo.mjs"));
        string executionModule = File.ReadAllText(Path.Combine(repository, "OfficeIMO.JavaScript", "bundles", "officeimo-execution.mjs"));
        string font = Convert.ToBase64String(File.ReadAllBytes(Path.Combine(repository, "Website", "Apps", "OfficeIMO.Web.Converter", "Assets", "Fonts", "Carlito-Regular.ttf")));
        var reports = new List<object>();
        foreach (HtmlBrowserEngine engine in DataTablesInterop.Engines(args)) {
            string output = Path.Combine(evidence, engine.ToString().ToLowerInvariant()); Directory.CreateDirectory(output);
            await using HtmlBrowserSession session = await HtmlBrowser.OpenSessionAsync(new Uri(moduleUri, "../host.html").AbsoluteUri,
                new HtmlBrowserLaunchOptions { Browser = engine, Headless = true, Timeout = 120000 });
            var errors = new List<string>(); session.Page.PageError += (_, error) => errors.Add(error);
            var files = new List<object>();
            await session.Page.ExposeFunctionAsync("writeExecutionFixture", (string name, string base64, string contract) => {
                if (Path.GetFileName(name) != name || !new[] { ".csv", ".xlsx", ".pdf" }.Contains(Path.GetExtension(name)))
                    throw new InvalidDataException("Invalid worker fixture name.");
                string path = Path.Combine(output, name);
                File.WriteAllBytes(path, Convert.FromBase64String(base64)); File.WriteAllText(path + ".json", contract);
                if (name.EndsWith(".pdf", StringComparison.Ordinal)) files.Add(PdfExportVerifier.Verify(path));
                else if (name.EndsWith(".xlsx", StringComparison.Ordinal)) {
                    WorkbookVerifier.Verify(path);
                    using var reader = ExcelDocument.OpenReader(path, new ExcelReaderOptions { A1Range = "A2:C3", HasHeaderRow = false });
                    WorkbookVerifier.Require(reader.Read() && Equals(reader.GetValue(0), "Łódź") && Convert.ToDouble(reader.GetValue(1), CultureInfo.InvariantCulture) == 12.5 && Equals(reader.GetValue(2), true), "Worker XLSX typed first row differs.");
                    WorkbookVerifier.Require(reader.Read() && Equals(reader.GetValue(0), "second") && Convert.ToDouble(reader.GetValue(1), CultureInfo.InvariantCulture) == 7 && Equals(reader.GetValue(2), false) && !reader.Read(), "Worker XLSX typed second row differs.");
                    files.Add(new { file = name, bytes = new FileInfo(path).Length, readers = 2, openXml = true });
                } else {
                    using var reader = CsvDocument.OpenDataReader(path);
                    WorkbookVerifier.Require(reader.FieldCount == 3 && reader.Read() && reader.GetString(0) == "Łódź" && reader.GetString(1) == "12.5" && reader.GetString(2) == "True", "Worker CSV first row differs.");
                    WorkbookVerifier.Require(reader.Read() && reader.GetString(0) == "second" && reader.GetString(1) == "7" && reader.GetString(2) == "False" && !reader.Read(), "Worker CSV second row differs.");
                    files.Add(new { file = name, bytes = new FileInfo(path).Length, reader = "OfficeIMO.CSV" });
                }
                return true;
            });
            await session.Page.AddScriptTagAsync(new() { Content = classicScript });
            await session.Page.AddScriptTagAsync(new() { Content = executionScript });
            await session.Page.AddScriptTagAsync(new() { Content = script });
            var results = new Dictionary<string, JsonElement>();
            foreach (string profile in new[] { "esm", "esm-bundled", "classic" })
                results[profile] = await session.Page.EvaluateAsync<JsonElement>("args => runExecutionContracts(args)",
                    new { moduleBase = moduleUri.AbsoluteUri.TrimEnd('/'), classicScript, executionScript, fontBase64 = font, profile });
            if (errors.Count != 0) throw new InvalidDataException(string.Join("; ", errors));
            var report = new { engine = engine.ToString(), browser = session.Browser?.Version, results, files, errors };
            reports.Add(report);
            File.WriteAllText(Path.Combine(output, "report.json"), JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
            Console.WriteLine($"Worker execution passed {engine}: {files.Count} independently read files, classic, dist ESM and standalone ESM workers.");
        }
        File.WriteAllText(Path.Combine(evidence, "execution-interop.json"), JsonSerializer.Serialize(new {
            classicSha256 = Hash(classicScript), executionSha256 = Hash(executionScript),
            moduleSha256 = Hash(moduleScript), executionModuleSha256 = Hash(executionModule), reports
        }, new JsonSerializerOptions { WriteIndented = true }));
    }
    private static string Hash(string text) => Convert.ToHexString(SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(text))).ToLowerInvariant();
}
