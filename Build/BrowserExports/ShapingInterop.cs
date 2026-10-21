using System.Security.Cryptography;
using System.Text.Json;
using HtmlTinkerX;
using OfficeIMO.Pdf;

/// <summary>Real optional HarfBuzz WASM/bidi qualification against the shared native/Node control corpus.</summary>
internal static class ShapingInterop {
    internal static async Task RunAsync(string repository, string evidence, string[] args) {
        string Option(string name) => args.FirstOrDefault(arg => arg.StartsWith(name + "=", StringComparison.Ordinal))?[(name.Length + 1)..]
            ?? throw new ArgumentException("Shaping qualification requires " + name + ".");
        var module = new Uri(Option("--module-url").TrimEnd('/') + "/");
        if (!module.IsLoopback || module.Scheme != "http") throw new ArgumentException("Shaping modules require loopback HTTP.");
        string assets = Path.GetFullPath(Option("--shaping-assets"));
        string referencePath = args.FirstOrDefault(arg => arg.StartsWith("--shaping-reference=", StringComparison.Ordinal))?[20..] ?? Path.Combine(assets, "node-proof.json");
        using var reference = JsonDocument.Parse(File.ReadAllText(referencePath));
        foreach ((string package, string version) in new[] { ("harfbuzzjs", "1.6.3"), ("bidi-js", "1.1.0") }) {
            using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(assets, "packages", package, "package", "package.json")));
            Require(manifest.RootElement.GetProperty("name").GetString() == package && manifest.RootElement.GetProperty("version").GetString() == version, "Wrong explicit optional package identity.");
        }
        string script = File.ReadAllText(Path.Combine(repository, "Build", "BrowserExports", "shaping-contracts.js"));
        string wasm = Path.Combine(assets, "packages", "harfbuzzjs", "package", "dist", "harfbuzz.wasm");
        string wasmHash = Hash(wasm);
        var reports = new List<object>();
        foreach (HtmlBrowserEngine engine in DataTablesInterop.Engines(args)) {
            string output = Path.Combine(evidence, engine.ToString().ToLowerInvariant()); Directory.CreateDirectory(output);
            await using var session = await HtmlBrowser.OpenSessionAsync(new Uri(module, "../host.html").AbsoluteUri,
                new HtmlBrowserLaunchOptions { Browser = engine, Headless = true, Timeout = 120000 });
            var errors = new List<string>(); var assetsRead = new List<object>(); int wasmResponses = 0;
            session.Page.PageError += (_, error) => errors.Add(error);
            session.Page.Response += (_, response) => {
                if (response.Url.Contains("/__shaping__/", StringComparison.Ordinal) || response.Url.Contains("/__fonts__/", StringComparison.Ordinal)) {
                    assetsRead.Add(new { url = response.Url, status = response.Status, contentType = response.Headers.GetValueOrDefault("content-type") });
                    if (response.Url.EndsWith("/harfbuzz.wasm", StringComparison.Ordinal) && response.Status == 200 && response.Headers.GetValueOrDefault("content-type")?.StartsWith("application/wasm", StringComparison.Ordinal) == true) wasmResponses++;
                }
            };
            object? pdfReadback = null;
            await session.Page.ExposeFunctionAsync("saveShapingFixture", (string name, string base64, string contract) => {
                if (name != "shaped-report.pdf") throw new InvalidDataException("Unknown shaping fixture.");
                string path = Path.Combine(output, name); File.WriteAllBytes(path, Convert.FromBase64String(base64)); File.WriteAllText(path + ".json", contract);
                using var parsed = JsonDocument.Parse(contract); pdfReadback = VerifyPdf(path, parsed.RootElement.GetProperty("corpus")); return true;
            });
            await session.Page.AddScriptTagAsync(new() { Content = script });
            JsonElement result = await session.Page.EvaluateAsync<JsonElement>("args => runShapingContracts(args)", new {
                moduleBase = module.AbsoluteUri.TrimEnd('/'), shapingBase = new Uri(module, "/__shaping__").AbsoluteUri,
                fontBase = new Uri(module, "/__fonts__").AbsoluteUri
            });
            foreach (JsonElement control in reference.RootElement.GetProperty("controls").EnumerateArray()) {
                string name = control.GetProperty("name").GetString()!;
                JsonElement actual = result.GetProperty("controls").EnumerateArray().Single(item => item.GetProperty("name").GetString() == name);
                Require(actual.GetProperty("sha256").GetString() == control.GetProperty("sha256").GetString(), engine + " glyph/control mismatch: " + name);
            }
            string pdfHash = Hash(Path.Combine(output, "shaped-report.pdf"));
            Require(pdfHash == result.GetProperty("pdfSha256").GetString() && pdfHash == reference.RootElement.GetProperty("report").GetProperty("sha256").GetString(), "Browser PDF differs from canonical Node/native-qualified producer.");
            Require(wasmResponses == 1 && pdfReadback is not null, "The actual optional WASM response or independent PDF readback is missing.");
            Require(errors.Count == 0, "Uncaught shaping browser errors: " + string.Join("; ", errors));
            var report = new { engine = engine.ToString(), browser = session.Browser?.Version, result, pdfReadback, wasmResponses, assetsRead, errors };
            reports.Add(report); File.WriteAllText(Path.Combine(output, "report.json"), JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
            Console.WriteLine($"Optional shaping passed {engine}: real WASM, native/Node glyph controls, exact report PDF and tagged/link readback.");
        }
        Require(Hash(wasm) == wasmHash, "Optional WASM bytes changed during qualification.");
        File.WriteAllText(Path.Combine(evidence, "shaping-interop.json"), JsonSerializer.Serialize(new {
            reference = Path.GetFileName(referencePath), referenceSha256 = Hash(referencePath), wasmSha256 = wasmHash,
            packageVersions = new { harfbuzzjs = "1.6.3", bidiJs = "1.1.0" }, reports
        }, new JsonSerializerOptions { WriteIndented = true }));
    }

    private static object VerifyPdf(string path, JsonElement corpus) {
        var document = PdfReadDocument.Open(path);
        Require(!document.RepairReport.HasRepairs && document.Pages.Count >= 1, "Shaped PDF required repairs or has no page.");
        PdfTaggedContentInfo tagged = document.TaggedContent ?? throw new InvalidDataException("Shaped PDF has no structure tree.");
        foreach (string type in new[] { "Document", "H1", "P", "Table", "TR", "TH", "TD", "Link" }) Require(tagged.StructureTypes.Contains(type), "Shaped PDF lost structure type " + type);
        int links = 0;
        foreach (var page in document.Pages) {
            var size = page.GetPageSize();
            foreach (var link in page.GetLinkAnnotations()) {
                links++;
                Require(link.Uri == corpus.GetProperty("link").GetString() && link.Width > 0 && link.Height > 0 && link.X1 >= 22 - .01 && link.X2 <= size.Width - 22 + .01 && link.Y1 >= 22 - .01 && link.Y2 <= size.Height - 22 + .01, "Shaped link URI/rectangle differs or escapes content bounds.");
            }
        }
        Require(links >= 3 && tagged.StructureElements.Where(element => element.StructureType == "Link").Sum(element => element.ObjectReferenceCount) == links, "Shaped link annotations lost their structure associations.");
        // The canonical independent-parser helper checks logical MCID/ActualText order separately from visual extraction.
        return new { file = Path.GetFileName(path), pages = document.Pages.Count, links, repairs = 0, logicalActualText = "separate independent-parser qualification", tagged.StructureTypes, tagged.ObjectReferenceCount };
    }
    private static void Require(bool condition, string message) => WorkbookVerifier.Require(condition, "Optional shaping: " + message);
    private static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();
}
