using System.Text.Json;
using HtmlTinkerX;
using OfficeIMO.Browser;

/// <summary>Opt-in native paged source qualification, retaining compact reports and removing validated large output.</summary>
internal static class CanopyScale {
    internal static async Task RunAsync(HtmlBrowserSession session, string repository, string output, string[] args) {
        string script = File.ReadAllText(Path.Combine(repository, "Build", "BrowserExports", "canopy-scale.js"));
        await session.Page.AddScriptTagAsync(new() { Content = script });
        FileStream? destination = null;
        await session.Page.ExposeFunctionAsync("acceptCanopyScaleChunk", async (string base64) => {
            byte[] bytes = Convert.FromBase64String(base64);
            if (bytes.Length > 65536 || destination is null) throw new InvalidDataException("Invalid Canopy scale chunk.");
            await destination.WriteAsync(bytes); return true;
        });
        bool pdf = args.Contains("--canopy-pdf-scale");
        var sizes = !pdf && args.Contains("--full") ? new[] { 10000, 100000, 250000, 1000000 } : new[] { 10000, 100000 };
        string? regular = pdf ? Convert.ToBase64String(File.ReadAllBytes(Path.Combine(repository, "Website", "Apps", "OfficeIMO.Web.Converter", "Assets", "Fonts", "Carlito-Regular.ttf"))) : null;
        string? bold = pdf ? Convert.ToBase64String(File.ReadAllBytes(Path.Combine(repository, "Website", "Apps", "OfficeIMO.Web.Converter", "Assets", "Fonts", "Carlito-Bold.ttf"))) : null;
        long removed = 0;
        foreach (int rows in sizes) foreach (int columns in rows > 100000 ? new[] { 4 } : new[] { 4, 20 })
            foreach (string format in pdf ? new[] { "pdf" } : new[] { "csv", "xlsx" }) foreach (bool styled in pdf ? new[] { false, true } : new[] { false }) {
            string navigation = rows == 10000 || rows == 250000 ? "offset" : "cursor";
            var spec = new ExportQualification.Case(format, rows, columns, Styled: styled, Unique: !pdf, Fallback: rows == 10000, SlowSink: rows == 10000);
            string name = "canopy-scale-" + spec.Name, path = Path.Combine(output, name + "." + format);
            Console.WriteLine($"Canopy paged qualification {session.Browser?.Version}/{name}/{navigation}");
            JsonElement? metrics = null;
            Exception? failure = null;
            try {
                await using (destination = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 65536, FileOptions.Asynchronous)) {
                    metrics = await session.Page.EvaluateAsync<JsonElement>("args => runCanopyScale(args)", new { rows, columns, format, navigation, styled, regular, bold, slowSink = spec.SlowSink, fallback = spec.Fallback });
                }
                destination = null;
                object proof;
                if (pdf) {
                    File.WriteAllText(path + ".json", JsonSerializer.Serialize(new { rows, columns, required = Array.Empty<string>(), repeated = new[] { "Column 0" } }));
                    proof = PdfExportVerifier.Verify(path);
                } else proof = QualificationVerifier.Verify(path, spec);
                long length = new FileInfo(path).Length;
                File.WriteAllText(Path.Combine(output, name + ".json"), JsonSerializer.Serialize(new { metrics, proof, asset = BrowserAssets.CanopyXScript.HashedFileName,
                    bytes = length, retained = false, timing = "diagnostic; includes native capture, acknowledged file bridge and output generation" }, new JsonSerializerOptions { WriteIndented = true }));
                Console.WriteLine($"Passed {name}: {rows * (long)columns} cells, {length} bytes to remove after independent readback.");
            } catch (Exception error) {
                failure = error;
                try {
                    File.WriteAllText(Path.Combine(output, name + ".json"), JsonSerializer.Serialize(new { metrics, error = error.ToString(), rows, columns, format, navigation,
                        asset = BrowserAssets.CanopyXScript.HashedFileName, bytes = File.Exists(path) ? new FileInfo(path).Length : 0, validated = false }, new JsonSerializerOptions { WriteIndented = true }));
                } catch (Exception reportError) { Console.Error.WriteLine("Could not retain compact failure report: " + reportError.Message); }
                throw;
            } finally {
                destination = null;
                try {
                    if (File.Exists(path)) { long length = new FileInfo(path).Length; File.Delete(path); removed += length; }
                    if (pdf && File.Exists(path + ".json")) File.Delete(path + ".json");
                    File.WriteAllText(Path.Combine(output, "canopy-scale-cleanup.json"), JsonSerializer.Serialize(new { logicalBytesRemoved = removed, successfulLargeOutputsRetained = 0 }));
                } catch (Exception cleanupError) when (failure is not null) { Console.Error.WriteLine("Scale cleanup failed; original failure preserved: " + cleanupError.Message); }
            }
        }
        File.WriteAllText(Path.Combine(output, "canopy-scale-cleanup.json"), JsonSerializer.Serialize(new { logicalBytesRemoved = removed, successfulLargeOutputsRetained = 0 }));
    }
}
