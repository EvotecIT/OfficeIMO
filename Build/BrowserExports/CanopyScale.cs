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
        var sizes = args.Contains("--full") ? new[] { 10000, 100000, 250000, 1000000 } : new[] { 10000, 100000 };
        long removed = 0;
        foreach (int rows in sizes) foreach (int columns in rows > 100000 ? new[] { 4 } : new[] { 4, 20 }) foreach (string format in new[] { "csv", "xlsx" }) {
            string navigation = rows == 10000 || rows == 250000 ? "offset" : "cursor";
            var spec = new ExportQualification.Case(format, rows, columns, Unique: true, Fallback: rows == 10000, SlowSink: rows == 10000);
            string name = "canopy-scale-" + spec.Name, path = Path.Combine(output, name + "." + format);
            Console.WriteLine($"Canopy paged qualification {session.Browser?.Version}/{name}/{navigation}");
            JsonElement metrics;
            await using (destination = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 65536, FileOptions.Asynchronous)) {
                metrics = await session.Page.EvaluateAsync<JsonElement>("args => runCanopyScale(args)", new { rows, columns, format, navigation, slowSink = spec.SlowSink, fallback = spec.Fallback });
            }
            destination = null;
            object proof = QualificationVerifier.Verify(path, spec);
            long length = new FileInfo(path).Length;
            File.WriteAllText(Path.Combine(output, name + ".json"), JsonSerializer.Serialize(new { metrics, proof, asset = BrowserAssets.CanopyXScript.HashedFileName,
                bytes = length, retained = false, timing = "diagnostic; includes native capture, acknowledged file bridge and output generation" }, new JsonSerializerOptions { WriteIndented = true }));
            File.Delete(path); removed += length;
            Console.WriteLine($"Passed {name}: {rows * (long)columns} cells, {length} bytes removed after independent readback.");
        }
        File.WriteAllText(Path.Combine(output, "canopy-scale-cleanup.json"), JsonSerializer.Serialize(new { logicalBytesRemoved = removed, successfulLargeOutputsRetained = 0 }));
    }
}
