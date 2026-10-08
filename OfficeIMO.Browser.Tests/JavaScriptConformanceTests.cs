#if NET8_0_OR_GREATER
using System;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Browser.Tests;

/// <summary>Independent C# readers and SDK validation qualify the shared TypeScript writer corpus.</summary>
public sealed class JavaScriptConformanceTests {
    [Fact]
    public async Task TypeScriptCorpusOpensInBothReadersPassesSdkValidationAndMatchesCsvBytes() {
        string repository = FindRepository();
        string? evidenceRoot = Environment.GetEnvironmentVariable("OFFICEIMO_JS_EVIDENCE_DIR");
        string output = evidenceRoot is null
            ? Path.Combine(Path.GetTempPath(), "officeimo-js-conformance-" + Guid.NewGuid().ToString("N"))
            : Path.Combine(Path.GetFullPath(evidenceRoot), "dotnet-" + System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription.Replace(' ', '-'));
        Directory.CreateDirectory(output);
        try {
            var start = new ProcessStartInfo("node") { WorkingDirectory = repository, RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false };
            start.ArgumentList.Add(Path.Combine(repository, "OfficeIMO.JavaScript", "scripts", "fixtures.mjs"));
            start.ArgumentList.Add(output);
            using Process process = Process.Start(start) ?? throw new InvalidOperationException("Cannot start Node for TypeScript conformance.");
            var stdout = process.StandardOutput.ReadToEndAsync();
            var stderr = process.StandardError.ReadToEndAsync();
            using var timeout = new CancellationTokenSource(TimeSpan.FromMinutes(2));
            try { await process.WaitForExitAsync(timeout.Token); }
            catch (OperationCanceledException) { process.Kill(entireProcessTree: true); await process.WaitForExitAsync(); throw new TimeoutException("TypeScript fixture generation timed out."); }
            Assert.True(process.ExitCode == 0, await stdout + await stderr);
            using JsonDocument fixtures = JsonDocument.Parse(File.ReadAllText(Path.Combine(repository, "OfficeIMO.TestAssets", "JavaScript", "xlsx-writer.json")));
            foreach (JsonElement spec in fixtures.RootElement.GetProperty("cases").EnumerateArray()) {
                foreach (string compression in new[] { "auto", "store" }) {
                    string path = Path.Combine(output, spec.GetProperty("name").GetString() + "-" + compression + ".xlsx");
                    JavaScriptWorkbookContract.Verify(path, spec);
                }
            }
            using JsonDocument csv = JsonDocument.Parse(File.ReadAllText(Path.Combine(repository, "OfficeIMO.TestAssets", "CSV", "browser-exports.json")));
            foreach (JsonElement vector in csv.RootElement.GetProperty("cases").EnumerateArray()) {
                byte[] expected = BrowserCsvVectorContract.Write(vector);
                Assert.Equal(BrowserCsvVectorContract.Expected(vector), expected);
                Assert.Equal(expected, File.ReadAllBytes(Path.Combine(output, vector.GetProperty("name").GetString() + ".csv")));
            }
        } finally {
            // A configured evidence directory is deliberately retained for the reviewer.
            if (evidenceRoot is null) Directory.Delete(output, recursive: true);
        }
    }

    private static string FindRepository([CallerFilePath] string sourceFile = "") {
        foreach (string start in new[] { Path.GetDirectoryName(sourceFile)!, AppContext.BaseDirectory })
            for (DirectoryInfo? directory = new DirectoryInfo(start); directory is not null; directory = directory.Parent)
                if (File.Exists(Path.Combine(directory.FullName, "OfficeIMO.JavaScript", "package.json"))) return directory.FullName;
        throw new InvalidOperationException("Run this test from a source checkout after npm ci and npm run build in OfficeIMO.JavaScript.");
    }
}
#endif
