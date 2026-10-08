#nullable enable
using System;
using System.Diagnostics;
using System.Threading.Tasks;

/// <summary>Test-only process lifetime and JSON-line transport for the shared benchmark runner.</summary>
public sealed class DataTablesBenchmarkClient : IDisposable {
    public static DataTablesBenchmarkClient? Current { get; set; }
    private readonly Process process;
    private readonly Task<string> errors;

    public DataTablesBenchmarkClient(string binary, string repository, string output, string assets) {
        var start = new ProcessStartInfo("dotnet") { UseShellExecute = false, CreateNoWindow = true,
            RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true, WorkingDirectory = repository };
        foreach (string argument in new[] { binary, repository, output, "--datatables-session", "--comparison-assets=" + assets }) start.ArgumentList.Add(argument);
        process = Process.Start(start) ?? throw new InvalidOperationException("Unable to start comparison session.");
        errors = process.StandardError.ReadToEndAsync();
        try {
            string ready = ReadLine();
            if (ready != "{\"ready\":true}") throw new InvalidOperationException("Unexpected comparison startup: " + ready);
        } catch { Dispose(); throw; }
    }

    public string Request(string json) {
        if (process.HasExited) throw new InvalidOperationException("Comparison session exited: " + errors.GetAwaiter().GetResult());
        process.StandardInput.WriteLine(json); process.StandardInput.Flush(); return ReadLine();
    }
    private string ReadLine() {
        string? line = process.StandardOutput.ReadLineAsync().WaitAsync(TimeSpan.FromMinutes(12)).GetAwaiter().GetResult();
        return line ?? throw new InvalidOperationException("Comparison session ended: " + errors.GetAwaiter().GetResult());
    }
    public void Dispose() {
        if (!process.HasExited) {
            try { process.StandardInput.WriteLine("{\"command\":\"close\"}"); process.StandardInput.Flush(); } catch (System.IO.IOException) { }
            if (!process.WaitForExit(15000)) process.Kill(true);
        }
        process.Dispose();
    }
}
