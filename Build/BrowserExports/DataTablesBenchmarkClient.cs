#nullable enable
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;

/// <summary>Test-only process lifetime and JSON-line transport for the shared benchmark runner.</summary>
public sealed class DataTablesBenchmarkClient : IDisposable {
    public static DataTablesBenchmarkClient? Current { get; set; }
    /// <summary>Task-keyed opt-in Node sessions; each wrapper removes and disposes its own keys.</summary>
    public static readonly Dictionary<string, DataTablesBenchmarkClient> Sessions = new(StringComparer.Ordinal);
    private readonly Process process;
    private readonly Task<string> errors;
    private bool disposed;

    /// <summary>Recorded child identity for task-owned lifecycle evidence.</summary>
    public int ProcessId => process.Id;

    public DataTablesBenchmarkClient(string binary, string repository, string output, string assets)
        : this("dotnet", new[] { binary, repository, output, "--datatables-session", "--comparison-assets=" + assets }, repository) { }

    /// <summary>Start one explicitly selected test-only executable and JSON-line adapter.</summary>
    public DataTablesBenchmarkClient(string executable, string[] arguments, string workingDirectory) {
        var start = new ProcessStartInfo(executable) { UseShellExecute = false, CreateNoWindow = true,
            RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true, WorkingDirectory = workingDirectory };
        foreach (string argument in arguments) start.ArgumentList.Add(argument);
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
        if (disposed) return;
        disposed = true;
        if (!process.HasExited) {
            try {
                process.StandardInput.WriteLine("{\"command\":\"close\"}"); process.StandardInput.Flush();
                // Closing our pipe also releases readers that keep stdin open after the close message.
                process.StandardInput.Close();
            } catch (System.IO.IOException) { }
            if (!process.WaitForExit(15000)) process.Kill(true);
        }
        process.Dispose();
    }
}
