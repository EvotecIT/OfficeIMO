using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeIMO.GoogleWorkspace.Benchmarks;

internal static class GoogleWorkspaceTransportEvidence {
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true,
        Converters = { new JsonStringEnumConverter() }
    };

    public static async Task<int> RunAsync(string[] args) {
        string outputPath = ResolveOutputPath(args);
        var results = new List<TransportEvidenceResult>();
        foreach (int payloadBytes in new[] { 64 * 1024, 4 * 1024 * 1024 }) {
            foreach (ResponseLengthMode lengthMode in Enum.GetValues<ResponseLengthMode>()) {
                results.Add(await RunChildProbeAsync(payloadBytes, lengthMode).ConfigureAwait(false));
            }
        }

        var report = new TransportEvidenceReport {
            CreatedUtc = DateTimeOffset.UtcNow,
            SourceCommit = RunGit("rev-parse HEAD"),
            SourceTreeDirty = RunGit("status --porcelain").Length > 0,
            Runtime = RuntimeInformation.FrameworkDescription,
            OperatingSystem = RuntimeInformation.OSDescription,
            Results = results
        };
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath) ?? ".");
        await File.WriteAllTextAsync(outputPath, JsonSerializer.Serialize(report, JsonOptions)).ConfigureAwait(false);
        Console.WriteLine("Google Workspace transport evidence: " + outputPath);
        return results.All(result => result.Validated) ? 0 : 1;
    }

    public static async Task<int> RunProbeAsync(string[] args) {
        if (args.Length != 2
            || !Enum.TryParse(args[0], ignoreCase: true, out ResponseLengthMode lengthMode)
            || !int.TryParse(args[1], out int payloadBytes)
            || payloadBytes < 0) {
            Console.Error.WriteLine("Usage: probe <declared|unknown> <payload-bytes>");
            return 2;
        }

        byte[] payload = GoogleWorkspaceTransportScenario.CreatePayload(payloadBytes);
        using GoogleWorkspaceHttpTransport transport =
            GoogleWorkspaceTransportScenario.CreateTransport(payload, lengthMode, out HttpClient client);
        using (client) {
            for (var index = 0; index < 2; index++) {
                byte[] warmup = await GoogleWorkspaceTransportScenario.DownloadAsync(transport, payloadBytes)
                    .ConfigureAwait(false);
                GC.KeepAlive(warmup);
            }

            GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
            long heapBefore = GC.GetTotalMemory(forceFullCollection: false);
            using Process process = Process.GetCurrentProcess();
            process.Refresh();
            long workingSetBefore = process.WorkingSet64;
            using var sampler = new TransportMemorySampler(process);
            long allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
            var stopwatch = Stopwatch.StartNew();
            byte[] result = await GoogleWorkspaceTransportScenario.DownloadAsync(transport, payloadBytes)
                .ConfigureAwait(false);
            stopwatch.Stop();
            TransportMemoryPeak peak = sampler.Stop();
            long allocated = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore;
            bool validated = result.AsSpan().SequenceEqual(payload);
            string hash = Convert.ToHexString(SHA256.HashData(result)).ToLowerInvariant();
            GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
            long retained = Math.Max(0, GC.GetTotalMemory(forceFullCollection: false) - heapBefore);
            process.Refresh();
            var evidence = new TransportEvidenceResult {
                PayloadBytes = payloadBytes,
                LengthMode = lengthMode,
                OutputBytes = result.LongLength,
                ElapsedMilliseconds = stopwatch.Elapsed.TotalMilliseconds,
                AllocatedBytes = allocated,
                RetainedManagedHeapGrowthBytes = retained,
                PeakManagedHeapGrowthBytes = Math.Max(0, peak.ManagedHeapBytes - heapBefore),
                PeakWorkingSetGrowthBytes = Math.Max(0, peak.WorkingSetBytes - workingSetBefore),
                AbsoluteProcessPeakWorkingSetBytes = process.PeakWorkingSet64,
                OutputSha256 = hash,
                Validated = validated
            };
            GC.KeepAlive(result);
            Console.WriteLine(JsonSerializer.Serialize(evidence, JsonOptions));
            return validated ? 0 : 1;
        }
    }

    private static async Task<TransportEvidenceResult> RunChildProbeAsync(
        int payloadBytes,
        ResponseLengthMode lengthMode) {
        string processPath = Environment.ProcessPath
            ?? throw new InvalidOperationException("Unable to resolve the benchmark process path.");
        var startInfo = new ProcessStartInfo {
            FileName = processPath,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        if (string.Equals(Path.GetFileNameWithoutExtension(processPath), "dotnet", StringComparison.OrdinalIgnoreCase)) {
            startInfo.ArgumentList.Add(Assembly.GetEntryAssembly()!.Location);
        }
        startInfo.ArgumentList.Add("probe");
        startInfo.ArgumentList.Add(lengthMode.ToString());
        startInfo.ArgumentList.Add(payloadBytes.ToString(System.Globalization.CultureInfo.InvariantCulture));
        using Process child = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Unable to start the transport memory probe.");
        string output = await child.StandardOutput.ReadToEndAsync().ConfigureAwait(false);
        string error = await child.StandardError.ReadToEndAsync().ConfigureAwait(false);
        await child.WaitForExitAsync().ConfigureAwait(false);
        if (child.ExitCode != 0) {
            throw new InvalidOperationException("Transport memory probe failed: " + error);
        }
        return JsonSerializer.Deserialize<TransportEvidenceResult>(output, JsonOptions)
            ?? throw new InvalidOperationException("Transport memory probe returned no evidence.");
    }

    private static string ResolveOutputPath(string[] args) {
        if (args.Length == 2 && string.Equals(args[0], "--json", StringComparison.OrdinalIgnoreCase)) {
            return Path.GetFullPath(args[1]);
        }
        if (args.Length != 0) {
            throw new ArgumentException("Usage: evidence [--json <path>]");
        }
        return Path.GetFullPath(Path.Combine(".benchmark-artifacts", "googleworkspace-transport", "evidence.json"));
    }

    private static string RunGit(string arguments) {
        using var process = Process.Start(new ProcessStartInfo {
            FileName = "git",
            Arguments = arguments,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        }) ?? throw new InvalidOperationException("Unable to start git.");
        string output = process.StandardOutput.ReadToEnd();
        process.WaitForExit();
        return process.ExitCode == 0 ? output.Trim() : string.Empty;
    }
}

internal sealed class TransportMemorySampler : IDisposable {
    private readonly Process _process;
    private readonly CancellationTokenSource _stop = new();
    private readonly Task _sampling;
    private long _managedHeapBytes;
    private long _workingSetBytes;

    public TransportMemorySampler(Process process) {
        _process = process;
        Sample();
        _sampling = Task.Run(async () => {
            while (!_stop.IsCancellationRequested) {
                Sample();
                try {
                    await Task.Delay(1, _stop.Token).ConfigureAwait(false);
                } catch (OperationCanceledException) {
                    break;
                }
            }
        });
    }

    public TransportMemoryPeak Stop() {
        _stop.Cancel();
        _sampling.GetAwaiter().GetResult();
        Sample();
        return new TransportMemoryPeak(_managedHeapBytes, _workingSetBytes);
    }

    public void Dispose() {
        if (!_stop.IsCancellationRequested) Stop();
        _stop.Dispose();
    }

    private void Sample() {
        UpdateMax(ref _managedHeapBytes, GC.GetTotalMemory(forceFullCollection: false));
        _process.Refresh();
        UpdateMax(ref _workingSetBytes, _process.WorkingSet64);
    }

    private static void UpdateMax(ref long target, long value) {
        long current;
        do {
            current = Volatile.Read(ref target);
            if (value <= current) return;
        } while (Interlocked.CompareExchange(ref target, value, current) != current);
    }
}

internal readonly record struct TransportMemoryPeak(long ManagedHeapBytes, long WorkingSetBytes);

internal sealed class TransportEvidenceReport {
    public DateTimeOffset CreatedUtc { get; set; }
    public string SourceCommit { get; set; } = string.Empty;
    public bool SourceTreeDirty { get; set; }
    public string Runtime { get; set; } = string.Empty;
    public string OperatingSystem { get; set; } = string.Empty;
    public IReadOnlyList<TransportEvidenceResult> Results { get; set; } = Array.Empty<TransportEvidenceResult>();
}

internal sealed class TransportEvidenceResult {
    public int PayloadBytes { get; set; }
    public ResponseLengthMode LengthMode { get; set; }
    public long OutputBytes { get; set; }
    public double ElapsedMilliseconds { get; set; }
    public long AllocatedBytes { get; set; }
    public long RetainedManagedHeapGrowthBytes { get; set; }
    public long PeakManagedHeapGrowthBytes { get; set; }
    public long PeakWorkingSetGrowthBytes { get; set; }
    public long AbsoluteProcessPeakWorkingSetBytes { get; set; }
    public string OutputSha256 { get; set; } = string.Empty;
    public bool Validated { get; set; }
}
