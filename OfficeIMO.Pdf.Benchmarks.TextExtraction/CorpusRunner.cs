using System.Collections.Concurrent;
using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Text.Json;

namespace OfficeIMO.Pdf.Benchmarks.TextExtraction;

/// <summary>
/// Runs every engine over every corpus file in child processes and records time, outcome and text.
/// <list type="bullet">
/// <item><c>cold</c>: a new process for each file and engine. Nothing is warmed up, so .NET engines also pay
/// for JIT compilation on every file.</item>
/// <item><c>warm</c>: one long-lived process per engine that receives files in sequence, the way a corpus
/// script loops in one interpreter. A crash or timeout replaces the process and is recorded for that file.</item>
/// </list>
/// Engines are interleaved per file so machine load drifts affect all of them equally.
/// </summary>
internal static class CorpusRunner {
    internal static int Run(string[] args) {
        string mode = Options.Get(args, "--mode") ?? "warm";
        if (mode is not ("cold" or "warm")) {
            throw new ArgumentException("--mode must be cold or warm.");
        }
        string output = Path.GetFullPath(Options.Get(args, "--output") ?? throw new ArgumentException("--output is required."));
        string[] engines = (Options.Get(args, "--engines") ?? string.Join(',', TextEngines.All))
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        foreach (string engine in engines) {
            TextEngines.Resolve(engine);
        }
        TimeSpan timeout = TimeSpan.FromSeconds(double.Parse(Options.Get(args, "--timeout") ?? "60", CultureInfo.InvariantCulture));
        (int shard, int shardCount) = ParseShard(Options.Get(args, "--shard") ?? "1/1");
        bool writeText = Options.Has(args, "--write-text");
        string textRoot = Path.Combine(output, "texts");
        List<(string Suite, string Root)> suites = Options.GetAll(args, "--suite")
            .Select(static value => {
                int separator = value.IndexOf('=');
                if (separator <= 0) {
                    throw new ArgumentException($"--suite expects name=directory, got '{value}'.");
                }
                return (value[..separator], Path.GetFullPath(value[(separator + 1)..]));
            })
            .ToList();
        if (suites.Count == 0) {
            throw new ArgumentException("At least one --suite name=directory is required.");
        }

        Directory.CreateDirectory(output);
        // Runs with different engine selections write separate files; the summary keeps the newest row per engine.
        string resultsPath = Path.Combine(output,
            $"results-{mode}-{string.Join('+', suites.Select(static s => s.Suite))}-{string.Join('+', engines)}-shard{shard}of{shardCount}.csv");
        using var csv = new StreamWriter(resultsPath, append: false, new UTF8Encoding(false));
        csv.WriteLine("suite,file,engine,mode,status,ms,wall_ms,chars,error_type,error_site,error_message");

        var workers = new Dictionary<string, EngineProcess>(StringComparer.Ordinal);
        int rotation = 0;
        try {
            foreach ((string suite, string root) in suites) {
                string[] files = Directory.EnumerateFiles(root, "*", SearchOption.AllDirectories)
                    .Where(static file => file.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                    .OrderBy(static file => file, StringComparer.Ordinal)
                    .ToArray();
                string[] selected = files.Where((_, index) => index % shardCount == shard - 1).ToArray();
                Console.WriteLine($"{suite}: {files.Length} files, shard {shard}/{shardCount} runs {selected.Length}.");
                var clock = Stopwatch.StartNew();
                for (int index = 0; index < selected.Length; index++) {
                    string file = selected[index];
                    string relative = Path.GetRelativePath(root, file);
                    for (int offset = 0; offset < engines.Length; offset++) {
                        string engine = engines[(offset + rotation) % engines.Length];
                        string? textPath = writeText ? Path.Combine(textRoot, engine, suite, relative + ".txt") : null;
                        DeleteText(textPath);
                        Outcome outcome;
                        if (mode == "cold") {
                            using var process = new EngineProcess(engine);
                            outcome = process.Extract(file, textPath, timeout);
                        } else {
                            if (!workers.TryGetValue(engine, out EngineProcess? worker) || !worker.IsAlive) {
                                worker?.Dispose();
                                worker = new EngineProcess(engine);
                                workers[engine] = worker;
                            }
                            outcome = worker.Extract(file, textPath, timeout);
                            if (outcome.Status is "timeout" or "crash") {
                                worker.Dispose();
                                workers.Remove(engine);
                            }
                        }
                        if (outcome.Status != "ok") {
                            // Only successful extractions may leave text for parity; the process has exited or
                            // answered, so nothing is still writing these paths.
                            DeleteText(textPath);
                        }
                        csv.WriteLine(string.Join(',',
                            Csv(suite), Csv(relative), engine, mode, outcome.Status,
                            outcome.Ms.ToString("0.###", CultureInfo.InvariantCulture),
                            outcome.WallMs.ToString("0.###", CultureInfo.InvariantCulture),
                            outcome.Chars.ToString(CultureInfo.InvariantCulture),
                            Csv(outcome.ErrorType), Csv(outcome.ErrorSite), Csv(outcome.ErrorMessage)));
                    }
                    rotation++;
                    if ((index + 1) % 250 == 0) {
                        csv.Flush();
                        Console.WriteLine($"  {suite}: {index + 1}/{selected.Length} in {clock.Elapsed.TotalSeconds:0}s");
                    }
                }
                csv.Flush();
                Console.WriteLine($"  {suite}: done in {clock.Elapsed.TotalSeconds:0}s");
            }
        } finally {
            foreach (EngineProcess worker in workers.Values) {
                worker.Dispose();
            }
        }
        Console.WriteLine($"Results: {resultsPath}");
        return 0;
    }

    private static (int Shard, int Count) ParseShard(string value) {
        string[] parts = value.Split('/');
        int shard = int.Parse(parts[0], CultureInfo.InvariantCulture);
        int count = int.Parse(parts[1], CultureInfo.InvariantCulture);
        if (count < 1 || shard < 1 || shard > count) {
            throw new ArgumentException("--shard expects i/n with 1 <= i <= n.");
        }
        return (shard, count);
    }

    private static void DeleteText(string? textPath) {
        if (textPath is null) {
            return;
        }
        foreach (string path in new[] { textPath, textPath + ExtractionWorker.PartialSuffix }) {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    internal static string Csv(string? value) {
        if (string.IsNullOrEmpty(value)) {
            return string.Empty;
        }
        string flattened = value.Replace('\r', ' ').Replace('\n', ' ');
        return flattened.IndexOfAny([',', '"']) >= 0 ? "\"" + flattened.Replace("\"", "\"\"") + "\"" : flattened;
    }

    internal sealed record Outcome(string Status, double Ms, double WallMs, int Chars, string? ErrorType, string? ErrorSite, string? ErrorMessage);

    /// <summary>One engine child process and its result channel.</summary>
    private sealed class EngineProcess : IDisposable {
        private readonly Process _process;
        private readonly BlockingCollection<string> _results = new();
        private readonly ConcurrentQueue<string> _stderr = new();
        private readonly long _started = Stopwatch.GetTimestamp();
        private bool _firstRequest = true;

        internal EngineProcess(string engine) {
            var start = new ProcessStartInfo {
                FileName = Environment.ProcessPath!,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                StandardInputEncoding = ExtractionWorker.Utf8,
                StandardOutputEncoding = ExtractionWorker.Utf8,
                StandardErrorEncoding = ExtractionWorker.Utf8
            };
            string entryAssembly = typeof(CorpusRunner).Assembly.Location;
            if (Path.GetFileNameWithoutExtension(Environment.ProcessPath!).Equals("dotnet", StringComparison.OrdinalIgnoreCase)) {
                start.ArgumentList.Add(entryAssembly);
            }
            start.ArgumentList.Add("worker");
            start.ArgumentList.Add("--engine");
            start.ArgumentList.Add(engine);
            _process = Process.Start(start) ?? throw new InvalidOperationException($"Could not start the {engine} worker.");
            _process.OutputDataReceived += (_, e) => {
                if (e.Data is null) {
                    _results.CompleteAdding();
                } else if (e.Data.StartsWith(ExtractionWorker.ResultPrefix, StringComparison.Ordinal)) {
                    _results.TryAdd(e.Data[ExtractionWorker.ResultPrefix.Length..]);
                }
            };
            _process.ErrorDataReceived += (_, e) => {
                if (e.Data is not null) {
                    _stderr.Enqueue(e.Data);
                    while (_stderr.Count > 20) {
                        _stderr.TryDequeue(out string? _);
                    }
                }
            };
            _process.BeginOutputReadLine();
            _process.BeginErrorReadLine();
        }

        internal bool IsAlive => !_process.HasExited && !_results.IsAddingCompleted;

        internal Outcome Extract(string file, string? textPath, TimeSpan timeout) {
            long requested = _firstRequest ? _started : Stopwatch.GetTimestamp();
            _firstRequest = false;
            try {
                _process.StandardInput.WriteLine(JsonSerializer.Serialize(new WorkerRequest(file, textPath)));
                _process.StandardInput.Flush();
            } catch (IOException) {
                return Crash(requested);
            }
            try {
                if (_results.TryTake(out string? json, timeout)) {
                    WorkerResult result = JsonSerializer.Deserialize<WorkerResult>(json)!;
                    return new Outcome(result.Status, result.Ms, Stopwatch.GetElapsedTime(requested).TotalMilliseconds,
                        result.Chars, result.ErrorType, result.ErrorSite, result.ErrorMessage);
                }
            } catch (InvalidOperationException) {
                // The output channel completed: the process exited without a result.
            }
            if (_results.IsAddingCompleted || _process.HasExited) {
                return Crash(requested);
            }
            Kill();
            return new Outcome("timeout", timeout.TotalMilliseconds, Stopwatch.GetElapsedTime(requested).TotalMilliseconds, 0, null, null, null);
        }

        private Outcome Crash(long requested) {
            _process.WaitForExit(5000);
            string exitCode = _process.HasExited ? "0x" + _process.ExitCode.ToString("X8", CultureInfo.InvariantCulture) : "unknown";
            string tail = string.Join(" | ", _stderr.TakeLast(3));
            return new Outcome("crash", 0, Stopwatch.GetElapsedTime(requested).TotalMilliseconds, 0,
                "exit " + exitCode, null, tail.Length > 240 ? tail[..240] : tail);
        }

        private void Kill() {
            try {
                if (!_process.HasExited) {
                    _process.Kill(entireProcessTree: true);
                    _process.WaitForExit(5000);
                }
            } catch (InvalidOperationException) {
            }
        }

        public void Dispose() {
            try {
                if (!_process.HasExited) {
                    _process.StandardInput.Close();
                    if (!_process.WaitForExit(5000)) {
                        Kill();
                    }
                }
            } catch (IOException) {
                Kill();
            }
            if (_process.HasExited) {
                // The parameterless wait returns only after the redirected output and error callbacks have drained,
                // so no callback can touch the result channel after it is disposed.
                _process.WaitForExit();
                _results.Dispose();
            }
            _process.Dispose();
        }
    }
}

internal static class Options {
    internal static string? Get(string[] args, string name) {
        for (int i = 0; i < args.Length - 1; i++) {
            if (string.Equals(args[i], name, StringComparison.OrdinalIgnoreCase)) {
                return args[i + 1];
            }
        }
        return null;
    }

    internal static IEnumerable<string> GetAll(string[] args, string name) {
        for (int i = 0; i < args.Length - 1; i++) {
            if (string.Equals(args[i], name, StringComparison.OrdinalIgnoreCase)) {
                yield return args[i + 1];
            }
        }
    }

    internal static bool Has(string[] args, string name) =>
        args.Any(arg => string.Equals(arg, name, StringComparison.OrdinalIgnoreCase));
}
