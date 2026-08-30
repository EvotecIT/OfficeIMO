using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text.Json;

namespace OfficeIMO.PdfQualityCorpus;

internal static class PdfQualityCorpusCoordinator {
    internal static async Task<QualityReport> RunAsync(QualityRunOptions options, CancellationToken cancellationToken = default) {
        DateTimeOffset started = DateTimeOffset.UtcNow;
        QualityManifest manifest = PdfQualityCorpusManifest.Load(options.ManifestPath);
        foreach (QualityCase item in manifest.Cases) {
            string path = PdfQualityCorpusManifest.ResolveCasePath(options.RootDirectory, item);
            if (!File.Exists(path)) throw new FileNotFoundException("Manifested PDF corpus file was not found: " + item.Id + ".", path);
        }

        using var gate = new SemaphoreSlim(options.Parallelism, options.Parallelism);
        Task<(int Index, QualityCaseResult Result)>[] tasks = manifest.Cases.Select((item, index) =>
            RunOneAsync(index, item, options, gate, cancellationToken)).ToArray();
        (int Index, QualityCaseResult Result)[] completed = await Task.WhenAll(tasks).ConfigureAwait(false);
        QualityCaseResult[] results = completed.OrderBy(item => item.Index).Select(item => item.Result).ToArray();
        DateTimeOffset finished = DateTimeOffset.UtcNow;
        return new QualityReport {
            StartedUtc = started,
            CompletedUtc = finished,
            Authority = manifest.Authority,
            Sources = manifest.Sources,
            Environment = new QualityEnvironment {
                Framework = RuntimeInformation.FrameworkDescription,
                OperatingSystem = RuntimeInformation.OSDescription,
                ProcessArchitecture = RuntimeInformation.ProcessArchitecture.ToString(),
                EngineAssemblyVersion = typeof(OfficeIMO.Pdf.PdfDocument).Assembly.GetName().Version?.ToString() ?? "unknown"
            },
            Configuration = new QualityReportConfiguration {
                ManifestFileName = Path.GetFileName(options.ManifestPath),
                ManifestSha256 = ComputeSha256(options.ManifestPath),
                MaxFileBytes = options.MaxFileBytes,
                MaxRenderPages = options.MaxRenderPages,
                TimeoutSeconds = options.TimeoutSeconds,
                Parallelism = options.Parallelism,
                MaxWorkerMemoryBytes = options.MaxWorkerMemoryBytes
            },
            Totals = BuildTotals(results, finished - started),
            Cases = results
        };
    }

    private static string ComputeSha256(string path) {
        using FileStream stream = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(stream));
    }

    private static async Task<(int Index, QualityCaseResult Result)> RunOneAsync(
        int index,
        QualityCase item,
        QualityRunOptions options,
        SemaphoreSlim gate,
        CancellationToken cancellationToken) {
        await gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try {
            return (index, await RunWorkerAsync(item, options, cancellationToken).ConfigureAwait(false));
        } finally {
            gate.Release();
        }
    }

    private static async Task<QualityCaseResult> RunWorkerAsync(
        QualityCase item,
        QualityRunOptions options,
        CancellationToken cancellationToken) {
        ProcessStartInfo startInfo = CreateStartInfo(options, item.Id);
        using var process = new Process { StartInfo = startInfo };
        Stopwatch stopwatch = Stopwatch.StartNew();
        if (!process.Start()) return Failure(item, "worker-start-failed", stopwatch.ElapsedMilliseconds);
        Task<WorkerProcessEvidence> processEvidenceTask = SampleProcessAsync(
            process.Id,
            TryGetStartTime(process),
            options.MaxWorkerMemoryBytes);
        Task<string> outputTask = process.StandardOutput.ReadToEndAsync();
        Task<string> errorTask = process.StandardError.ReadToEndAsync();
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(TimeSpan.FromSeconds(options.TimeoutSeconds));
        try {
            await process.WaitForExitAsync(timeout.Token).ConfigureAwait(false);
        } catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested) {
            TryKill(process);
            await Task.WhenAll(outputTask, errorTask).ConfigureAwait(false);
            WorkerProcessEvidence processEvidence = CompleteProcessEvidence(
                process,
                await processEvidenceTask.ConfigureAwait(false),
                options.MaxWorkerMemoryBytes);
            var timedOut = Failure(item, "worker-timeout", stopwatch.ElapsedMilliseconds);
            AddProcessEvidence(timedOut, processEvidence, stopwatch.ElapsedMilliseconds);
            timedOut.TimedOut = true;
            timedOut.Outcome = "timed-out";
            return timedOut;
        }
        stopwatch.Stop();
        string output = await outputTask.ConfigureAwait(false);
        string error = await errorTask.ConfigureAwait(false);
        WorkerProcessEvidence evidence = CompleteProcessEvidence(
            process,
            await processEvidenceTask.ConfigureAwait(false),
            options.MaxWorkerMemoryBytes);
        if (evidence.MemoryBudgetExceeded) {
            QualityCaseResult memoryFailure = Failure(item, "worker-memory-budget-exceeded", stopwatch.ElapsedMilliseconds);
            AddProcessEvidence(memoryFailure, evidence, stopwatch.ElapsedMilliseconds);
            return memoryFailure;
        }
        try {
            QualityCaseResult? result = JsonSerializer.Deserialize<QualityCaseResult>(output, QualityJson.Options);
            if (result is null) {
                QualityCaseResult failure = Failure(item, "worker-empty-result", stopwatch.ElapsedMilliseconds);
                AddProcessEvidence(failure, evidence, stopwatch.ElapsedMilliseconds);
                return failure;
            }
            if (result.DurationMilliseconds == 0) result.DurationMilliseconds = stopwatch.ElapsedMilliseconds;
            AddProcessEvidence(result, evidence, stopwatch.ElapsedMilliseconds);
            if (process.ExitCode != 0 && string.Equals(result.Outcome, "passed", StringComparison.Ordinal)) {
                result.Outcome = "failed";
                result.FailureCode = "worker-nonzero-exit";
            }
            EnsureFailureEvidence(result);
            return result;
        } catch (JsonException) {
            QualityCaseResult failure = Failure(item, string.IsNullOrWhiteSpace(error) ? "worker-invalid-json" : "worker-error", stopwatch.ElapsedMilliseconds);
            AddProcessEvidence(failure, evidence, stopwatch.ElapsedMilliseconds);
            return failure;
        }
    }

    private static void AddProcessEvidence(QualityCaseResult result, WorkerProcessEvidence evidence, long wallClockMilliseconds) {
        result.WorkerWallClockMilliseconds = wallClockMilliseconds;
        result.PeakWorkingSetBytes = evidence.PeakWorkingSetBytes;
        result.WorkerCpuMilliseconds = evidence.CpuMilliseconds;
    }

    private static DateTime? TryGetStartTime(Process process) {
        try {
            return process.StartTime;
        } catch (InvalidOperationException) {
            return null;
        } catch (NotSupportedException) {
            return null;
        } catch (System.ComponentModel.Win32Exception) {
            return null;
        }
    }

    private static async Task<WorkerProcessEvidence> SampleProcessAsync(
        int processId,
        DateTime? expectedStartTime,
        long maxWorkingSetBytes) {
        long peakWorkingSetBytes = 0L;
        long cpuMilliseconds = 0L;
        bool memoryBudgetExceeded = false;
        bool canTerminateWorker = CanTerminateWorker(expectedStartTime);
        while (true) {
            try {
                using Process sample = Process.GetProcessById(processId);
                if (expectedStartTime.HasValue && sample.StartTime != expectedStartTime.Value) break;
                sample.Refresh();
                peakWorkingSetBytes = Math.Max(
                    peakWorkingSetBytes,
                    Math.Max(sample.WorkingSet64, TryGetPeakWorkingSet(sample)));
                cpuMilliseconds = Math.Max(cpuMilliseconds, (long)sample.TotalProcessorTime.TotalMilliseconds);
                if (IsMemoryBudgetExceeded(peakWorkingSetBytes, maxWorkingSetBytes)) {
                    memoryBudgetExceeded = true;
                    if (canTerminateWorker) {
                        TryKill(sample);
                        break;
                    }
                }
            } catch (ArgumentException) {
                break;
            } catch (InvalidOperationException) {
                break;
            } catch (NotSupportedException) {
                break;
            } catch (System.ComponentModel.Win32Exception) {
                break;
            }
            await Task.Delay(25).ConfigureAwait(false);
        }
        return new WorkerProcessEvidence(peakWorkingSetBytes, cpuMilliseconds, memoryBudgetExceeded);
    }

    private static WorkerProcessEvidence CompleteProcessEvidence(
        Process process,
        WorkerProcessEvidence sampled,
        long maxWorkingSetBytes) {
        long peakWorkingSetBytes = Math.Max(sampled.PeakWorkingSetBytes, TryGetPeakWorkingSet(process));
        long cpuMilliseconds = Math.Max(sampled.CpuMilliseconds, TryGetCpuMilliseconds(process));
        return new WorkerProcessEvidence(
            peakWorkingSetBytes,
            cpuMilliseconds,
            sampled.MemoryBudgetExceeded || IsMemoryBudgetExceeded(peakWorkingSetBytes, maxWorkingSetBytes));
    }

    private static long TryGetPeakWorkingSet(Process process) {
        try {
            return process.PeakWorkingSet64;
        } catch (InvalidOperationException) {
            return 0L;
        } catch (NotSupportedException) {
            return 0L;
        } catch (System.ComponentModel.Win32Exception) {
            return 0L;
        }
    }

    private static long TryGetCpuMilliseconds(Process process) {
        try {
            return (long)process.TotalProcessorTime.TotalMilliseconds;
        } catch (InvalidOperationException) {
            return 0L;
        } catch (NotSupportedException) {
            return 0L;
        } catch (System.ComponentModel.Win32Exception) {
            return 0L;
        }
    }

    private static bool IsMemoryBudgetExceeded(long peakWorkingSetBytes, long maxWorkingSetBytes) =>
        peakWorkingSetBytes > maxWorkingSetBytes;

    private static bool CanTerminateWorker(DateTime? verifiedStartTime) => verifiedStartTime.HasValue;

    private readonly record struct WorkerProcessEvidence(
        long PeakWorkingSetBytes,
        long CpuMilliseconds,
        bool MemoryBudgetExceeded);

    private static ProcessStartInfo CreateStartInfo(QualityRunOptions options, string caseId) {
        string entryAssembly = Assembly.GetEntryAssembly()?.Location
            ?? throw new InvalidOperationException("PDF quality corpus entry assembly path is unavailable.");
        var startInfo = new ProcessStartInfo {
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };
        if (string.Equals(Path.GetExtension(entryAssembly), ".dll", StringComparison.OrdinalIgnoreCase)) {
            startInfo.FileName = "dotnet";
            startInfo.ArgumentList.Add(entryAssembly);
        } else {
            startInfo.FileName = entryAssembly;
        }
        startInfo.ArgumentList.Add("probe");
        Add(startInfo, "--manifest", options.ManifestPath);
        Add(startInfo, "--root", options.RootDirectory);
        Add(startInfo, "--case-id", caseId);
        Add(startInfo, "--max-file-bytes", options.MaxFileBytes.ToString(System.Globalization.CultureInfo.InvariantCulture));
        Add(startInfo, "--max-render-pages", options.MaxRenderPages.ToString(System.Globalization.CultureInfo.InvariantCulture));
        return startInfo;
    }

    private static void Add(ProcessStartInfo startInfo, string name, string value) {
        startInfo.ArgumentList.Add(name);
        startInfo.ArgumentList.Add(value);
    }

    private static void TryKill(Process process) {
        try {
            if (!process.HasExited) process.Kill(entireProcessTree: true);
        } catch (InvalidOperationException) {
        } catch (System.ComponentModel.Win32Exception) {
        }
    }

    private static QualityCaseResult Failure(QualityCase item, string code, long duration) {
        var result = new QualityCaseResult {
            Id = item.Id,
            SourceId = item.Source,
            SourcePath = item.SourcePath,
            Sha256 = item.Sha256,
            ByteLength = item.ByteLength,
            Features = item.Features,
            Outcome = "failed",
            FailureCode = code,
            DurationMilliseconds = duration
        };
        EnsureFailureEvidence(result);
        return result;
    }

    private static void EnsureFailureEvidence(QualityCaseResult result) {
        if (string.Equals(result.Outcome, "passed", StringComparison.Ordinal)) return;
        string failureCode = string.IsNullOrWhiteSpace(result.FailureCode)
            ? "worker-reported-failure"
            : result.FailureCode;
        if (result.Checks.All(check => check.Succeeded)) {
            result.Checks = result.Checks.Concat(new[] {
                new QualityCheckResult {
                    Name = "coordinator.case-completed",
                    Succeeded = false,
                    Message = failureCode
                }
            }).ToArray();
        }
        if (result.Expectations.All(expectation => expectation.Succeeded)) {
            result.Expectations = result.Expectations.Concat(new[] {
                new QualityExpectationResult {
                    Name = "coordinator.case-outcome",
                    Succeeded = false,
                    Expected = "passed",
                    Actual = failureCode
                }
            }).ToArray();
        }
    }

    internal static void VerifyFailureScoringContract() {
        var item = new QualityCase { Id = "failure-contract", Source = "contract", ByteLength = 1L };
        QualityCaseResult result = Failure(item, "worker-contract-failure", 1L);
        QualityTotals totals = BuildTotals(new[] { result }, TimeSpan.FromMilliseconds(1));
        if (result.OperationalScore != 0D || result.ExpectationScore != 0D ||
            totals.OperationalScore != 0D || totals.ExpectationScore != 0D) {
            throw new InvalidOperationException("Fatal worker outcomes must fail both operational and expectation scores.");
        }
        if (!IsMemoryBudgetExceeded(2L, 1L) || IsMemoryBudgetExceeded(1L, 1L)) {
            throw new InvalidOperationException("Worker memory budget boundaries are not fail-closed.");
        }
        if (CanTerminateWorker(null) || !CanTerminateWorker(DateTime.UtcNow)) {
            throw new InvalidOperationException("Active worker termination requires verified process identity.");
        }
    }

    private static QualityTotals BuildTotals(IReadOnlyList<QualityCaseResult> results, TimeSpan duration) => new() {
        Cases = results.Count,
        Passed = results.Count(result => result.Outcome == "passed"),
        Failed = results.Count(result => result.Outcome == "failed"),
        TimedOut = results.Count(result => result.TimedOut),
        OperationalChecks = results.Sum(result => result.Checks.Count),
        OperationalChecksPassed = results.Sum(result => result.Checks.Count(check => check.Succeeded)),
        Expectations = results.Sum(result => result.Expectations.Count),
        ExpectationsPassed = results.Sum(result => result.Expectations.Count(expectation => expectation.Succeeded)),
        InputBytes = results.Sum(result => result.ByteLength),
        Pages = results.Sum(result => result.Metrics.PageCount),
        DurationMilliseconds = (long)duration.TotalMilliseconds,
        PeakWorkingSetBytes = results.Count == 0 ? 0L : results.Max(result => result.PeakWorkingSetBytes)
    };
}
