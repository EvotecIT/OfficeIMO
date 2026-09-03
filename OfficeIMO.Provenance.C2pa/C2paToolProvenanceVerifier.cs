using System.ComponentModel;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Provenance;

namespace OfficeIMO.Provenance.C2pa;

/// <summary>
/// Provides optional C2PA content-binding, signature, and trust verification through the official
/// <c>c2patool</c> command-line application. The executable is supplied by the host and is not bundled.
/// </summary>
public sealed class C2paToolProvenanceVerifier : IOfficeProvenanceVerifier {
    private static readonly string[] NonObjectReportFinding = { "c2patool returned a non-object JSON report." };
    private static readonly string[] MalformedActiveManifestFinding = { "c2patool returned malformed active_manifest data." };
    private static readonly string[] DuplicateCriticalReportFieldFinding = { "c2patool returned duplicate security-critical report fields." };
    private static readonly HashSet<string> SuccessfulValidationCodes = new(StringComparer.Ordinal) {
        "claimSignature.validated",
        "claimSignature.insideValidity",
        "signingCredential.trusted",
        "signingCredential.ocsp.notRevoked",
        "timeStamp.trusted",
        "timeStamp.validated",
        "assertion.hashedURI.match",
        "assertion.dataHash.match",
        "assertion.bmffHash.match",
        "assertion.accessible",
        "assertion.boxesHash.match",
        "assertion.collectionHash.match",
        "ingredient.manifest.validated",
        "ingredient.claimSignature.validated"
    };
    private readonly IC2paToolProcessRunner _runner;

    /// <summary>Creates a verifier for an installed or explicitly downloaded c2patool executable.</summary>
    public C2paToolProvenanceVerifier(string executablePath)
        : this(executablePath, new C2paToolProcessRunner()) { }

    internal C2paToolProvenanceVerifier(string executablePath, IC2paToolProcessRunner runner) {
        if (string.IsNullOrWhiteSpace(executablePath)) throw new ArgumentException("A c2patool executable path or command name is required.", nameof(executablePath));
        ExecutablePath = File.Exists(executablePath) ? Path.GetFullPath(executablePath) : executablePath;
        _runner = runner ?? throw new ArgumentNullException(nameof(runner));
    }

    /// <summary>Gets the executable path or command name used for verification.</summary>
    public string ExecutablePath { get; }

    /// <inheritdoc />
    public string Name => "c2patool";

    /// <inheritdoc />
    public OfficeProvenanceVerificationResult Verify(
        string filePath,
        OfficeProvenanceVerificationOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("An asset path is required.", nameof(filePath));
        string fullPath = Path.GetFullPath(filePath);
        if (!File.Exists(fullPath)) throw new FileNotFoundException("The asset to verify was not found.", fullPath);
        options ??= new OfficeProvenanceVerificationOptions();
        Validate(options);

        string settingsPath = Path.Combine(Path.GetTempPath(), ".officeimo-c2pa-" + Guid.NewGuid().ToString("N") + ".json");
        try {
            cancellationToken.ThrowIfCancellationRequested();
            File.WriteAllText(settingsPath, CreateSettings(options.AllowNetworkAccess), new UTF8Encoding(false));
            string workingDirectory = Path.GetDirectoryName(fullPath) ?? Directory.GetCurrentDirectory();
            var request = new C2paToolProcessRequest(
                ExecutablePath,
                BuildArguments(fullPath, settingsPath, workingDirectory, options),
                workingDirectory,
                options.Timeout,
                options.MaxReportBytes);
            C2paToolProcessResult processResult;
            try {
                processResult = _runner.Run(request, cancellationToken);
            } catch (Win32Exception exception) {
                return Result(OfficeProvenanceVerificationStatus.ProviderUnavailable, new[] { exception.Message }, null, options);
            } catch (TimeoutException exception) {
                return Result(OfficeProvenanceVerificationStatus.Error, new[] { exception.Message }, null, options);
            } catch (InvalidDataException exception) {
                return Result(OfficeProvenanceVerificationStatus.Error, new[] { exception.Message }, null, options);
            } catch (IOException exception) {
                return Result(OfficeProvenanceVerificationStatus.Error, new[] { exception.Message }, null, options);
            } catch (InvalidOperationException exception) {
                return Result(OfficeProvenanceVerificationStatus.Error, new[] { exception.Message }, null, options);
            }
            return Interpret(processResult, options);
        } finally {
            try { if (File.Exists(settingsPath)) File.Delete(settingsPath); } catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    private static OfficeProvenanceVerificationResult Interpret(C2paToolProcessResult process, OfficeProvenanceVerificationOptions options) {
        if (string.IsNullOrWhiteSpace(process.StandardOutput)) {
            string message = string.IsNullOrWhiteSpace(process.StandardError)
                ? $"c2patool exited with code {process.ExitCode} without a JSON report."
                : process.StandardError.Trim();
            return Result(OfficeProvenanceVerificationStatus.Error, new[] { message }, null, options);
        }
        try {
            using JsonDocument document = JsonDocument.Parse(process.StandardOutput, new JsonDocumentOptions {
                AllowTrailingCommas = false,
                CommentHandling = JsonCommentHandling.Disallow,
                MaxDepth = 128
            });
            if (document.RootElement.ValueKind != JsonValueKind.Object) {
                return Result(OfficeProvenanceVerificationStatus.Error,
                    NonObjectReportFinding, process.StandardOutput, options);
            }
            if (!TryGetUniqueProperty(document.RootElement, "active_manifest", out JsonElement activeManifestElement, out bool hasActiveManifest) ||
                !TryGetUniqueProperty(document.RootElement, "validation_status", out JsonElement validationStatus, out bool hasValidationStatus)) {
                return Result(OfficeProvenanceVerificationStatus.Error,
                    DuplicateCriticalReportFieldFinding, process.StandardOutput, options);
            }
            if (!hasActiveManifest) {
                return Result(OfficeProvenanceVerificationStatus.Error,
                    MalformedActiveManifestFinding, process.StandardOutput, options);
            }
            string? activeManifest = null;
            if (hasActiveManifest) {
                if (activeManifestElement.ValueKind == JsonValueKind.String) {
                    activeManifest = activeManifestElement.GetString();
                    if (string.IsNullOrWhiteSpace(activeManifest)) {
                        return Result(OfficeProvenanceVerificationStatus.Error,
                            MalformedActiveManifestFinding, process.StandardOutput, options);
                    }
                }
                else if (activeManifestElement.ValueKind != JsonValueKind.Null) {
                    return Result(OfficeProvenanceVerificationStatus.Error,
                        MalformedActiveManifestFinding, process.StandardOutput, options);
                }
            }
            var findings = new List<string>();
            var findingSet = new HashSet<string>(StringComparer.Ordinal);
            if (hasValidationStatus &&
                !TryCollectValidationFindings(validationStatus, findings, findingSet)) {
                findings.Add("c2patool returned malformed validation_status data.");
                return Result(OfficeProvenanceVerificationStatus.Error, findings, process.StandardOutput, options);
            }
            if (string.IsNullOrWhiteSpace(activeManifest)) {
                if (findings.Count > 0) {
                    bool onlyNoManifestTrustFailures = findings.All(IsTrustFinding);
                    return Result(
                        onlyNoManifestTrustFailures ? OfficeProvenanceVerificationStatus.Untrusted : OfficeProvenanceVerificationStatus.Invalid,
                        findings,
                        process.StandardOutput,
                        options);
                }
                if (process.ExitCode != 0 && findings.Count == 0) {
                    findings.Add(string.IsNullOrWhiteSpace(process.StandardError)
                        ? $"c2patool exited with code {process.ExitCode}."
                        : process.StandardError.Trim());
                    return Result(OfficeProvenanceVerificationStatus.Error, findings, process.StandardOutput, options);
                }
                return Result(OfficeProvenanceVerificationStatus.NotPresent, findings, process.StandardOutput, options);
            }
            if (findings.Count == 0) {
                if (process.ExitCode != 0) findings.Add($"c2patool exited with code {process.ExitCode} after producing a manifest report.");
                return Result(process.ExitCode == 0 ? OfficeProvenanceVerificationStatus.Valid : OfficeProvenanceVerificationStatus.Error,
                    findings, process.StandardOutput, options);
            }
            bool onlyTrustFailures = findings.All(IsTrustFinding);
            return Result(
                onlyTrustFailures ? OfficeProvenanceVerificationStatus.Untrusted : OfficeProvenanceVerificationStatus.Invalid,
                findings,
                process.StandardOutput,
                options);
        } catch (JsonException exception) {
            return Result(OfficeProvenanceVerificationStatus.Error,
                new[] { "c2patool returned malformed JSON: " + exception.Message },
                process.StandardOutput,
                options);
        }
    }

    private static bool TryCollectValidationFindings(JsonElement validationStatus, List<string> findings, HashSet<string> findingSet) {
        if (validationStatus.ValueKind != JsonValueKind.Array) return false;
        foreach (JsonElement status in validationStatus.EnumerateArray()) {
            if (status.ValueKind != JsonValueKind.Object) return false;
            if (!TryGetUniqueProperty(status, "code", out JsonElement codeElement, out bool hasCode) ||
                !TryGetUniqueProperty(status, "explanation", out JsonElement explanationElement, out bool hasExplanation) ||
                !TryGetUniqueProperty(status, "success", out JsonElement successElement, out bool hasSuccess)) return false;
            if (!hasCode || codeElement.ValueKind != JsonValueKind.String ||
                hasExplanation && explanationElement.ValueKind != JsonValueKind.String ||
                hasSuccess && successElement.ValueKind is not JsonValueKind.True and not JsonValueKind.False) return false;
            string? code = hasCode ? codeElement.GetString() : null;
            if (string.IsNullOrWhiteSpace(code)) return false;
            string? explanation = hasExplanation ? explanationElement.GetString() : null;
            bool? explicitSuccess = hasSuccess ? successElement.GetBoolean() : null;
            bool codeIndicatesSuccess = SuccessfulValidationCodes.Contains(code!);
            if (explicitSuccess.HasValue && explicitSuccess.Value != codeIndicatesSuccess) return false;
            if (codeIndicatesSuccess) continue;
            string finding = string.IsNullOrWhiteSpace(code) ? "unknown validation failure" : code!;
            if (!string.IsNullOrWhiteSpace(explanation)) finding += ": " + explanation;
            if (findingSet.Add(finding)) findings.Add(finding);
        }
        return true;
    }

    private static bool TryGetUniqueProperty(
        JsonElement element,
        string propertyName,
        out JsonElement value,
        out bool found) {
        value = default;
        found = false;
        foreach (JsonProperty property in element.EnumerateObject()) {
            if (!string.Equals(property.Name, propertyName, StringComparison.Ordinal)) continue;
            if (found) return false;
            value = property.Value;
            found = true;
        }
        return true;
    }

    private static bool IsTrustFinding(string finding) {
        int separator = finding.IndexOf(':');
        string code = separator < 0 ? finding : finding.Substring(0, separator);
        return code.StartsWith("signingCredential.", StringComparison.Ordinal) ||
            code == "timeStamp.untrusted" ||
            code == "timeStamp.outsideValidity" ||
            code == "cawg.ica.untrusted_issuer";
    }

    private static OfficeProvenanceVerificationResult Result(
        OfficeProvenanceVerificationStatus status,
        IReadOnlyList<string> findings,
        string? report,
        OfficeProvenanceVerificationOptions options) =>
        new(status, "c2patool", findings, options.IncludeRawReport ? report : null);

    private static ReadOnlyCollection<string> BuildArguments(
        string assetPath,
        string settingsPath,
        string workingDirectory,
        OfficeProvenanceVerificationOptions options) {
        var arguments = new List<string> { assetPath, "--settings", settingsPath };
        if (options.TrustAnchorsPath != null || options.AllowedListPath != null || options.TrustConfigurationPath != null) {
            arguments.Add("trust");
            AddTrustArgument(arguments, "--trust_anchors", options.TrustAnchorsPath, workingDirectory, options.AllowNetworkAccess);
            AddTrustArgument(arguments, "--allowed_list", options.AllowedListPath, workingDirectory, options.AllowNetworkAccess);
            AddTrustArgument(arguments, "--trust_config", options.TrustConfigurationPath, workingDirectory, options.AllowNetworkAccess);
        }
        return arguments.AsReadOnly();
    }

    private static void AddTrustArgument(List<string> arguments, string name, string? value, string workingDirectory, bool allowNetwork) {
        if (value == null) return;
        string argumentValue = value;
        if (Uri.TryCreate(value, UriKind.Absolute, out Uri? uri) && (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps)) {
            if (!allowNetwork) throw new ArgumentException($"Remote trust material for {name} requires AllowNetworkAccess.");
        } else {
            string fullPath = Path.GetFullPath(value);
            if (!File.Exists(fullPath)) throw new FileNotFoundException($"The trust material for {name} was not found.", value);
            argumentValue = GetRelativePath(workingDirectory, fullPath);
        }
        arguments.Add(name);
        arguments.Add(argumentValue);
    }

    internal static string GetRelativePath(string directoryPath, string filePath) {
        string directory = Path.GetFullPath(directoryPath);
        if (!directory.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)) directory += Path.DirectorySeparatorChar;
        var directoryUri = new Uri(directory);
        var fileUri = new Uri(Path.GetFullPath(filePath));
        Uri relativeUri = directoryUri.MakeRelativeUri(fileUri);
        if (relativeUri.IsAbsoluteUri) return fileUri.LocalPath;
        return Uri.UnescapeDataString(relativeUri.ToString())
            .Replace('/', Path.DirectorySeparatorChar);
    }

    private static string CreateSettings(bool allowNetwork) =>
        "{\"version\":1,\"verify\":{\"remote_manifest_fetch\":" + (allowNetwork ? "true" : "false") + ",\"ocsp_fetch\":" + (allowNetwork ? "true" : "false") + "}}";

    private static void Validate(OfficeProvenanceVerificationOptions options) {
        if (options.Timeout <= TimeSpan.Zero || options.Timeout > TimeSpan.FromMinutes(10)) {
            throw new ArgumentOutOfRangeException(nameof(options), "Timeout must be between zero and ten minutes.");
        }
        if (options.MaxReportBytes <= 0 || options.MaxReportBytes > int.MaxValue) {
            throw new ArgumentOutOfRangeException(nameof(options), "MaxReportBytes must be between one and Int32.MaxValue.");
        }
    }
}

internal interface IC2paToolProcessRunner {
    C2paToolProcessResult Run(C2paToolProcessRequest request, CancellationToken cancellationToken = default);
}

internal sealed class C2paToolProcessRequest {
    internal C2paToolProcessRequest(string executablePath, IReadOnlyList<string> arguments, string workingDirectory, TimeSpan timeout, long maximumOutputBytes) {
        ExecutablePath = executablePath;
        Arguments = arguments;
        WorkingDirectory = workingDirectory;
        Timeout = timeout;
        MaximumOutputBytes = maximumOutputBytes;
    }
    internal string ExecutablePath { get; }
    internal IReadOnlyList<string> Arguments { get; }
    internal string WorkingDirectory { get; }
    internal TimeSpan Timeout { get; }
    internal long MaximumOutputBytes { get; }
}

internal sealed class C2paToolProcessResult {
    internal C2paToolProcessResult(int exitCode, string standardOutput, string standardError) {
        ExitCode = exitCode;
        StandardOutput = standardOutput;
        StandardError = standardError;
    }
    internal int ExitCode { get; }
    internal string StandardOutput { get; }
    internal string StandardError { get; }
}

internal sealed class C2paToolProcessRunner : IC2paToolProcessRunner {
    private static readonly char[] ProcessSnapshotLineSeparators = { '\r', '\n' };
    private readonly bool _useExternalUnixSessionLauncher;

    internal C2paToolProcessRunner(bool useExternalUnixSessionLauncher = true) {
        _useExternalUnixSessionLauncher = useExternalUnixSessionLauncher;
    }

    public C2paToolProcessResult Run(
        C2paToolProcessRequest request,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        string targetExecutable = ResolveUnixExecutable(request.ExecutablePath, request.WorkingDirectory);
        string executable = targetExecutable;
        string arguments = string.Join(" ", request.Arguments.Select(QuoteArgument));
        string? sessionLauncher = _useExternalUnixSessionLauncher ? FindUnixSessionLauncher() : null;
        bool ownsUnixProcessGroup = sessionLauncher != null;
        if (Environment.OSVersion.Platform != PlatformID.Win32NT && sessionLauncher == null) {
            throw new Win32Exception(2,
                "c2patool execution on Unix requires a setsid executable so child processes can be contained safely.");
        }
        if (sessionLauncher != null) {
            executable = sessionLauncher;
            arguments = QuoteArgument(targetExecutable) + (arguments.Length == 0 ? string.Empty : " " + arguments);
        }
        var startInfo = new ProcessStartInfo {
            FileName = executable,
            Arguments = arguments,
            WorkingDirectory = request.WorkingDirectory,
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        using var process = new Process { StartInfo = startInfo };
        process.Start();
        using C2paToolProcessContainment containment = C2paToolProcessContainment.Create(process, ownsUnixProcessGroup);
        Task<string> stdout = ReadBoundedAsync(process.StandardOutput.BaseStream, request.MaximumOutputBytes, "standard output");
        Task<string> stderr = ReadBoundedAsync(process.StandardError.BaseStream, request.MaximumOutputBytes, "standard error");
        Stopwatch timer = Stopwatch.StartNew();
        while (true) {
            if (process.WaitForExit(50)) break;
            ThrowIfCancellationRequested(process, containment, cancellationToken);
            if (stdout.IsFaulted || stderr.IsFaulted) {
                Terminate(process, containment);
                throw stdout.Exception?.GetBaseException() ?? stderr.Exception?.GetBaseException() ?? new InvalidDataException("c2patool output failed.");
            }
            if (timer.Elapsed > request.Timeout) {
                Terminate(process, containment);
                throw new TimeoutException($"c2patool exceeded the configured timeout of {request.Timeout}.");
            }
        }
        try {
            var outputTasks = new Task[] { stdout, stderr };
            while (true) {
                TimeSpan remaining = request.Timeout - timer.Elapsed;
                if (remaining <= TimeSpan.Zero) {
                    Terminate(process, containment);
                    throw new TimeoutException($"c2patool exceeded the configured timeout of {request.Timeout}.");
                }
                int waitMilliseconds = (int)Math.Min(50D, Math.Ceiling(remaining.TotalMilliseconds));
                if (Task.WaitAll(outputTasks, waitMilliseconds, CancellationToken.None)) break;
                ThrowIfCancellationRequested(process, containment, cancellationToken);
            }
        } catch (AggregateException exception) {
            throw exception.GetBaseException();
        }
        cancellationToken.ThrowIfCancellationRequested();
        return new C2paToolProcessResult(process.ExitCode, stdout.Result, stderr.Result);
    }

    private static void ThrowIfCancellationRequested(
        Process process,
        C2paToolProcessContainment containment,
        CancellationToken cancellationToken) {
        if (!cancellationToken.IsCancellationRequested) return;
        Terminate(process, containment);
        cancellationToken.ThrowIfCancellationRequested();
    }

    internal static Task<string> ReadBoundedAsync(Stream stream, long maximumBytes, string streamName) => Task.Run(() => {
        try {
            using var reader = new StreamReader(
                stream,
                new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true),
                detectEncodingFromByteOrderMarks: true,
                bufferSize: 4096,
                leaveOpen: true);
            var builder = new StringBuilder();
            char[] buffer = new char[4096];
            long bytes = 0;
            while (true) {
                int read = reader.Read(buffer, 0, buffer.Length);
                if (read <= 0) break;
                bytes += Encoding.UTF8.GetByteCount(buffer, 0, read);
                if (bytes > maximumBytes) throw new InvalidDataException($"c2patool {streamName} exceeds the configured limit of {maximumBytes} bytes.");
                builder.Append(buffer, 0, read);
            }
            return builder.ToString();
        } catch (DecoderFallbackException exception) {
            throw new InvalidDataException($"c2patool {streamName} is not valid UTF-8.", exception);
        }
    });

    private static void Terminate(Process process, C2paToolProcessContainment containment) {
        try {
            containment.Terminate();
            if (!TryKillEntireProcessTree(process) && !process.HasExited) {
                process.Kill();
                process.WaitForExit(1000);
            }
        } catch (InvalidOperationException) { }
        catch (Win32Exception) { }
        finally {
            try { process.StandardOutput.Dispose(); } catch (InvalidOperationException) { }
            try { process.StandardError.Dispose(); } catch (InvalidOperationException) { }
        }
    }

    private sealed class C2paToolProcessContainment : IDisposable {
        private const uint JobObjectLimitKillOnJobClose = 0x00002000;
        private IntPtr _job;
        private readonly int _unixProcessGroupId;

        private C2paToolProcessContainment(IntPtr job, int unixProcessGroupId = 0) {
            _job = job;
            _unixProcessGroupId = unixProcessGroupId;
        }

        internal static C2paToolProcessContainment Create(Process process, bool ownsUnixProcessGroup) {
            if (Environment.OSVersion.Platform != PlatformID.Win32NT) {
                return new C2paToolProcessContainment(IntPtr.Zero, ownsUnixProcessGroup ? process.Id : 0);
            }
            IntPtr job = CreateJobObject(IntPtr.Zero, null);
            if (job == IntPtr.Zero) return new C2paToolProcessContainment(IntPtr.Zero);
            var information = new JobObjectExtendedLimitInformation {
                BasicLimitInformation = new JobObjectBasicLimitInformation { LimitFlags = JobObjectLimitKillOnJobClose }
            };
            int length = Marshal.SizeOf<JobObjectExtendedLimitInformation>();
            if (!SetInformationJobObject(job, 9, ref information, length) || !AssignProcessToJobObject(job, process.Handle)) {
                CloseHandle(job);
                return new C2paToolProcessContainment(IntPtr.Zero);
            }
            return new C2paToolProcessContainment(job);
        }

        internal void Terminate() => Dispose();

        public void Dispose() {
            if (_unixProcessGroupId > 0) _ = KillUnixProcessGroup(-_unixProcessGroupId, 9);
            IntPtr job = _job;
            if (job == IntPtr.Zero) return;
            _job = IntPtr.Zero;
            CloseHandle(job);
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct IoCounters {
            internal ulong ReadOperationCount;
            internal ulong WriteOperationCount;
            internal ulong OtherOperationCount;
            internal ulong ReadTransferCount;
            internal ulong WriteTransferCount;
            internal ulong OtherTransferCount;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct JobObjectBasicLimitInformation {
            internal long PerProcessUserTimeLimit;
            internal long PerJobUserTimeLimit;
            internal uint LimitFlags;
            internal UIntPtr MinimumWorkingSetSize;
            internal UIntPtr MaximumWorkingSetSize;
            internal uint ActiveProcessLimit;
            internal UIntPtr Affinity;
            internal uint PriorityClass;
            internal uint SchedulingClass;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct JobObjectExtendedLimitInformation {
            internal JobObjectBasicLimitInformation BasicLimitInformation;
            internal IoCounters IoInfo;
            internal UIntPtr ProcessMemoryLimit;
            internal UIntPtr JobMemoryLimit;
            internal UIntPtr PeakProcessMemoryUsed;
            internal UIntPtr PeakJobMemoryUsed;
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr CreateJobObject(IntPtr securityAttributes, string? name);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool SetInformationJobObject(IntPtr job, int informationClass, ref JobObjectExtendedLimitInformation information, int informationLength);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool AssignProcessToJobObject(IntPtr job, IntPtr process);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CloseHandle(IntPtr handle);

        [DllImport("libc", EntryPoint = "kill", SetLastError = true)]
        private static extern int KillUnixProcessGroup(int processId, int signal);
    }

    private static string? FindUnixSessionLauncher() {
        if (Environment.OSVersion.Platform == PlatformID.Win32NT) return null;
        foreach (string path in new[] {
            "/usr/bin/setsid",
            "/bin/setsid",
            "/usr/local/bin/setsid",
            "/opt/homebrew/opt/util-linux/bin/setsid",
            "/usr/local/opt/util-linux/bin/setsid"
        }) {
            if (File.Exists(path) && IsUnixExecutable(path)) return path;
        }
        string searchPath = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
        foreach (string directory in searchPath.Split(Path.PathSeparator)) {
            if (string.IsNullOrWhiteSpace(directory)) continue;
            string candidate;
            try { candidate = Path.GetFullPath(Path.Combine(directory.Trim(), "setsid")); }
            catch (Exception exception) when (exception is ArgumentException || exception is NotSupportedException || exception is PathTooLongException) { continue; }
            if (File.Exists(candidate) && IsUnixExecutable(candidate)) return candidate;
        }
        return null;
    }

    private static string ResolveUnixExecutable(string configuredPath, string workingDirectory) {
        if (Environment.OSVersion.Platform == PlatformID.Win32NT) return configuredPath;
        bool containsSeparator = configuredPath.Contains(Path.DirectorySeparatorChar.ToString()) ||
            configuredPath.Contains(Path.AltDirectorySeparatorChar.ToString());
        if (containsSeparator || Path.IsPathRooted(configuredPath)) {
            string candidate = Path.IsPathRooted(configuredPath)
                ? Path.GetFullPath(configuredPath)
                : Path.GetFullPath(Path.Combine(workingDirectory, configuredPath));
            if (File.Exists(candidate) && IsUnixExecutable(candidate)) return candidate;
            throw new Win32Exception(2, $"The configured c2patool executable '{configuredPath}' was not found or is not executable.");
        }

        string path = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
        foreach (string directory in path.Split(Path.PathSeparator)) {
            if (string.IsNullOrWhiteSpace(directory)) continue;
            string candidate;
            try { candidate = Path.GetFullPath(Path.Combine(directory.Trim(), configuredPath)); }
            catch (Exception exception) when (exception is ArgumentException || exception is NotSupportedException || exception is PathTooLongException) { continue; }
            if (File.Exists(candidate) && IsUnixExecutable(candidate)) return candidate;
        }
        throw new Win32Exception(2, $"The configured c2patool executable '{configuredPath}' was not found or is not executable.");
    }

    private static bool IsUnixExecutable(string path) => UnixAccess(path, 1) == 0;

    [DllImport("libc", EntryPoint = "access", SetLastError = true, CharSet = CharSet.Ansi,
        BestFitMapping = false, ThrowOnUnmappableChar = true)]
    private static extern int UnixAccess([MarshalAs(UnmanagedType.LPStr)] string path, int mode);

    private static bool TryKillEntireProcessTree(Process process) {
        System.Reflection.MethodInfo? method = typeof(Process).GetMethod("Kill", new[] { typeof(bool) });
        if (method != null) {
            try {
                method.Invoke(process, new object[] { true });
                return true;
            } catch (System.Reflection.TargetInvocationException exception) when (
                exception.InnerException is InvalidOperationException || exception.InnerException is Win32Exception) { }
        }
        if (Environment.OSVersion.Platform == PlatformID.Win32NT) return TryKillWithTaskKill(process);
        return TryKillUnixProcessTree(process);
    }

    private static bool TryKillWithTaskKill(Process process) {
        try {
            using Process? killer = Process.Start(new ProcessStartInfo {
                FileName = "taskkill.exe",
                Arguments = $"/PID {process.Id} /T /F",
                UseShellExecute = false,
                CreateNoWindow = true
            });
            if (killer == null || !killer.WaitForExit(2000)) return false;
            return killer.ExitCode == 0 || process.HasExited;
        } catch (InvalidOperationException) { return false; }
        catch (Win32Exception) { return false; }
    }

    private static bool TryKillUnixProcessTree(Process process) {
        try {
            var children = new Dictionary<int, List<int>>();
            using Process? snapshot = Process.Start(new ProcessStartInfo {
                FileName = "/bin/ps",
                Arguments = "-e -o pid= -o ppid=",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            });
            if (snapshot == null) return false;
            string output = ReadProcessSnapshot(snapshot.StandardOutput, 4 * 1024 * 1024);
            if (!snapshot.WaitForExit(2000) || snapshot.ExitCode != 0) return false;
            foreach (string line in output.Split(ProcessSnapshotLineSeparators, StringSplitOptions.RemoveEmptyEntries)) {
                string[] fields = line.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
                if (fields.Length != 2 || !int.TryParse(fields[0], out int pid) || !int.TryParse(fields[1], out int parent)) continue;
                if (!children.TryGetValue(parent, out List<int>? list)) children.Add(parent, list = new List<int>());
                list.Add(pid);
            }
            var descendants = new List<int>();
            var pending = new Stack<int>();
            pending.Push(process.Id);
            while (pending.Count > 0) {
                int parent = pending.Pop();
                if (!children.TryGetValue(parent, out List<int>? direct)) continue;
                foreach (int child in direct) { descendants.Add(child); pending.Push(child); }
            }
            process.Kill();
            for (int index = descendants.Count - 1; index >= 0; index--) {
                try { using Process child = Process.GetProcessById(descendants[index]); child.Kill(); }
                catch (ArgumentException) { }
                catch (InvalidOperationException) { }
                catch (Win32Exception) { }
            }
            return true;
        } catch (InvalidOperationException) { return false; }
        catch (Win32Exception) { return false; }
        catch (InvalidDataException) { return false; }
    }

    private static string ReadProcessSnapshot(TextReader reader, int maximumCharacters) {
        var builder = new StringBuilder();
        char[] buffer = new char[4096];
        while (true) {
            int read = reader.Read(buffer, 0, buffer.Length);
            if (read <= 0) return builder.ToString();
            if (builder.Length > maximumCharacters - read) throw new InvalidDataException("The process-tree snapshot exceeds its safety limit.");
            builder.Append(buffer, 0, read);
        }
    }

    private static string QuoteArgument(string value) {
        if (value.Length > 0 && value.All(character => !char.IsWhiteSpace(character) && character != '"')) return value;
        var builder = new StringBuilder("\"");
        int backslashes = 0;
        foreach (char character in value) {
            if (character == '\\') { backslashes++; continue; }
            if (character == '"') {
                builder.Append('\\', backslashes * 2 + 1).Append('"');
                backslashes = 0;
                continue;
            }
            builder.Append('\\', backslashes).Append(character);
            backslashes = 0;
        }
        builder.Append('\\', backslashes * 2).Append('"');
        return builder.ToString();
    }
}
