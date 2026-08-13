using System.ComponentModel;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using OfficeIMO.Provenance;

namespace OfficeIMO.Security;

/// <summary>
/// Provides optional C2PA content-binding, signature, and trust verification through the official
/// <c>c2patool</c> command-line application. The executable is supplied by the host and is not bundled.
/// </summary>
public sealed class C2paToolProvenanceVerifier : IOfficeProvenanceVerifier {
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
    public OfficeProvenanceVerificationResult Verify(string filePath, OfficeProvenanceVerificationOptions? options = null) {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("An asset path is required.", nameof(filePath));
        string fullPath = Path.GetFullPath(filePath);
        if (!File.Exists(fullPath)) throw new FileNotFoundException("The asset to verify was not found.", fullPath);
        options ??= new OfficeProvenanceVerificationOptions();
        Validate(options);

        string settingsPath = Path.Combine(Path.GetTempPath(), ".officeimo-c2pa-" + Guid.NewGuid().ToString("N") + ".json");
        try {
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
                processResult = _runner.Run(request);
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
            string? activeManifest = document.RootElement.ValueKind == JsonValueKind.Object &&
                document.RootElement.TryGetProperty("active_manifest", out JsonElement activeManifestElement) &&
                activeManifestElement.ValueKind == JsonValueKind.String
                    ? activeManifestElement.GetString()
                    : null;
            var findings = new List<string>();
            var findingSet = new HashSet<string>(StringComparer.Ordinal);
            CollectValidationFindings(document.RootElement, findings, findingSet);
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

    private static void CollectValidationFindings(JsonElement element, List<string> findings, HashSet<string> findingSet) {
        if (element.ValueKind == JsonValueKind.Object) {
            foreach (JsonProperty property in element.EnumerateObject()) {
                if (property.NameEquals("validation_status") && property.Value.ValueKind == JsonValueKind.Array) {
                    foreach (JsonElement status in property.Value.EnumerateArray()) {
                        if (status.ValueKind != JsonValueKind.Object) continue;
                        string? code = status.TryGetProperty("code", out JsonElement codeElement) && codeElement.ValueKind == JsonValueKind.String
                            ? codeElement.GetString()
                            : null;
                        string? explanation = status.TryGetProperty("explanation", out JsonElement explanationElement) && explanationElement.ValueKind == JsonValueKind.String
                            ? explanationElement.GetString()
                            : null;
                        bool? explicitSuccess = status.TryGetProperty("success", out JsonElement successElement) &&
                            (successElement.ValueKind == JsonValueKind.True || successElement.ValueKind == JsonValueKind.False)
                            ? successElement.GetBoolean()
                            : null;
                        if (explicitSuccess == true || explicitSuccess != false && code != null && SuccessfulValidationCodes.Contains(code)) continue;
                        string finding = string.IsNullOrWhiteSpace(code) ? "unknown validation failure" : code!;
                        if (!string.IsNullOrWhiteSpace(explanation)) finding += ": " + explanation;
                        if (findingSet.Add(finding)) findings.Add(finding);
                    }
                } else {
                    CollectValidationFindings(property.Value, findings, findingSet);
                }
            }
        } else if (element.ValueKind == JsonValueKind.Array) {
            foreach (JsonElement item in element.EnumerateArray()) CollectValidationFindings(item, findings, findingSet);
        }
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
    C2paToolProcessResult Run(C2paToolProcessRequest request);
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

    public C2paToolProcessResult Run(C2paToolProcessRequest request) {
        var startInfo = new ProcessStartInfo {
            FileName = request.ExecutablePath,
            Arguments = string.Join(" ", request.Arguments.Select(QuoteArgument)),
            WorkingDirectory = request.WorkingDirectory,
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8
        };
        using var process = new Process { StartInfo = startInfo };
        process.Start();
        Task<string> stdout = ReadBoundedAsync(process.StandardOutput, request.MaximumOutputBytes, "standard output");
        Task<string> stderr = ReadBoundedAsync(process.StandardError, request.MaximumOutputBytes, "standard error");
        Stopwatch timer = Stopwatch.StartNew();
        while (!process.WaitForExit(50)) {
            if (stdout.IsFaulted || stderr.IsFaulted) {
                Terminate(process);
                throw stdout.Exception?.GetBaseException() ?? stderr.Exception?.GetBaseException() ?? new InvalidDataException("c2patool output failed.");
            }
            if (timer.Elapsed > request.Timeout) {
                Terminate(process);
                throw new TimeoutException($"c2patool exceeded the configured timeout of {request.Timeout}.");
            }
        }
        try {
            TimeSpan remaining = request.Timeout - timer.Elapsed;
            if (remaining <= TimeSpan.Zero || !Task.WaitAll(new Task[] { stdout, stderr }, remaining)) {
                Terminate(process);
                throw new TimeoutException($"c2patool exceeded the configured timeout of {request.Timeout}.");
            }
        } catch (AggregateException exception) {
            throw exception.GetBaseException();
        }
        return new C2paToolProcessResult(process.ExitCode, stdout.Result, stderr.Result);
    }

    private static Task<string> ReadBoundedAsync(TextReader reader, long maximumBytes, string streamName) => Task.Run(() => {
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
    });

    private static void Terminate(Process process) {
        try {
            if (!process.HasExited) {
                if (!TryKillEntireProcessTree(process)) process.Kill();
                process.WaitForExit(1000);
            }
        } catch (InvalidOperationException) { }
        catch (Win32Exception) { }
        finally {
            try { process.StandardOutput.Dispose(); } catch (InvalidOperationException) { }
            try { process.StandardError.Dispose(); } catch (InvalidOperationException) { }
        }
    }

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
