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
            string? activeManifest = FindStringProperty(document.RootElement, "active_manifest");
            var findings = new List<string>();
            CollectValidationFindings(document.RootElement, findings);
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

    private static string? FindStringProperty(JsonElement element, string name) {
        if (element.ValueKind == JsonValueKind.Object) {
            foreach (JsonProperty property in element.EnumerateObject()) {
                if (property.NameEquals(name) && property.Value.ValueKind == JsonValueKind.String) return property.Value.GetString();
                string? nested = FindStringProperty(property.Value, name);
                if (nested != null) return nested;
            }
        } else if (element.ValueKind == JsonValueKind.Array) {
            foreach (JsonElement item in element.EnumerateArray()) {
                string? nested = FindStringProperty(item, name);
                if (nested != null) return nested;
            }
        }
        return null;
    }

    private static void CollectValidationFindings(JsonElement element, List<string> findings) {
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
                        string finding = string.IsNullOrWhiteSpace(code) ? "unknown validation failure" : code!;
                        if (!string.IsNullOrWhiteSpace(explanation)) finding += ": " + explanation;
                        if (!findings.Contains(finding, StringComparer.Ordinal)) findings.Add(finding);
                    }
                } else {
                    CollectValidationFindings(property.Value, findings);
                }
            }
        } else if (element.ValueKind == JsonValueKind.Array) {
            foreach (JsonElement item in element.EnumerateArray()) CollectValidationFindings(item, findings);
        }
    }

    private static bool IsTrustFinding(string finding) {
        string normalized = finding.ToUpperInvariant();
        return normalized.Contains("UNTRUSTED") || normalized.Contains("TRUST");
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

    private static string GetRelativePath(string directoryPath, string filePath) {
        string directory = Path.GetFullPath(directoryPath);
        if (!directory.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)) directory += Path.DirectorySeparatorChar;
        var directoryUri = new Uri(directory);
        var fileUri = new Uri(Path.GetFullPath(filePath));
        return Uri.UnescapeDataString(directoryUri.MakeRelativeUri(fileUri).ToString())
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
                TryKill(process);
                throw stdout.Exception?.GetBaseException() ?? stderr.Exception?.GetBaseException() ?? new InvalidDataException("c2patool output failed.");
            }
            if (timer.Elapsed > request.Timeout) {
                TryKill(process);
                throw new TimeoutException($"c2patool exceeded the configured timeout of {request.Timeout}.");
            }
        }
        try {
            Task.WaitAll(stdout, stderr);
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

    private static void TryKill(Process process) {
        try { if (!process.HasExited) process.Kill(); } catch (InvalidOperationException) { }
        catch (Win32Exception) { }
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
