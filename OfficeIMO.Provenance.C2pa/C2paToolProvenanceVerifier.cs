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
public sealed partial class C2paToolProvenanceVerifier : ICancellableOfficeProvenanceVerifier {
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
        OfficeProvenanceVerificationOptions? options = null) => Verify(
            filePath,
            options,
            CancellationToken.None);

    /// <inheritdoc />
    public OfficeProvenanceVerificationResult Verify(
        string filePath,
        OfficeProvenanceVerificationOptions? options,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("An asset path is required.", nameof(filePath));
        string fullPath = Path.GetFullPath(filePath);
        if (!File.Exists(fullPath)) throw new FileNotFoundException("The asset to verify was not found.", fullPath);
        options ??= new OfficeProvenanceVerificationOptions();
        Validate(options);
        var executionBudget = new C2paToolExecutionBudget(options.Timeout, cancellationToken);

        string settingsPath = Path.Combine(Path.GetTempPath(), ".officeimo-c2pa-" + Guid.NewGuid().ToString("N") + ".json");
        try {
            cancellationToken.ThrowIfCancellationRequested();
            File.WriteAllText(settingsPath, CreateSettings(options.AllowNetworkAccess), new UTF8Encoding(false));
            string workingDirectory = Path.GetDirectoryName(fullPath) ?? Directory.GetCurrentDirectory();
            try {
                var request = new C2paToolProcessRequest(
                    ExecutablePath,
                    BuildArguments(fullPath, settingsPath, workingDirectory, options),
                    workingDirectory,
                    executionBudget.GetRemainingTimeout(),
                    options.MaxReportBytes);
                C2paToolProcessResult processResult = _runner.Run(request, cancellationToken);
                return executionBudget.RunInterpretation(
                    interpretationToken => Interpret(processResult, options, interpretationToken));
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
        } finally {
            try { if (File.Exists(settingsPath)) File.Delete(settingsPath); } catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    private static OfficeProvenanceVerificationResult Interpret(
        C2paToolProcessResult process,
        OfficeProvenanceVerificationOptions options,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (string.IsNullOrWhiteSpace(process.StandardOutput)) {
            // c2patool 0.27 reports ordinary absence on stderr rather than as JSON.
            // Match only the observed domain response, never arbitrary failures containing these words.
            if (process.ExitCode == 1 && string.Equals(process.StandardError.Trim(), "Error: No claim found", StringComparison.Ordinal))
                return Result(OfficeProvenanceVerificationStatus.NotPresent, Array.Empty<string>(), null, options);
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
            cancellationToken.ThrowIfCancellationRequested();
            if (document.RootElement.ValueKind != JsonValueKind.Object) {
                return Result(OfficeProvenanceVerificationStatus.Error,
                    NonObjectReportFinding, process.StandardOutput, options);
            }
            if (!TryGetUniqueProperty(document.RootElement, "active_manifest", cancellationToken, out JsonElement activeManifestElement, out bool hasActiveManifest) ||
                !TryGetUniqueProperty(document.RootElement, "validation_status", cancellationToken, out JsonElement validationStatus, out bool hasValidationStatus)) {
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
                !TryCollectValidationFindings(validationStatus, findings, findingSet, cancellationToken)) {
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

    private static bool TryCollectValidationFindings(
        JsonElement validationStatus,
        List<string> findings,
        HashSet<string> findingSet,
        CancellationToken cancellationToken) {
        if (validationStatus.ValueKind != JsonValueKind.Array) return false;
        foreach (JsonElement status in validationStatus.EnumerateArray()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (status.ValueKind != JsonValueKind.Object) return false;
            if (!TryGetUniqueProperty(status, "code", cancellationToken, out JsonElement codeElement, out bool hasCode) ||
                !TryGetUniqueProperty(status, "explanation", cancellationToken, out JsonElement explanationElement, out bool hasExplanation) ||
                !TryGetUniqueProperty(status, "success", cancellationToken, out JsonElement successElement, out bool hasSuccess)) return false;
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
        CancellationToken cancellationToken,
        out JsonElement value,
        out bool found) {
        value = default;
        found = false;
        foreach (JsonProperty property in element.EnumerateObject()) {
            cancellationToken.ThrowIfCancellationRequested();
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
