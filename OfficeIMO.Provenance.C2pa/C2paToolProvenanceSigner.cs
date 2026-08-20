using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.Json;
using OfficeIMO.Core.Internal;
using OfficeIMO.Provenance;

namespace OfficeIMO.Provenance.C2pa;

/// <summary>
/// Creates C2PA Content Credentials through the official <c>c2patool</c> command-line application.
/// Version 0.27.0 or newer, plus a production signer process, are supplied by the host and are not bundled.
/// </summary>
public sealed class C2paToolProvenanceSigner : IOfficeProvenanceSigner {
    private static readonly Version MinimumSupportedVersion = new Version(0, 27, 0);
    private static readonly string[] VersionArguments = { "--version" };
    private readonly IC2paToolProcessRunner _runner;

    /// <summary>Creates a production signer that delegates private-key operations to a separate signer process.</summary>
    public C2paToolProvenanceSigner(string executablePath, string signerPath)
        : this(executablePath, signerPath, useBuiltInTestCredentials: false, new C2paToolProcessRunner()) { }

    private C2paToolProvenanceSigner(string executablePath, bool useBuiltInTestCredentials)
        : this(executablePath, null, useBuiltInTestCredentials, new C2paToolProcessRunner()) { }

    internal C2paToolProvenanceSigner(
        string executablePath,
        string? signerPath,
        bool useBuiltInTestCredentials,
        IC2paToolProcessRunner runner) {
        if (string.IsNullOrWhiteSpace(executablePath)) {
            throw new ArgumentException("A c2patool executable path or command name is required.", nameof(executablePath));
        }
        if (!useBuiltInTestCredentials && string.IsNullOrWhiteSpace(signerPath)) {
            throw new ArgumentException("A production signer path or command name is required.", nameof(signerPath));
        }
        if (useBuiltInTestCredentials && !string.IsNullOrWhiteSpace(signerPath)) {
            throw new ArgumentException("A signer path cannot be combined with built-in test credentials.", nameof(signerPath));
        }

        ExecutablePath = NormalizeCommand(executablePath);
        SignerPath = string.IsNullOrWhiteSpace(signerPath) ? null : NormalizeCommand(signerPath!);
        UsesBuiltInTestCredentials = useBuiltInTestCredentials;
        _runner = runner ?? throw new ArgumentNullException(nameof(runner));
    }

    /// <summary>
    /// Creates a development-only signer that uses c2patool's public test credentials.
    /// Outputs created by this signer must not be represented as production credentials.
    /// </summary>
    public static C2paToolProvenanceSigner CreateWithBuiltInTestCredentials(string executablePath) =>
        new(executablePath, useBuiltInTestCredentials: true);

    /// <summary>Gets the executable path or command name used for signing.</summary>
    public string ExecutablePath { get; }
    /// <summary>Gets the external signer process path or command name used for production key operations.</summary>
    public string? SignerPath { get; }
    /// <summary>Gets whether c2patool's public development credentials are used.</summary>
    public bool UsesBuiltInTestCredentials { get; }
    /// <inheritdoc />
    public string Name => "c2patool";

    /// <inheritdoc />
    public OfficeProvenanceSigningResult Sign(
        OfficeProvenanceSigningRequest request,
        OfficeProvenanceSigningOptions? options = null) {
#if NET6_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(request);
#else
        if (request == null) throw new ArgumentNullException(nameof(request));
#endif
        options ??= new OfficeProvenanceSigningOptions();
        Validate(options);

        string inputPath = Path.GetFullPath(request.InputPath);
        string outputPath = Path.GetFullPath(request.OutputPath);
        string? parentPath = string.IsNullOrWhiteSpace(request.ParentPath)
            ? null
            : Path.GetFullPath(request.ParentPath!);
        if (!File.Exists(inputPath)) throw new FileNotFoundException("The asset to sign was not found.", inputPath);
        if (parentPath != null && !File.Exists(parentPath)) {
            throw new FileNotFoundException("The parent asset was not found.", parentPath);
        }
        if (PathsEqual(inputPath, outputPath)) {
            throw new ArgumentException("Signing requires a separate output path so the input cannot be partially replaced.", nameof(request));
        }
        if (parentPath != null && PathsEqual(parentPath, outputPath)) {
            throw new ArgumentException("The signed output cannot replace its parent asset.", nameof(request));
        }
        if (!string.Equals(Path.GetExtension(inputPath), Path.GetExtension(outputPath), StringComparison.OrdinalIgnoreCase)) {
            throw new ArgumentException("The input and output extensions must match. Format conversion is not part of signing.", nameof(request));
        }
        if (!options.ReplaceExistingOutput && File.Exists(outputPath)) {
            return Result(OfficeProvenanceSigningStatus.Rejected, "The destination already exists.", options);
        }

        string? creationType = ValidateIntent(request.Claim, parentPath);
        OfficeProvenanceSigningResult? versionFailure = ProbeVersion(options, Path.GetDirectoryName(inputPath) ?? Directory.GetCurrentDirectory());
        if (versionFailure != null) return versionFailure;

        OfficeFileCommit.EnsureTargetDirectory(outputPath);
        string manifestPath = Path.Combine(Path.GetTempPath(), ".officeimo-c2pa-manifest-" + Guid.NewGuid().ToString("N") + ".json");
        string stagingPath = OfficeFileCommit.CreateStagingPath(outputPath);
        try {
            File.WriteAllBytes(manifestPath, CreateManifest(request.Claim));

            var arguments = new List<string> {
                inputPath,
                "--manifest", manifestPath,
                "--output", stagingPath
            };
            if (parentPath != null) {
                arguments.Add("--parent");
                arguments.Add(parentPath);
            } else {
                arguments.Add("--create");
                arguments.Add(creationType!);
            }
            if (SignerPath != null) {
                arguments.Add("--signer-path");
                arguments.Add(SignerPath);
            }

            C2paToolProcessResult process;
            try {
                process = _runner.Run(new C2paToolProcessRequest(
                    ExecutablePath,
                    arguments.AsReadOnly(),
                    Path.GetDirectoryName(inputPath) ?? Directory.GetCurrentDirectory(),
                    options.Timeout,
                    options.MaxReportBytes));
            } catch (Win32Exception exception) {
                return Result(OfficeProvenanceSigningStatus.ProviderUnavailable, exception.Message, options);
            } catch (TimeoutException exception) {
                return Result(OfficeProvenanceSigningStatus.Error, exception.Message, options);
            } catch (InvalidDataException exception) {
                return Result(OfficeProvenanceSigningStatus.Error, exception.Message, options);
            } catch (IOException exception) {
                return Result(OfficeProvenanceSigningStatus.Error, exception.Message, options);
            } catch (InvalidOperationException exception) {
                return Result(OfficeProvenanceSigningStatus.Error, exception.Message, options);
            }

            string? rawReport = options.IncludeRawReport ? JoinOutput(process) : null;
            if (process.ExitCode != 0) {
                return Result(
                    OfficeProvenanceSigningStatus.Rejected,
                    ProviderMessage(process, $"c2patool exited with code {process.ExitCode}."),
                    options,
                    rawReport);
            }
            if (!File.Exists(stagingPath)) {
                return Result(OfficeProvenanceSigningStatus.Error, "c2patool reported success without creating an output asset.", options, rawReport);
            }

            OfficeProvenanceReport structuralReport;
            try {
                structuralReport = OfficeProvenanceInspector.InspectFile(stagingPath);
            } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or InvalidDataException) {
                return Result(OfficeProvenanceSigningStatus.Error, "The signed output could not be inspected: " + exception.Message, options, rawReport);
            }
            if (!HasStructurallyValidEmbeddedManifest(structuralReport)) {
                return Result(OfficeProvenanceSigningStatus.Error, "c2patool output did not contain a structurally valid embedded C2PA manifest.", options, rawReport);
            }

            try {
                OfficeFileCommit.CommitTemporaryFileAtomically(
                    stagingPath,
                    outputPath,
                    options.ReplaceExistingOutput
                        ? OfficeFileCommit.ConflictPolicy.Replace
                        : OfficeFileCommit.ConflictPolicy.FailIfExists);
                stagingPath = string.Empty;
            } catch (IOException exception) {
                return Result(OfficeProvenanceSigningStatus.Error, "The signed output could not be committed atomically: " + exception.Message, options, rawReport);
            } catch (UnauthorizedAccessException exception) {
                return Result(OfficeProvenanceSigningStatus.Error, "The signed output could not be committed atomically: " + exception.Message, options, rawReport);
            }

            var findings = new List<string>();
            if (UsesBuiltInTestCredentials) {
                findings.Add("The asset was signed with c2patool's public development credentials and is not a production credential.");
            }
            return new OfficeProvenanceSigningResult(
                OfficeProvenanceSigningStatus.Signed,
                Name,
                findings.AsReadOnly(),
                outputPath,
                structuralReport,
                rawReport);
        } finally {
            OfficeFileCommit.DeleteIfExists(stagingPath);
            OfficeFileCommit.DeleteIfExists(manifestPath);
        }
    }

    private static byte[] CreateManifest(OfficeProvenanceClaim claim) {
        using var stream = new MemoryStream();
        using (var writer = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = false })) {
            writer.WriteStartObject();
            writer.WriteStartArray("claim_generator_info");
            writer.WriteStartObject();
            writer.WriteString("name", claim.ClaimGenerator);
            writer.WriteEndObject();
            writer.WriteEndArray();
            if (claim.Title != null) writer.WriteString("title", claim.Title);
            writer.WriteStartArray("assertions");
            if (claim.Actions.Count > 1) {
                writer.WriteStartObject();
                writer.WriteString("label", "c2pa.actions.v2");
                writer.WriteStartObject("data");
                writer.WriteStartArray("actions");
                for (int index = 1; index < claim.Actions.Count; index++) {
                    OfficeProvenanceAction action = claim.Actions[index];
                    writer.WriteStartObject();
                    writer.WriteString("action", GetActionName(action.Kind));
                    if (OfficeProvenanceDigitalSourceTypes.TryGetUri(action.DigitalSourceKind, out string? digitalSourceType)) {
                        writer.WriteString("digitalSourceType", digitalSourceType);
                    }
                    writer.WriteEndObject();
                }
                writer.WriteEndArray();
                writer.WriteEndObject();
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
            writer.WriteEndObject();
        }
        return stream.ToArray();
    }

    private static string GetActionName(OfficeProvenanceActionKind kind) => kind switch {
        OfficeProvenanceActionKind.Created => "c2pa.created",
        OfficeProvenanceActionKind.Opened => "c2pa.opened",
        OfficeProvenanceActionKind.Edited => "c2pa.edited",
        OfficeProvenanceActionKind.EditedMetadata => "c2pa.edited.metadata",
        OfficeProvenanceActionKind.Converted => "c2pa.converted",
        OfficeProvenanceActionKind.Repackaged => "c2pa.repackaged",
        OfficeProvenanceActionKind.Transcoded => "c2pa.transcoded",
        OfficeProvenanceActionKind.AddedText => "c2pa.addedText",
        OfficeProvenanceActionKind.Cropped => "c2pa.cropped",
        OfficeProvenanceActionKind.Resized => "c2pa.resized",
        OfficeProvenanceActionKind.Published => "c2pa.published",
        _ => "c2pa.unknown"
    };

    private static bool HasStructurallyValidEmbeddedManifest(OfficeProvenanceReport report) {
        foreach (OfficeProvenanceEvidence evidence in report.Evidence) {
            if (evidence.Carrier == OfficeProvenanceCarrierKind.C2paManifest && evidence.IsStructurallyValid) {
                return true;
            }
        }
        return false;
    }

    private OfficeProvenanceSigningResult? ProbeVersion(
        OfficeProvenanceSigningOptions options,
        string workingDirectory) {
        C2paToolProcessResult process;
        try {
            process = _runner.Run(new C2paToolProcessRequest(
                ExecutablePath,
                VersionArguments,
                workingDirectory,
                options.Timeout,
                Math.Min(options.MaxReportBytes, 64L * 1024L)));
        } catch (Win32Exception exception) {
            return Result(OfficeProvenanceSigningStatus.ProviderUnavailable, exception.Message, options);
        } catch (Exception exception) when (exception is TimeoutException or InvalidDataException or IOException or InvalidOperationException) {
            return Result(OfficeProvenanceSigningStatus.ProviderUnavailable, "c2patool version discovery failed: " + exception.Message, options);
        }

        if (process.ExitCode != 0) {
            return Result(
                OfficeProvenanceSigningStatus.ProviderUnavailable,
                ProviderMessage(process, $"c2patool --version exited with code {process.ExitCode}."),
                options);
        }

        string value = process.StandardOutput.Trim();
        const string prefix = "c2patool ";
        if (!value.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) ||
            !Version.TryParse(value.Substring(prefix.Length).Trim(), out Version? version)) {
            return Result(
                OfficeProvenanceSigningStatus.ProviderUnavailable,
                "c2patool returned an unrecognized version string.",
                options);
        }
        if (version < MinimumSupportedVersion) {
            return Result(
                OfficeProvenanceSigningStatus.ProviderUnavailable,
                $"c2patool {version} is unsupported; version {MinimumSupportedVersion} or newer is required.",
                options);
        }
        return null;
    }

    private static string? ValidateIntent(OfficeProvenanceClaim claim, string? parentPath) {
        OfficeProvenanceAction first = claim.Actions[0];
        if (parentPath == null) {
            if (first.Kind != OfficeProvenanceActionKind.Created ||
                !OfficeProvenanceDigitalSourceTypes.TryGetUri(first.DigitalSourceKind, out string? sourceType)) {
                throw new ArgumentException(
                    "A new-asset claim must begin with Created and declare a concrete digital source type.",
                    nameof(claim));
            }
            RejectRepeatedIntentActions(claim);
            int separator = sourceType!.LastIndexOf('/');
            return separator >= 0 ? sourceType.Substring(separator + 1) : sourceType;
        }

        if (first.Kind != OfficeProvenanceActionKind.Opened) {
            throw new ArgumentException("A derived-asset claim must begin with Opened when a parent is supplied.", nameof(claim));
        }
        RejectRepeatedIntentActions(claim);
        return null;
    }

    private static void RejectRepeatedIntentActions(OfficeProvenanceClaim claim) {
        for (int index = 1; index < claim.Actions.Count; index++) {
            OfficeProvenanceActionKind kind = claim.Actions[index].Kind;
            if (kind == OfficeProvenanceActionKind.Created || kind == OfficeProvenanceActionKind.Opened) {
                throw new ArgumentException("Created and Opened are reserved for the first claim action.", nameof(claim));
            }
        }
    }

    private static OfficeProvenanceSigningResult Result(
        OfficeProvenanceSigningStatus status,
        string finding,
        OfficeProvenanceSigningOptions options,
        string? rawReport = null) =>
        new(status, "c2patool", new[] { finding }, rawReport: options.IncludeRawReport ? rawReport : null);

    private static string ProviderMessage(C2paToolProcessResult process, string fallback) {
        if (!string.IsNullOrWhiteSpace(process.StandardError)) return process.StandardError.Trim();
        if (!string.IsNullOrWhiteSpace(process.StandardOutput)) return process.StandardOutput.Trim();
        return fallback;
    }

    private static string JoinOutput(C2paToolProcessResult process) {
        if (string.IsNullOrWhiteSpace(process.StandardError)) return process.StandardOutput;
        if (string.IsNullOrWhiteSpace(process.StandardOutput)) return process.StandardError;
        return process.StandardOutput + Environment.NewLine + process.StandardError;
    }

    private static string NormalizeCommand(string value) => File.Exists(value) ? Path.GetFullPath(value) : value;

    private static bool PathsEqual(string first, string second) => string.Equals(
        first,
        second,
        RuntimeInformation.IsOSPlatform(OSPlatform.Windows) ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal);

    private static void Validate(OfficeProvenanceSigningOptions options) {
        if (options.Timeout <= TimeSpan.Zero || options.Timeout > TimeSpan.FromMinutes(10)) {
            throw new ArgumentOutOfRangeException(nameof(options), "Timeout must be between zero and ten minutes.");
        }
        if (options.MaxReportBytes <= 0 || options.MaxReportBytes > int.MaxValue) {
            throw new ArgumentOutOfRangeException(nameof(options), "MaxReportBytes must be between one and Int32.MaxValue.");
        }
    }
}
