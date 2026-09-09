using OfficeIMO.Core.Internal;
using OfficeIMO.Internal;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private static PreparedRequest PrepareRequest(OfficeWorkflowRequest request) {
        string id = request.Id;
        OfficeWorkflowOperation operation = request.Operation;
        try {
            ValidatedRequest validated = ValidateRequest(request);
            return new PreparedRequest(validated.Id, validated.Operation, validated, null);
        } catch (Exception exception) when (exception is not OutOfMemoryException and not StackOverflowException) {
            return new PreparedRequest(id, operation, null, exception);
        }
    }

    private static ValidatedRequest ValidateRequest(OfficeWorkflowRequest request) {
        if (string.IsNullOrWhiteSpace(request.Id)) throw new ArgumentException("Request id cannot be empty.", nameof(request));
        if (!Enum.IsDefined(typeof(OfficeWorkflowOperation), request.Operation)) {
            throw new ArgumentOutOfRangeException(nameof(request), request.Operation, "Choose a supported workflow operation.");
        }
        if (!Enum.IsDefined(typeof(OfficeWorkflowConflictPolicy), request.ConflictPolicy)) {
            throw new ArgumentOutOfRangeException(nameof(request), request.ConflictPolicy, "Choose a supported output conflict policy.");
        }
        if (!Enum.IsDefined(typeof(OfficeWorkflowOutputProfile), request.OutputProfile)) {
            throw new ArgumentOutOfRangeException(nameof(request), request.OutputProfile, "Choose a supported workflow output profile.");
        }
        if (string.IsNullOrWhiteSpace(request.InputPath)) throw new ArgumentException("Input path cannot be empty.", nameof(request));
        string inputPath = ValidateInputLocation(request.InputPath, request.InputStream);
        string inputName = request.InputStream?.Name ?? inputPath;
        OfficeWorkflowLimits limits = (request.Limits ?? throw new ArgumentException("Workflow limits cannot be null.", nameof(request))).CloneAndValidate();
        OfficeWorkflowRoute? route = null;
        string? comparisonPath = null;
        string? outputPath = string.IsNullOrWhiteSpace(request.OutputPath) ? null
            : request.OutputStream is null ? ValidateLocalOutput(request.OutputPath) : OfficeStorageIdentity.Normalize(request.OutputPath);
        string? outputName = request.OutputStream?.Name ?? outputPath;
        if (request.OutputStream is not null && (outputPath is null || request.ConflictPolicy != OfficeWorkflowConflictPolicy.Replace)) {
            throw new ArgumentException("A provider output requires an explicit destination and the Replace policy after direct-write confirmation.", nameof(request));
        }

        if (request.InputStream is not null && outputPath is null &&
            request.Operation is not (OfficeWorkflowOperation.Inspect or OfficeWorkflowOperation.RepairPlan or OfficeWorkflowOperation.Compare)) {
            throw new ArgumentException("A provider input requires an explicit output destination.", nameof(request));
        }
        if (request.ComparisonStream is not null && request.Operation != OfficeWorkflowOperation.Compare) {
            throw new ArgumentException("A comparison provider is valid only for a comparison operation.", nameof(request));
        }

        if (request.Operation == OfficeWorkflowOperation.Convert) {
            route = OfficeWorkflowCatalog.FindExecutable(request.ConversionRouteId)
                ?? throw new ArgumentException("Choose a supported conversion route.", nameof(request));
            string extension = Path.GetExtension(inputName);
            if (!route.SourceExtensions.Any(item => string.Equals(NormalizeExtension(item), extension, StringComparison.OrdinalIgnoreCase))) {
                throw new ArgumentException($"Route '{route.Id}' does not accept '{extension}' input.", nameof(request));
            }
            outputPath ??= Path.ChangeExtension(inputPath, NormalizeExtension(route.TargetExtension));
            if (!string.Equals(Path.GetExtension(outputName ?? outputPath), NormalizeExtension(route.TargetExtension), StringComparison.OrdinalIgnoreCase)) {
                throw new ArgumentException($"Route '{route.Id}' requires a '{NormalizeExtension(route.TargetExtension)}' output.", nameof(request));
            }
            if (!route.SupportedOutputProfiles.Contains(request.OutputProfile)) {
                throw new ArgumentException(
                    $"The {route.Id} route currently supports only the Faithful output profile.",
                    nameof(request));
            }
        } else if (request.Operation == OfficeWorkflowOperation.Compare) {
            if (string.IsNullOrWhiteSpace(request.ComparisonPath)) throw new ArgumentException("PDF comparison requires a second input path.", nameof(request));
            comparisonPath = ValidateInputLocation(request.ComparisonPath, request.ComparisonStream);
            EnsurePdfExtension(inputName);
            EnsurePdfExtension(request.ComparisonStream?.Name ?? comparisonPath);
            if (outputPath is not null && !string.Equals(Path.GetExtension(outputName ?? outputPath), ".html", StringComparison.OrdinalIgnoreCase)) {
                throw new ArgumentException("Comparison output must be an HTML gallery.", nameof(request));
            }
        } else {
            EnsurePdfExtension(inputName);
            if (request.Operation == OfficeWorkflowOperation.Optimize &&
                request.OutputProfile == OfficeWorkflowOutputProfile.TextOnly) {
                throw new ArgumentException(
                    "Lossless PDF optimization does not support the TextOnly output profile.",
                    nameof(request));
            }
            if (request.Operation is OfficeWorkflowOperation.Optimize or OfficeWorkflowOperation.Repair or OfficeWorkflowOperation.Sanitize or OfficeWorkflowOperation.ExtractPages or OfficeWorkflowOperation.ProtectPdf or OfficeWorkflowOperation.RemovePdfProtection or OfficeWorkflowOperation.SignPdf or OfficeWorkflowOperation.ScanCleanup) {
                outputPath ??= Path.Combine(
                    Path.GetDirectoryName(inputPath)!,
                    Path.GetFileNameWithoutExtension(inputPath) + "." + request.Operation.ToString().ToLowerInvariant() + ".pdf");
                EnsurePdfExtension(outputName ?? outputPath);
            } else if (request.Operation is OfficeWorkflowOperation.Inspect or OfficeWorkflowOperation.RepairPlan && outputPath is not null) {
                throw new ArgumentException("The selected report-only operation does not publish an artifact.", nameof(request));
            }
        }

        if (request.ConversionOptions is not null && route is null)
            throw new ArgumentException("Conversion settings are valid only for conversion operations.", nameof(request));
        OfficeWorkflowConversionOptions? conversionOptions = request.ConversionOptions?.Snapshot(route!);
        OfficeScanCleanupOptions? scanCleanup = request.ScanCleanup?.Snapshot();
        if ((request.Operation == OfficeWorkflowOperation.ScanCleanup) != (scanCleanup != null))
            throw new ArgumentException("ScanCleanup requires scan settings; other operations cannot accept them.", nameof(request));
        if (scanCleanup != null && request.OutputProfile != OfficeWorkflowOutputProfile.Faithful)
            throw new ArgumentException("Scan preparation does not support conversion output profiles.", nameof(request));

        if (request.PageNumbers is { Length: > 100000 })
            throw new ArgumentException("Page extraction is limited to 100,000 selected pages.", nameof(request));
        int[]? pages = request.PageNumbers?.ToArray();
        if (request.Operation == OfficeWorkflowOperation.ExtractPages) {
            if (pages is not { Length: > 0 and <= 100000 } || pages.Any(page => page <= 0))
                throw new ArgumentException("Extraction requires 1 to 100,000 positive one-based page numbers.", nameof(request));
            if (request.OutputProfile != OfficeWorkflowOutputProfile.Faithful)
                throw new ArgumentException("Page extraction supports only the Faithful output profile.", nameof(request));
        } else if (pages is not null) {
            throw new ArgumentException("Page numbers are valid only for page extraction.", nameof(request));
        }

        var encryption = request.OutputEncryption?.Clone();
        bool signing = request.Operation == OfficeWorkflowOperation.SignPdf;
        var signatureOptions = SnapshotSignatureOptions(request.OutputSignatureOptions);
        if (signing) {
            if (request.OutputSigner is null || request.OutputSignatureValidator is null || signatureOptions is null)
                throw new ArgumentException("Signing requires a signer, signature settings, and a cryptographic verifier.", nameof(request));
            if (request.OutputProfile != OfficeWorkflowOutputProfile.Faithful)
                throw new ArgumentException("Signing supports only the Faithful output profile.", nameof(request));
        } else if (request.OutputSigner is not null || signatureOptions is not null || request.OutputSignatureValidator is not null)
            throw new ArgumentException("Signature settings are valid only for SignPdf.", nameof(request));
        bool securityOutput = request.Operation is OfficeWorkflowOperation.ProtectPdf or OfficeWorkflowOperation.RemovePdfProtection;
        if (request.Operation == OfficeWorkflowOperation.ProtectPdf) {
            if (encryption is null || !Enum.IsDefined(encryption.Algorithm))
                throw new ArgumentException("Choose valid PDF output encryption settings.", nameof(request));
        } else if (encryption is not null) throw new ArgumentException("Output encryption is valid only for ProtectPdf.", nameof(request));
        if (!securityOutput && request.PdfOwnerPassword is not null)
            throw new ArgumentException("The owner password is valid only for protection workflows.", nameof(request));
        if (securityOutput && request.OutputProfile != OfficeWorkflowOutputProfile.Faithful)
            throw new ArgumentException("Protection workflows support only the Faithful output profile.", nameof(request));
        string? outputPassword = securityOutput ? encryption?.OwnerPassword ?? encryption?.UserPassword : request.PdfPassword;
        var outputOptions = CreatePdfLoadOptions(outputPassword, limits.MaximumOutputBytes);
        if (encryption?.AesCryptographyProvider is not null) outputOptions = OfficeIMO.Pdf.PdfLoadOptions.WithAesCryptographyProvider(outputOptions, encryption.AesCryptographyProvider);
        OfficeWorkflowStreamInput? inputStream = request.InputStream;
        if ((request.Operation is OfficeWorkflowOperation.ExtractPages or OfficeWorkflowOperation.ScanCleanup || securityOutput || signing) && inputStream is null) {
            inputStream = new OfficeWorkflowStreamInput(Path.GetFileName(inputPath), token => {
                token.ThrowIfCancellationRequested();
                return Task.FromResult<Stream>(new FileStream(inputPath, FileMode.Open, FileAccess.Read, FileShare.Read));
            });
        }

        return new ValidatedRequest(
            request.Id,
            request.Operation,
            inputPath,
            comparisonPath,
            outputPath,
            route,
            request.ConflictPolicy,
            request.OutputProfile,
            limits,
            CreatePdfLoadOptions(securityOutput ? request.PdfOwnerPassword ?? request.PdfPassword : request.PdfPassword, limits.MaximumInputBytes),
            CreatePdfLoadOptions(request.ComparisonPdfPassword ?? request.PdfPassword, limits.MaximumInputBytes),
            outputOptions,
            request.PublicationGuard,
            inputStream, request.ComparisonStream, request.OutputStream, pages, encryption, request.PdfOwnerPassword ?? request.PdfPassword,
            request.OutputSigner, signatureOptions, request.OutputSignatureValidator, conversionOptions, scanCleanup);
    }

    private static string ValidateInputLocation(string location, OfficeWorkflowStreamInput? stream) {
        if (stream is not null) return OfficeStorageIdentity.Normalize(location);
        string path = OfficeStorageIdentity.GetLocalPath(location)
            ?? throw new ArgumentException("A provider input requires a stream access contract.", nameof(location));
        if (!File.Exists(path)) throw new FileNotFoundException("The workflow input file does not exist.", path);
        return path;
    }

    private static string ValidateLocalOutput(string location) => OfficeStorageIdentity.GetLocalPath(location)
        ?? throw new NotSupportedException("This workflow requires a filesystem output destination. Choose a local destination.");

    private static void ReportInputStagingCleanupFailure(Exception error, List<OfficeWorkflowDiagnostic> diagnostics) {
        if (error.Data[OfficeStreamFileSnapshot.CleanupFailureDataKey] is string directory) {
            diagnostics.Add(new OfficeWorkflowDiagnostic("InputStagingCleanupFailed",
                "Private input staging could not be removed: " + directory,
                OfficeWorkflowDiagnosticSeverity.Warning, "cleanup"));
        }
    }

    private sealed partial class WorkflowInputSnapshots : IDisposable {
        private readonly List<(OfficeStreamFileSnapshot Snapshot, OfficeWorkflowStreamInput Source, WorkflowSourceAccess Access)> _snapshots = new();

        internal async Task<ValidatedImageExportRequest> CaptureAsync(ValidatedImageExportRequest request, CancellationToken token) {
            string inputPath = await CaptureOneAsync(request.InputPath, request.InputStream, request.Limits.MaximumInputBytes, token).ConfigureAwait(false);
            return request with { InputPath = inputPath, PublicationGuard = Guard(request.PublicationGuard, request.Limits.MaximumInputBytes, [request.InputPath]) };
        }

        internal async Task<ValidatedRequest> CaptureAsync(ValidatedRequest request, CancellationToken token) {
            string inputPath = await CaptureOneAsync(request.InputPath, request.InputStream, request.Limits.MaximumInputBytes, token).ConfigureAwait(false);
            string? comparisonPath = request.ComparisonPath is null ? null
                : await CaptureOneAsync(request.ComparisonPath, request.ComparisonStream, request.Limits.MaximumInputBytes, token).ConfigureAwait(false);
            string[]? protectedSources = request.InputStream is null && request.ComparisonStream is null && request.OutputStream is null
                ? null : new[] { request.InputPath, request.ComparisonPath }.OfType<string>().ToArray();
            return request with {
                InputPath = inputPath,
                ComparisonPath = comparisonPath,
                PublicationGuard = Guard(request.PublicationGuard, request.Limits.MaximumInputBytes, protectedSources, request.OutputStream)
            };
        }

        internal async Task<ValidatedAssemblyRequest> CaptureAsync(ValidatedAssemblyRequest request, CancellationToken token) {
            var sources = new List<string>(request.Sources.Count);
            var stagedStreams = new Dictionary<string, OfficeWorkflowStreamInput>(StringComparer.Ordinal);
            var sourceLocations = new Dictionary<string, string>(StringComparer.Ordinal);
            long remainingBytes = request.Limits.MaximumInputBytes;
            int remainingEntries = request.Options.MaximumDiscoveredEntries;
            foreach (string location in request.Sources) {
                if (request.SourceDirectories?.TryGetValue(location, out OfficeWorkflowDirectoryInput? directory) == true) {
                    var captured = await CaptureDirectoryAsync(directory, request.Options.IncludeSubdirectories,
                        remainingEntries, remainingBytes, token).ConfigureAwait(false);
                    remainingBytes -= captured.Bytes;
                    remainingEntries -= captured.Entries;
                    sources.Add(captured.Path);
                    sourceLocations[captured.Path] = location;
                    continue;
                }
                request.SourceStreams.TryGetValue(location, out OfficeWorkflowStreamInput? source);
                if (source is not null && remainingBytes <= 0) throw new InvalidDataException("The provider inputs exceed the workflow input limit.");
                string path = await CaptureOneAsync(location, source, remainingBytes, token).ConfigureAwait(false);
                if (source is not null) {
                    remainingBytes -= _snapshots[^1].Snapshot.Length;
                    stagedStreams[path] = source;
                    sourceLocations[path] = location;
                }
                sources.Add(path);
            }
            return request with {
                Sources = sources,
                SourceStreams = stagedStreams,
                SourceLocations = sourceLocations,
                PublicationGuard = Guard(request.PublicationGuard, request.Limits.MaximumInputBytes,
                    request.Sources.Where(source => request.SourceDirectories?.ContainsKey(source) != true)
                        .Concat(_directoryFiles.Select(item => item.Access.Location)).ToArray(), request.OutputStream)
            };
        }

        internal IOfficeWorkflowPublicationGuard? Guard(IOfficeWorkflowPublicationGuard? host, long maximumBytes,
            string[]? protectedSources = null, OfficeWorkflowStreamOutput? output = null) {
            if (_directoryManifests.Count > 0) host = new DirectoryMembershipPublicationGuard(host, _directoryManifests.ToArray());
            if (protectedSources is not null) host = new WorkflowScopedSourcePublicationGuard(host, protectedSources,
                _snapshots.Select(item => item.Access).Concat(_directoryFiles.Select(item => item.Access)).ToArray(), output);
            var fingerprints = _snapshots.Select(item => (item.Source, item.Snapshot.Fingerprint))
                .Concat(_directoryFiles.Select(item => (item.Source, item.Fingerprint))).ToArray();
            return fingerprints.Length == 0 ? host : new VerifiedProviderPublicationGuard(host, fingerprints, maximumBytes);
        }

        internal async Task<string> CaptureOneAsync(string location, OfficeWorkflowStreamInput? source, long maximumBytes, CancellationToken token) {
            if (source is null) return location;
            var access = new WorkflowSourceAccess(location, source);
            source = access.CreateInput();
            var snapshot = await OfficeStreamFileSnapshot.CaptureAsync(source.OpenRead, Path.GetExtension(source.Name),
                maximumBytes, source.ExpectedSha256, token).ConfigureAwait(false);
            _snapshots.Add((snapshot, source, access));
            return snapshot.FilePath;
        }

        internal async Task VerifyAsync(long maximumBytes, CancellationToken token) {
            foreach (var item in _snapshots) await item.Snapshot.VerifySourceAsync(item.Source.OpenRead, maximumBytes, token).ConfigureAwait(false);
        }

        internal void Cleanup(List<OfficeWorkflowDiagnostic> diagnostics) {
            try { Dispose(); } catch (IOException error) {
                diagnostics.Add(new OfficeWorkflowDiagnostic("InputStagingCleanupFailed", error.Message,
                    OfficeWorkflowDiagnosticSeverity.Warning, "cleanup"));
            }
        }

        public void Dispose() {
            List<Exception>? failures = null;
            CleanupDirectories(ref failures);
            foreach (var item in _snapshots) {
                try { item.Snapshot.Dispose(); } catch (Exception error) when (error is IOException or UnauthorizedAccessException) {
                    (failures ??= new()).Add(error);
                }
            }
            if (failures is not null) throw new IOException("Private workflow inputs could not be removed.", new AggregateException(failures));
            _snapshots.Clear();
        }
    }

    private sealed class VerifiedProviderPublicationGuard(IOfficeWorkflowPublicationGuard? host,
        (OfficeWorkflowStreamInput Source, string Fingerprint)[] inputs, long maximumBytes) : IOfficeWorkflowPublicationGuard {
        public async ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken token) {
            if (host is not null && !await host.CanPublishAsync(path, isDirectory, token).ConfigureAwait(false)) return false;
            foreach (var input in inputs) {
                await OfficeStreamPublication.VerifyFingerprintAsync(input.Source.OpenRead, input.Fingerprint, maximumBytes, token).ConfigureAwait(false);
            }
            return true;
        }
    }

}