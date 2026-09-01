using System.Diagnostics;
using System.IO.Compression;
using OfficeIMO.Core.Internal;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private static readonly HashSet<string> AssemblyImageExtensions = new(StringComparer.OrdinalIgnoreCase) {
        ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".webp", ".ico", ".pcx"
    };
    private static readonly IReadOnlyDictionary<string, string> AssemblyOfficeRoutes =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
            [".docx"] = "docx-pdf",
            [".xlsx"] = "xlsx-pdf",
            [".pptx"] = "pptx-pdf",
            [".html"] = "html-pdf",
            [".htm"] = "html-pdf"
        };

    /// <inheritdoc />
    public async Task<PdfAssemblyResult> AssemblePdfAsync(
        PdfAssemblyRequest request,
        IProgress<OfficeWorkflowProgress>? progress = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(request);
        var stopwatch = Stopwatch.StartNew();
        var diagnostics = new List<OfficeWorkflowDiagnostic>();
        string? stagingPath = null;
        string? extractionRoot = null;
        int sourceCount = 0;
        int pageCount = 0;
        long inputBytes = 0L;
        long normalizedBytes = 0L;
        WorkflowFailureStage failureStage = WorkflowFailureStage.Validation;

        try {
            ValidatedAssemblyRequest validated = ValidateAssemblyRequest(request);
            failureStage = WorkflowFailureStage.Input;
            Report(progress, validated.Id, "discover", "Discovering supported input documents", 0.04D);
            cancellationToken.ThrowIfCancellationRequested();

            extractionRoot = Path.Combine(Path.GetTempPath(), "officeimo-assembly-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(extractionRoot);
            IReadOnlyList<AssemblySource> sources = ExpandAssemblySources(validated, extractionRoot, diagnostics, cancellationToken);
            if (validated.OutputProfile != OfficeWorkflowOutputProfile.Faithful &&
                sources.Any(static source => source.Route?.Id == "html-pdf")) {
                throw new ArgumentException(
                    "HTML assembly sources currently support only the Faithful output profile.",
                    nameof(request));
            }
            sourceCount = sources.Count;
            inputBytes = sources.Sum(static source => source.SizeBytes);
            if (inputBytes > validated.Limits.MaximumInputBytes) {
                throw new InvalidOperationException(
                    $"Expanded inputs total {inputBytes:N0} bytes, above the configured {validated.Limits.MaximumInputBytes:N0}-byte limit.");
            }

            var documents = new List<PdfDocument>(sources.Count);
            for (int index = 0; index < sources.Count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                AssemblySource source = sources[index];
                double fraction = 0.12D + (double)index / Math.Max(1, sources.Count) * 0.55D;
                Report(progress, validated.Id, "normalize", $"Preparing {index + 1:N0} of {sources.Count:N0} inputs", fraction);
                long remainingOutputBytes = validated.Limits.MaximumOutputBytes - normalizedBytes;
                if (remainingOutputBytes <= 0L) {
                    throw new InvalidOperationException(
                        $"Normalized inputs exceed the configured {validated.Limits.MaximumOutputBytes:N0}-byte output limit.");
                }
                PdfDocument normalized = NormalizeAssemblySource(
                    source,
                    validated,
                    remainingOutputBytes,
                    diagnostics,
                    cancellationToken);
                using (var measurement = new OfficeWorkflowBoundedCountingStream(remainingOutputBytes)) {
                    await normalized.SaveAsync(measurement, cancellationToken).ConfigureAwait(false);
                    normalizedBytes = checked(normalizedBytes + measurement.Length);
                }
                documents.Add(normalized);
            }

            Report(progress, validated.Id, "merge", "Combining normalized PDF pages", 0.7D);
            failureStage = WorkflowFailureStage.Operation;
            cancellationToken.ThrowIfCancellationRequested();
            PdfDocument merged = documents.Count == 1 ? documents[0] : PdfDocument.Merge(documents);

            string outputDirectory = Path.GetDirectoryName(validated.OutputPath)!;
            failureStage = WorkflowFailureStage.Output;
            Directory.CreateDirectory(outputDirectory);
            stagingPath = Path.Combine(
                outputDirectory,
                "." + Path.GetFileName(validated.OutputPath) + "." + Guid.NewGuid().ToString("N") + ".tmp");
            await using (var outputStream = new FileStream(
                stagingPath,
                FileMode.CreateNew,
                FileAccess.Write,
                FileShare.None,
                81920,
                FileOptions.Asynchronous | FileOptions.SequentialScan))
            await using (var boundedOutput = new OfficeWorkflowBoundedWriteStream(
                outputStream,
                validated.Limits.MaximumOutputBytes,
                leaveOpen: false)) {
                await merged.SaveAsync(boundedOutput, cancellationToken).ConfigureAwait(false);
            }
            cancellationToken.ThrowIfCancellationRequested();
            long stagedOutputBytes = new FileInfo(stagingPath).Length;

            Report(progress, validated.Id, "validate-output", "Reopening the assembled PDF", 0.84D);
            PdfDocumentInfo info = PdfDocument.Load(stagingPath, validated.PdfLoadOptions).Inspect();
            pageCount = info.PageCount;
            if (pageCount < 1) throw new InvalidOperationException("The assembled PDF has no pages.");
            diagnostics.Add(new OfficeWorkflowDiagnostic(
                "AssemblyReopened",
                "The staged PDF was reopened through OfficeIMO.Pdf before publication.",
                stage: "validate-output",
                details: new Dictionary<string, string>(StringComparer.Ordinal) {
                    ["sourceCount"] = sources.Count.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    ["pageCount"] = pageCount.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    ["normalizedBytes"] = normalizedBytes.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    ["outputBytes"] = stagedOutputBytes.ToString(System.Globalization.CultureInfo.InvariantCulture)
                }));

            Report(progress, validated.Id, "publish", "Publishing the validated PDF", 0.93D);
            string publishedPath = Publish(stagingPath, validated.OutputPath, validated.ConflictPolicy);
            stagingPath = null;
            long outputBytes = new FileInfo(publishedPath).Length;
            Report(progress, validated.Id, "complete", "Assembled PDF is ready", 1D);
            return new PdfAssemblyResult(
                validated.Id,
                OfficeWorkflowStatus.Completed,
                OfficeWorkflowFailureKind.None,
                publishedPath,
                sourceCount,
                pageCount,
                inputBytes,
                outputBytes,
                stopwatch.Elapsed,
                $"Assembled {sourceCount:N0} {(sourceCount == 1 ? "input" : "inputs")} into {pageCount:N0} PDF {(pageCount == 1 ? "page" : "pages")}.",
                diagnostics);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            diagnostics.Add(new OfficeWorkflowDiagnostic(
                "Cancelled",
                "PDF assembly was cancelled before publication.",
                OfficeWorkflowDiagnosticSeverity.Information,
                "cancel"));
            return new PdfAssemblyResult(
                request.Id,
                OfficeWorkflowStatus.Cancelled,
                OfficeWorkflowFailureKind.None,
                null,
                sourceCount,
                pageCount,
                inputBytes,
                0L,
                stopwatch.Elapsed,
                "Cancelled",
                diagnostics);
        } catch (Exception ex) when (ex is not OutOfMemoryException and not StackOverflowException) {
            diagnostics.Add(new OfficeWorkflowDiagnostic(
                "PdfAssemblyFailed",
                ex.Message,
                OfficeWorkflowDiagnosticSeverity.Error,
                "execute",
                new Dictionary<string, string>(StringComparer.Ordinal) { ["exceptionType"] = ex.GetType().Name }));
            return new PdfAssemblyResult(
                request.Id,
                OfficeWorkflowStatus.Failed,
                ClassifyFailure(ex, failureStage),
                null,
                sourceCount,
                pageCount,
                inputBytes,
                0L,
                stopwatch.Elapsed,
                "PDF assembly failed: " + ex.Message,
                diagnostics);
        } finally {
            if (stagingPath is not null) TryDelete(stagingPath);
            if (extractionRoot is not null) TryDeleteDirectory(extractionRoot);
        }
    }

    private static ValidatedAssemblyRequest ValidateAssemblyRequest(PdfAssemblyRequest request) {
        if (string.IsNullOrWhiteSpace(request.Id)) throw new ArgumentException("Request id cannot be empty.", nameof(request));
        if (request.Sources == null || request.Sources.Count == 0) throw new ArgumentException("At least one source is required.", nameof(request));
        if (request.Sources.Any(string.IsNullOrWhiteSpace)) throw new ArgumentException("Source paths cannot be empty.", nameof(request));
        if (string.IsNullOrWhiteSpace(request.OutputPath)) throw new ArgumentException("Output path cannot be empty.", nameof(request));
        PdfAssemblyOptions options = (request.Options ?? throw new ArgumentException("Assembly options cannot be null.", nameof(request))).CloneAndValidate();
        if (request.Sources.Count > options.MaximumSourceCount) {
            throw new ArgumentException(
                $"Source count exceeds the configured {options.MaximumSourceCount:N0}-item limit.",
                nameof(request));
        }
        string outputPath = Path.GetFullPath(request.OutputPath);
        EnsurePdfExtension(outputPath);
        string[] sources = request.Sources.Select(Path.GetFullPath).ToArray();
        if (sources.Any(path => OfficeWorkflowPathIdentity.AreEquivalent(path, outputPath))) {
            throw new ArgumentException("The output PDF cannot also be an explicit input.", nameof(request));
        }
        foreach (string path in sources) {
            if (!File.Exists(path) && !Directory.Exists(path)) {
                throw new FileNotFoundException("An assembly source does not exist.", path);
            }
        }
        if (!Enum.IsDefined(request.ConflictPolicy)) throw new ArgumentOutOfRangeException(nameof(request.ConflictPolicy));
        if (!Enum.IsDefined(request.OutputProfile)) throw new ArgumentOutOfRangeException(nameof(request.OutputProfile));
        OfficeWorkflowLimits limits = (request.Limits ?? throw new ArgumentException("Workflow limits cannot be null.", nameof(request))).CloneAndValidate();
        return new ValidatedAssemblyRequest(
            request.Id,
            sources,
            outputPath,
            request.ConflictPolicy,
            request.OutputProfile,
            options,
            limits,
            new PdfLoadOptions { Password = request.PdfPassword });
    }

    private static IReadOnlyList<AssemblySource> ExpandAssemblySources(
        ValidatedAssemblyRequest request,
        string extractionRoot,
        List<OfficeWorkflowDiagnostic> diagnostics,
        CancellationToken cancellationToken) {
        var sources = new List<AssemblySource>();
        long archiveBytes = 0L;
        int archiveIndex = 0;
        int discoveredEntryCount = 0;
        foreach (string input in request.Sources) {
            cancellationToken.ThrowIfCancellationRequested();
            if (Directory.Exists(input)) {
                List<string> files = [];
                try {
                    foreach (string file in Directory.EnumerateFiles(input, "*", new EnumerationOptions {
                            RecurseSubdirectories = request.Options.IncludeSubdirectories,
                            IgnoreInaccessible = false,
                            AttributesToSkip = FileAttributes.ReparsePoint,
                            ReturnSpecialDirectories = false
                        })) {
                        cancellationToken.ThrowIfCancellationRequested();
                        CountDiscoveredEntry();
                        files.Add(file);
                    }
                } catch (UnauthorizedAccessException ex) {
                    throw new IOException("A source folder could not be enumerated safely: " + input, ex);
                }
                StringComparer pathComparer = OfficeWorkflowPathIdentity.GetComparer(input);
                foreach (string file in files
                             .OrderBy(static path => path, pathComparer)
                             .ThenBy(static path => path, StringComparer.Ordinal)) {
                    if (OfficeWorkflowPathIdentity.AreEquivalent(file, request.OutputPath)) continue;
                    AddDiscoveredSource(file, input, discovered: true);
                }
            } else {
                AddDiscoveredSource(input, input, discovered: false);
            }
        }
        if (sources.Count == 0) throw new InvalidOperationException("No supported documents were found in the requested sources.");
        return sources;

        void AddDiscoveredSource(string path, string origin, bool discovered) {
            cancellationToken.ThrowIfCancellationRequested();
            string extension = Path.GetExtension(path);
            if (string.Equals(extension, ".zip", StringComparison.OrdinalIgnoreCase)) {
                archiveIndex++;
                ExpandArchive(path, origin, archiveIndex);
                return;
            }
            if (!TryClassifyAssemblySource(path, out AssemblySourceKind kind, out OfficeWorkflowRoute? route)) {
                if (discovered && request.Options.IgnoreDiscoveredUnsupportedFiles) return;
                throw new NotSupportedException("No PDF assembly intake route is available for '" + path + "'.");
            }
            long size = new FileInfo(path).Length;
            EnforceInputLimit(path, size, request.Limits);
            Add(new AssemblySource(path, origin, Path.GetFileName(path), kind, route, size));
        }

        void ExpandArchive(string archivePath, string origin, int index) {
            long archiveFileBytes = new FileInfo(archivePath).Length;
            EnforceInputLimit(archivePath, archiveFileBytes, request.Limits);
            string destinationRoot = Path.Combine(extractionRoot, "archive-" + index.ToString("D4", System.Globalization.CultureInfo.InvariantCulture));
            Directory.CreateDirectory(destinationRoot);
            string canonicalDestinationRoot = Path.GetFullPath(destinationRoot)
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
            StringComparison destinationComparison = OfficeWorkflowPathIdentity.GetComparison(destinationRoot);
            using var archiveStream = new FileStream(archivePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            int remainingDiscoveryCapacity = request.Options.MaximumDiscoveredEntries - discoveredEntryCount;
            int preflightLimit = Math.Min(request.Options.MaximumArchiveEntries, remainingDiscoveryCapacity);
            OfficeArchiveSafety.ZipCentralDirectoryScanResult preflight =
                OfficeArchiveSafety.ScanZipCentralDirectory(archiveStream, archiveStream.Length, preflightLimit);
            if (!preflight.IsValid) {
                throw new InvalidDataException(preflight.Error ?? "The ZIP central directory is malformed.");
            }
            if (preflight.LimitExceeded) {
                if (preflight.EntryCount > request.Options.MaximumArchiveEntries) {
                    throw new InvalidDataException(
                        $"Archive '{Path.GetFileName(archivePath)}' declares {preflight.EntryCount:N0} entries, above the configured {request.Options.MaximumArchiveEntries:N0}-entry limit.");
                }
                throw new InvalidOperationException(
                    $"Discovered entry count exceeds the configured {request.Options.MaximumDiscoveredEntries:N0}-item limit.");
            }
            discoveredEntryCount = checked(discoveredEntryCount + checked((int)preflight.EntryCount));

            using var archive = new ZipArchive(archiveStream, ZipArchiveMode.Read, leaveOpen: false);
            if (archive.Entries.Count > request.Options.MaximumArchiveEntries) {
                throw new InvalidDataException(
                    $"Archive '{Path.GetFileName(archivePath)}' contains {archive.Entries.Count:N0} entries, above the configured {request.Options.MaximumArchiveEntries:N0}-entry limit.");
            }
            if (archive.Entries.Count != preflight.EntryCount) {
                throw new InvalidDataException("The ZIP entry count changed after bounded central-directory preflight.");
            }

            int observedArchiveEntries = 0;
            foreach (ZipArchiveEntry entry in archive.Entries
                         .OrderBy(static item => item.FullName, StringComparer.OrdinalIgnoreCase)
                         .ThenBy(static item => item.FullName, StringComparer.Ordinal)) {
                cancellationToken.ThrowIfCancellationRequested();
                observedArchiveEntries = checked(observedArchiveEntries + 1);
                if (observedArchiveEntries > preflight.EntryCount ||
                    observedArchiveEntries > request.Options.MaximumArchiveEntries) {
                    throw new InvalidDataException("The ZIP produced more entries than its bounded central-directory preflight declared.");
                }
                if (string.IsNullOrEmpty(entry.Name)) continue;
                string extension = Path.GetExtension(entry.Name);
                if (string.Equals(extension, ".zip", StringComparison.OrdinalIgnoreCase)) {
                    throw new InvalidDataException("Nested ZIP archives are not expanded.");
                }
                if (!TryClassifyAssemblySource(entry.Name, out AssemblySourceKind kind, out OfficeWorkflowRoute? route)) {
                    if (request.Options.IgnoreDiscoveredUnsupportedFiles) continue;
                    throw new NotSupportedException("No PDF assembly intake route is available for archive entry '" + entry.FullName + "'.");
                }
                if (entry.Length > request.Options.MaximumArchiveEntryBytes) {
                    throw new InvalidDataException("An archive entry exceeds the configured uncompressed-size limit.");
                }
                if (entry.Length > 0L) {
                    if (entry.CompressedLength == 0L || (double)entry.Length / entry.CompressedLength > request.Options.MaximumArchiveCompressionRatio) {
                        throw new InvalidDataException("An archive entry exceeds the configured compression-ratio limit.");
                    }
                }
                archiveBytes = checked(archiveBytes + entry.Length);
                if (archiveBytes > request.Options.MaximumArchiveBytes || archiveBytes > request.Limits.MaximumInputBytes) {
                    throw new InvalidDataException("Expanded archive content exceeds the configured aggregate input limit.");
                }

                string normalizedName = entry.FullName.Replace('/', Path.DirectorySeparatorChar).Replace('\\', Path.DirectorySeparatorChar);
                string destination = Path.GetFullPath(Path.Combine(destinationRoot, normalizedName));
                if (!destination.StartsWith(canonicalDestinationRoot, destinationComparison)) {
                    throw new InvalidDataException("An archive entry resolves outside its extraction directory.");
                }
                Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
                using (Stream source = entry.Open())
                using (FileStream target = new(destination, FileMode.CreateNew, FileAccess.Write, FileShare.None)) {
                    CopyBounded(source, target, entry.Length, request.Options.MaximumArchiveEntryBytes, cancellationToken);
                }
                Add(new AssemblySource(destination, origin, entry.FullName, kind, route, entry.Length));
            }
            if (observedArchiveEntries != preflight.EntryCount) {
                throw new InvalidDataException("The ZIP produced fewer entries than its bounded central-directory preflight declared.");
            }
            diagnostics.Add(new OfficeWorkflowDiagnostic(
                "ArchiveExpanded",
                "A ZIP source was expanded with entry-count, path, size, and compression-ratio guards.",
                stage: "discover",
                details: new Dictionary<string, string>(StringComparer.Ordinal) {
                    ["archive"] = Path.GetFileName(archivePath),
                    ["entryCount"] = archive.Entries.Count.ToString(System.Globalization.CultureInfo.InvariantCulture)
                }));
        }

        void Add(AssemblySource source) {
            if (sources.Count >= request.Options.MaximumSourceCount) {
                throw new InvalidOperationException(
                    $"Expanded source count exceeds the configured {request.Options.MaximumSourceCount:N0}-item limit.");
            }
            sources.Add(source);
        }

        void CountDiscoveredEntry() {
            cancellationToken.ThrowIfCancellationRequested();
            discoveredEntryCount = checked(discoveredEntryCount + 1);
            if (discoveredEntryCount > request.Options.MaximumDiscoveredEntries) {
                throw new InvalidOperationException(
                    $"Discovered entry count exceeds the configured {request.Options.MaximumDiscoveredEntries:N0}-item limit.");
            }
        }
    }

    private static PdfDocument NormalizeAssemblySource(
        AssemblySource source,
        ValidatedAssemblyRequest request,
        long maximumNormalizedBytes,
        List<OfficeWorkflowDiagnostic> diagnostics,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        switch (source.Kind) {
            case AssemblySourceKind.Pdf:
                PdfDocument opened = PdfDocument.Load(source.Path, request.PdfLoadOptions);
                _ = opened.Inspect();
                AddAssemblySourceDiagnostic(source, "PDF pages retained", diagnostics);
                return opened;
            case AssemblySourceKind.Image:
                PdfDocument imageDocument = PdfDocument.CreateFromImages(
                    [new PdfImageDocumentSource(File.ReadAllBytes(source.Path), source.DisplayName)],
                    request.Options.ImageOptions);
                AddAssemblySourceDiagnostic(source, "Image composed as one PDF page", diagnostics);
                return imageDocument;
            case AssemblySourceKind.Office:
                var conversionRequest = new ValidatedRequest(
                    request.Id,
                    OfficeWorkflowOperation.Convert,
                    source.Path,
                    ComparisonPath: null,
                    OutputPath: null,
                    source.Route,
                    request.ConflictPolicy,
                    request.OutputProfile,
                    new OfficeWorkflowLimits {
                        MaximumInputBytes = request.Limits.MaximumInputBytes,
                        MaximumOutputBytes = maximumNormalizedBytes
                    },
                    request.PdfLoadOptions,
                    request.PdfLoadOptions);
                OperationArtifact artifact = Convert(conversionRequest, diagnostics, cancellationToken);
                if (artifact.Bytes == null) throw new InvalidOperationException("An Office input did not produce PDF bytes.");
                AddAssemblySourceDiagnostic(source, "Office document normalized to PDF", diagnostics);
                return PdfDocument.Load(artifact.Bytes);
            default:
                throw new ArgumentOutOfRangeException(nameof(source));
        }
    }

    private static bool TryClassifyAssemblySource(
        string path,
        out AssemblySourceKind kind,
        out OfficeWorkflowRoute? route) {
        string extension = Path.GetExtension(path);
        if (string.Equals(extension, ".pdf", StringComparison.OrdinalIgnoreCase)) {
            kind = AssemblySourceKind.Pdf;
            route = null;
            return true;
        }
        if (AssemblyImageExtensions.Contains(extension)) {
            kind = AssemblySourceKind.Image;
            route = null;
            return true;
        }
        route = AssemblyOfficeRoutes.TryGetValue(extension, out string? routeId)
            ? OfficeWorkflowCatalog.Find(routeId)
            : null;
        if (route != null) {
            kind = AssemblySourceKind.Office;
            return true;
        }
        kind = default;
        return false;
    }

    private static void AddAssemblySourceDiagnostic(
        AssemblySource source,
        string action,
        ICollection<OfficeWorkflowDiagnostic> diagnostics) {
        diagnostics.Add(new OfficeWorkflowDiagnostic(
            "AssemblySourceNormalized",
            action + ": " + source.DisplayName,
            stage: "normalize",
            details: new Dictionary<string, string>(StringComparer.Ordinal) {
                ["name"] = source.DisplayName,
                ["kind"] = source.Kind.ToString(),
                ["origin"] = Path.GetFileName(source.Origin),
                ["bytes"] = source.SizeBytes.ToString(System.Globalization.CultureInfo.InvariantCulture)
            }));
    }

    private static void CopyBounded(
        Stream source,
        Stream target,
        long declaredLength,
        long maximumBytes,
        CancellationToken cancellationToken) {
        byte[] buffer = new byte[81_920];
        long written = 0L;
        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            int read = source.Read(buffer, 0, buffer.Length);
            if (read == 0) break;
            written = checked(written + read);
            if (written > maximumBytes || written > declaredLength) {
                throw new InvalidDataException("An archive entry produced more bytes than its declared or configured limit.");
            }
            target.Write(buffer, 0, read);
        }
        if (written != declaredLength) throw new InvalidDataException("An archive entry did not match its declared length.");
    }

    private enum AssemblySourceKind {
        Pdf,
        Image,
        Office
    }

    private sealed record AssemblySource(
        string Path,
        string Origin,
        string DisplayName,
        AssemblySourceKind Kind,
        OfficeWorkflowRoute? Route,
        long SizeBytes);

    private sealed record ValidatedAssemblyRequest(
        string Id,
        IReadOnlyList<string> Sources,
        string OutputPath,
        OfficeWorkflowConflictPolicy ConflictPolicy,
        OfficeWorkflowOutputProfile OutputProfile,
        PdfAssemblyOptions Options,
        OfficeWorkflowLimits Limits,
        PdfLoadOptions PdfLoadOptions);
}
