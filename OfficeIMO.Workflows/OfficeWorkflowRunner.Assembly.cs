using System.Diagnostics;
using System.IO.Compression;
using System.Security.Cryptography;
using OfficeIMO.Core.Internal;
using OfficeIMO.Html;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private const int AssemblyInputBufferSize = 81920;
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
        ValidatedAssemblyRequest? validated = null;

        try {
            validated = ValidateAssemblyRequest(request);
            failureStage = WorkflowFailureStage.Input;
            Report(progress, validated.Id, "discover", "Discovering supported input documents", 0.04D);
            cancellationToken.ThrowIfCancellationRequested();

            extractionRoot = Path.Combine(Path.GetTempPath(), "officeimo-assembly-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(extractionRoot);
            IReadOnlyList<AssemblySource> sources = ExpandAssemblySources(
                validated,
                extractionRoot,
                diagnostics,
                out inputBytes,
                cancellationToken);
            if (validated.OutputProfile != OfficeWorkflowOutputProfile.Faithful &&
                sources.Any(static source => source.Route?.Id == "html-pdf")) {
                throw new ArgumentException(
                    "HTML assembly sources currently support only the Faithful output profile.",
                    nameof(request));
            }
            sourceCount = sources.Count;
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
            PdfDocument reopened = await PdfDocument
                .LoadAsync(stagingPath, validated.OutputPdfLoadOptions, cancellationToken)
                .ConfigureAwait(false);
            PdfDocumentInfo info = reopened.Inspect(validated.OutputPdfLoadOptions, cancellationToken);
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
            cancellationToken.ThrowIfCancellationRequested();
            string publishedPath = Publish(stagingPath, validated.OutputPath, validated.ConflictPolicy, cancellationToken);
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
                validated?.Id ?? request.Id,
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
                GetDiagnosticStage(failureStage),
                new Dictionary<string, string>(StringComparer.Ordinal) { ["exceptionType"] = ex.GetType().Name }));
            return new PdfAssemblyResult(
                validated?.Id ?? request.Id,
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
            CreatePdfLoadOptions(request.PdfPassword, limits.MaximumInputBytes),
            CreatePdfLoadOptions(request.PdfPassword, limits.MaximumOutputBytes));
    }

    private static IReadOnlyList<AssemblySource> ExpandAssemblySources(
        ValidatedAssemblyRequest request,
        string extractionRoot,
        List<OfficeWorkflowDiagnostic> diagnostics,
        out long inputBytes,
        CancellationToken cancellationToken) {
        var sources = new List<AssemblySource>();
        long expandedInputBytes = 0L;
        long archiveBytes = 0L;
        int archiveIndex = 0;
        int discoveredEntryCount = 0;
        foreach (string input in request.Sources) {
            cancellationToken.ThrowIfCancellationRequested();
            if (Directory.Exists(input)) {
                string physicalRoot = OfficeWorkflowPathIdentity.ResolvePhysicalPath(input);
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
                HtmlDependencyDiscovery dependencyDiscovery = FindReferencedHtmlDependencies(
                    files,
                    physicalRoot,
                    request.Limits.MaximumInputBytes,
                    pathComparer,
                    cancellationToken);
                AddInputBytes(dependencyDiscovery.TotalBytes);
                foreach (string file in files
                             .OrderBy(static path => path, pathComparer)
                             .ThenBy(static path => path, StringComparer.Ordinal)) {
                    if (IsAssemblyOutputCandidate(file, request.OutputPath, request.ConflictPolicy)) continue;
                    if (dependencyDiscovery.Paths.Contains(Path.GetFullPath(file))) continue;
                    AddDiscoveredSource(file, input, discovered: true, physicalRoot, dependencyDiscovery.Snapshots);
                }
            } else {
                string sourceDirectory = Path.GetDirectoryName(input)!;
                string physicalRoot = OfficeWorkflowPathIdentity.ResolvePhysicalPath(sourceDirectory);
                IReadOnlyDictionary<string, byte[]>? dependencySnapshots = null;
                string extension = Path.GetExtension(input);
                if (string.Equals(extension, ".html", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(extension, ".htm", StringComparison.OrdinalIgnoreCase)) {
                    StringComparer pathComparer = OfficeWorkflowPathIdentity.GetComparer(sourceDirectory);
                    string fullInputPath = Path.GetFullPath(input);
                    var dependencyCandidates = new List<string> { fullInputPath };
                    try {
                        foreach (string file in Directory.EnumerateFiles(sourceDirectory, "*", new EnumerationOptions {
                                     RecurseSubdirectories = true,
                                     IgnoreInaccessible = false,
                                     AttributesToSkip = FileAttributes.ReparsePoint,
                                     ReturnSpecialDirectories = false
                                 })) {
                            cancellationToken.ThrowIfCancellationRequested();
                            string fullPath = Path.GetFullPath(file);
                            if (pathComparer.Equals(fullPath, fullInputPath)) continue;
                            CountDiscoveredEntry();
                            if (OfficeWorkflowHtmlResourceResolver.IsSupportedDependency(fullPath)) dependencyCandidates.Add(fullPath);
                        }
                    } catch (UnauthorizedAccessException ex) {
                        throw new IOException("An HTML source folder could not be enumerated safely: " + sourceDirectory, ex);
                    }
                    HtmlDependencyDiscovery dependencyDiscovery = FindReferencedHtmlDependencies(
                        dependencyCandidates,
                        physicalRoot,
                        request.Limits.MaximumInputBytes,
                        pathComparer,
                        cancellationToken);
                    AddInputBytes(dependencyDiscovery.TotalBytes);
                    dependencySnapshots = dependencyDiscovery.Snapshots;
                }
                AddDiscoveredSource(input, input, discovered: false, physicalRoot, dependencySnapshots);
            }
        }
        if (sources.Count == 0) throw new InvalidOperationException("No supported documents were found in the requested sources.");
        inputBytes = expandedInputBytes;
        return sources;

        void AddDiscoveredSource(
            string path,
            string origin,
            bool discovered,
            string physicalRoot,
            IReadOnlyDictionary<string, byte[]>? dependencySnapshots = null) {
            cancellationToken.ThrowIfCancellationRequested();
            string extension = Path.GetExtension(path);
            if (string.Equals(extension, ".zip", StringComparison.OrdinalIgnoreCase)) {
                archiveIndex++;
                ExpandArchive(path, origin, archiveIndex, physicalRoot);
                return;
            }
            if (!TryClassifyAssemblySource(path, out AssemblySourceKind kind, out OfficeWorkflowRoute? route)) {
                if (discovered && request.Options.IgnoreDiscoveredUnsupportedFiles) return;
                throw new NotSupportedException("No PDF assembly intake route is available for '" + path + "'.");
            }
            Add(CaptureSource(path, origin, Path.GetFileName(path), kind, route, physicalRoot, dependencySnapshots));
        }

        void ExpandArchive(string archivePath, string origin, int index, string physicalRoot) {
            using FileStream archiveStream = OfficeWorkflowPathIdentity.OpenRegularFileForRead(
                archivePath,
                physicalRoot,
                AssemblyInputBufferSize);
            long archiveFileBytes = archiveStream.Length;
            EnforceInputLimit(archivePath, archiveFileBytes, request.Limits);
            string destinationRoot = Path.Combine(extractionRoot, "archive-" + index.ToString("D4", System.Globalization.CultureInfo.InvariantCulture));
            Directory.CreateDirectory(destinationRoot);
            string physicalDestinationRoot = OfficeWorkflowPathIdentity.ResolvePhysicalPath(destinationRoot);
            string canonicalDestinationRoot = Path.GetFullPath(destinationRoot)
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
            StringComparison destinationComparison = OfficeWorkflowPathIdentity.GetComparison(destinationRoot);
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

            var extractedFiles = new List<string>();
            var archiveSources = new List<AssemblySource>();
            var archiveDependencyOnlyEntries = new List<(string Path, string DisplayName)>();
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
                string normalizedName = entry.FullName.Replace('/', Path.DirectorySeparatorChar).Replace('\\', Path.DirectorySeparatorChar);
                bool isAssemblySource = TryClassifyAssemblySource(entry.Name, out AssemblySourceKind kind, out OfficeWorkflowRoute? route);
                bool isPotentialHtmlDependency = OfficeWorkflowHtmlResourceResolver.IsSupportedDependency(entry.Name);
                if (!isAssemblySource && !isPotentialHtmlDependency && !request.Options.IgnoreDiscoveredUnsupportedFiles) {
                    throw new NotSupportedException("No PDF assembly intake route is available for archive entry '" + entry.FullName + "'.");
                }
                if (!isAssemblySource && !isPotentialHtmlDependency) continue;
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

                string destination = Path.GetFullPath(Path.Combine(destinationRoot, normalizedName));
                if (!destination.StartsWith(canonicalDestinationRoot, destinationComparison)) {
                    throw new InvalidDataException("An archive entry resolves outside its extraction directory.");
                }
                Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
                using (Stream source = entry.Open())
                using (FileStream target = new(destination, FileMode.CreateNew, FileAccess.Write, FileShare.None)) {
                    CopyBounded(source, target, entry.Length, request.Options.MaximumArchiveEntryBytes, cancellationToken);
                }
                extractedFiles.Add(destination);
                if (isAssemblySource) {
                    archiveSources.Add(CaptureSource(destination, origin, entry.FullName, kind, route, physicalDestinationRoot));
                } else {
                    archiveDependencyOnlyEntries.Add((destination, entry.FullName));
                }
            }
            if (observedArchiveEntries != preflight.EntryCount) {
                throw new InvalidDataException("The ZIP produced fewer entries than its bounded central-directory preflight declared.");
            }
            HtmlDependencyDiscovery dependencyDiscovery = FindReferencedHtmlDependencies(
                extractedFiles,
                physicalDestinationRoot,
                request.Limits.MaximumInputBytes,
                OfficeWorkflowPathIdentity.GetComparer(destinationRoot),
                cancellationToken);
            AddInputBytes(dependencyDiscovery.TotalBytes);
            if (!request.Options.IgnoreDiscoveredUnsupportedFiles) {
                foreach ((string path, string displayName) in archiveDependencyOnlyEntries) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!dependencyDiscovery.Paths.Contains(Path.GetFullPath(path))) {
                        throw new NotSupportedException("No PDF assembly intake route is available for archive entry '" + displayName + "'.");
                    }
                }
            }
            foreach (AssemblySource source in archiveSources) {
                if (!dependencyDiscovery.Paths.Contains(Path.GetFullPath(source.Path))) {
                    Add(string.Equals(source.Route?.Id, "html-pdf", StringComparison.Ordinal)
                        ? source with { HtmlDependencySnapshots = dependencyDiscovery.Snapshots }
                        : source);
                }
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
            AddInputBytes(source.SizeBytes);
            sources.Add(source);
        }

        void AddInputBytes(long bytes) {
            expandedInputBytes = checked(expandedInputBytes + bytes);
            if (expandedInputBytes > request.Limits.MaximumInputBytes) {
                throw new InvalidOperationException(
                    $"Expanded inputs total {expandedInputBytes:N0} bytes, above the configured {request.Limits.MaximumInputBytes:N0}-byte limit.");
            }
        }

        AssemblySource CaptureSource(
            string path,
            string origin,
            string displayName,
            AssemblySourceKind kind,
            OfficeWorkflowRoute? route,
            string physicalRoot,
            IReadOnlyDictionary<string, byte[]>? dependencySnapshots = null) {
            using FileStream stream = OfficeWorkflowPathIdentity.OpenRegularFileForRead(
                path,
                physicalRoot,
                AssemblyInputBufferSize);
            long size = stream.Length;
            EnforceInputLimit(path, size, request.Limits);
            string identity = OfficeWorkflowPathIdentity.GetPhysicalIdentityKey(path, stream);
            string contentSha256 = ComputeAssemblySourceSha256(stream, cancellationToken);
            IReadOnlyDictionary<string, byte[]>? htmlDependencySnapshots =
                string.Equals(route?.Id, "html-pdf", StringComparison.Ordinal) ? dependencySnapshots : null;
            return new AssemblySource(path, origin, displayName, kind, route, size, physicalRoot, identity, contentSha256, htmlDependencySnapshots);
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

    private static HtmlDependencyDiscovery FindReferencedHtmlDependencies(
        IReadOnlyCollection<string> files,
        string physicalRoot,
        long maximumInputBytes,
        StringComparer pathComparer,
        CancellationToken cancellationToken) {
        var candidates = new Dictionary<string, string>(pathComparer);
        var htmlFiles = new List<string>();
        foreach (string file in files) {
            cancellationToken.ThrowIfCancellationRequested();
            string fullPath = Path.GetFullPath(file);
            if (OfficeWorkflowHtmlResourceResolver.IsSupportedDependency(fullPath)) {
                candidates[fullPath] = fullPath;
            }
            string extension = Path.GetExtension(fullPath);
            if (string.Equals(extension, ".html", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(extension, ".htm", StringComparison.OrdinalIgnoreCase)) {
                htmlFiles.Add(fullPath);
            }
        }

        var referenced = new HashSet<string>(pathComparer);
        var snapshots = new Dictionary<string, byte[]>(pathComparer);
        long referencedBytes = 0L;
        var pendingStylesheets = new Queue<string>();
        var processedStylesheets = new HashSet<string>(pathComparer);
        foreach (string htmlPath in htmlFiles) {
            cancellationToken.ThrowIfCancellationRequested();
            byte[] htmlBytes = ReadDependencyBytes(htmlPath);
            AddManifestReferences(
                HtmlResourcePipeline.BuildManifest(
                    DecodeHtmlInput(htmlBytes),
                    OfficeWorkflowHtmlResourceResolver.CreatePdfResourcePipelineOptions(new Uri(htmlPath))),
                pendingStylesheets);
        }
        while (pendingStylesheets.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            string stylesheetPath = pendingStylesheets.Dequeue();
            if (!processedStylesheets.Add(stylesheetPath)) continue;
            byte[] cssBytes = snapshots[stylesheetPath];
            using var source = new MemoryStream(cssBytes, writable: false);
            using var reader = new StreamReader(source, System.Text.Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
            HtmlResourceManifest manifest = HtmlResourcePipeline.BuildStylesheetManifest(
                reader.ReadToEnd(),
                new Uri(stylesheetPath),
                OfficeWorkflowHtmlResourceResolver.CreatePdfResourcePipelineOptions());
            AddManifestReferences(manifest, pendingStylesheets);
        }
        return new HtmlDependencyDiscovery(referenced, snapshots, referencedBytes);

        byte[] ReadDependencyBytes(string path) {
            using FileStream stream = OfficeWorkflowPathIdentity.OpenRegularFileForRead(
                path,
                physicalRoot,
                AssemblyInputBufferSize);
            return OfficeWorkflowInputReader.ReadAllBytes(
                stream,
                Path.GetFileName(path),
                maximumInputBytes,
                cancellationToken);
        }

        void AddManifestReferences(HtmlResourceManifest manifest, Queue<string> stylesheets) {
            foreach (HtmlResourceReference resource in manifest.Resources) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!resource.IsAllowed || !IsLoadableHtmlDependency(resource.Kind) ||
                    !Uri.TryCreate(resource.ResolvedSource, UriKind.Absolute, out Uri? uri) ||
                    !uri.IsFile) {
                    continue;
                }
                string fullPath = Path.GetFullPath(uri.LocalPath);
                if (!candidates.TryGetValue(fullPath, out string? candidate)) continue;
                if (referenced.Add(candidate)) {
                    byte[] snapshot = ReadDependencyBytes(candidate);
                    referencedBytes = checked(referencedBytes + snapshot.LongLength);
                    if (referencedBytes > maximumInputBytes) {
                        throw new InvalidOperationException(
                            $"Referenced HTML dependencies total {referencedBytes:N0} bytes, above the configured {maximumInputBytes:N0}-byte limit.");
                    }
                    snapshots[candidate] = snapshot;
                }
                if (resource.Kind == HtmlResourceKind.Stylesheet &&
                    string.Equals(Path.GetExtension(candidate), ".css", StringComparison.OrdinalIgnoreCase)) {
                    stylesheets.Enqueue(candidate);
                }
            }
        }

        static bool IsLoadableHtmlDependency(HtmlResourceKind kind) =>
            kind is HtmlResourceKind.Image or HtmlResourceKind.Stylesheet or HtmlResourceKind.Font;
    }

    private sealed record HtmlDependencyDiscovery(
        HashSet<string> Paths,
        IReadOnlyDictionary<string, byte[]> Snapshots,
        long TotalBytes);

    private static bool IsAssemblyOutputCandidate(
        string path,
        string requestedOutputPath,
        OfficeWorkflowConflictPolicy conflictPolicy) {
        if (OfficeWorkflowPathIdentity.AreEquivalent(path, requestedOutputPath)) return true;
        if (conflictPolicy != OfficeWorkflowConflictPolicy.Rename) return false;

        string fullPath = Path.GetFullPath(path);
        string outputPath = Path.GetFullPath(requestedOutputPath);
        string? parent = Path.GetDirectoryName(outputPath);
        if (parent is null || !OfficeWorkflowPathIdentity.AreEquivalent(Path.GetDirectoryName(fullPath)!, parent)) return false;
        if (!string.Equals(Path.GetExtension(fullPath), Path.GetExtension(outputPath), OfficeWorkflowPathIdentity.GetComparison(parent))) return false;

        string outputStem = Path.GetFileNameWithoutExtension(outputPath);
        string candidateStem = Path.GetFileNameWithoutExtension(fullPath);
        if (!candidateStem.StartsWith(outputStem + " (", OfficeWorkflowPathIdentity.GetComparison(parent)) ||
            !candidateStem.EndsWith(")", StringComparison.Ordinal)) {
            return false;
        }
        ReadOnlySpan<char> suffix = candidateStem.AsSpan(outputStem.Length + 2, candidateStem.Length - outputStem.Length - 3);
        return int.TryParse(
            suffix,
            System.Globalization.NumberStyles.None,
            System.Globalization.CultureInfo.InvariantCulture,
            out int value) && value > 0;
    }

    private static PdfDocument NormalizeAssemblySource(
        AssemblySource source,
        ValidatedAssemblyRequest request,
        long maximumNormalizedBytes,
        List<OfficeWorkflowDiagnostic> diagnostics,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        using FileStream inputStream = OfficeWorkflowPathIdentity.OpenRegularFileForRead(
            source.Path,
            source.PhysicalRoot,
            AssemblyInputBufferSize);
        string observedIdentity = OfficeWorkflowPathIdentity.GetPhysicalIdentityKey(source.Path, inputStream);
        if (!string.Equals(observedIdentity, source.PhysicalIdentityKey, StringComparison.Ordinal) ||
            inputStream.Length != source.SizeBytes) {
            throw new InvalidDataException("An assembly source changed after discovery: " + source.DisplayName);
        }
        byte[] input = OfficeWorkflowInputReader.ReadAllBytes(
            inputStream,
            source.DisplayName,
            request.Limits.MaximumInputBytes,
            cancellationToken);
        string observedContentSha256 = System.Convert.ToHexString(SHA256.HashData(input));
        if (!string.Equals(observedContentSha256, source.ContentSha256, StringComparison.Ordinal)) {
            throw new InvalidDataException("An assembly source changed after discovery: " + source.DisplayName);
        }
        switch (source.Kind) {
            case AssemblySourceKind.Pdf:
                PdfDocument opened = PdfDocument.Load(input, request.PdfLoadOptions);
                _ = opened.Inspect(request.PdfLoadOptions, cancellationToken);
                AddAssemblySourceDiagnostic(source, "PDF pages retained", diagnostics);
                return opened;
            case AssemblySourceKind.Image:
                PdfDocument imageDocument = PdfDocument.CreateFromImages(
                    [new PdfImageDocumentSource(input, source.DisplayName)],
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
                    request.PdfLoadOptions,
                    CreatePdfLoadOptions(request.PdfLoadOptions.Password, maximumNormalizedBytes));
                OperationArtifact artifact = Convert(
                    conversionRequest,
                    input,
                    diagnostics,
                    cancellationToken,
                    emitHtmlTaggedStructure: source.Route?.Id != "html-pdf",
                    htmlResourceSnapshots: source.HtmlDependencySnapshots);
                if (artifact.Bytes == null) throw new InvalidOperationException("An Office input did not produce PDF bytes.");
                AddAssemblySourceDiagnostic(source, "Office document normalized to PDF", diagnostics);
                return PdfDocument.Load(artifact.Bytes, conversionRequest.OutputPdfLoadOptions);
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

    private static string ComputeAssemblySourceSha256(FileStream stream, CancellationToken cancellationToken) {
        stream.Position = 0L;
        using IncrementalHash hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        var buffer = new byte[AssemblyInputBufferSize];
        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            int read = stream.Read(buffer, 0, buffer.Length);
            if (read == 0) break;
            hash.AppendData(buffer, 0, read);
        }
        cancellationToken.ThrowIfCancellationRequested();
        return System.Convert.ToHexString(hash.GetHashAndReset());
    }

    private sealed record AssemblySource(
        string Path,
        string Origin,
        string DisplayName,
        AssemblySourceKind Kind,
        OfficeWorkflowRoute? Route,
        long SizeBytes,
        string PhysicalRoot,
        string PhysicalIdentityKey,
        string ContentSha256,
        IReadOnlyDictionary<string, byte[]>? HtmlDependencySnapshots);

    private sealed record ValidatedAssemblyRequest(
        string Id,
        IReadOnlyList<string> Sources,
        string OutputPath,
        OfficeWorkflowConflictPolicy ConflictPolicy,
        OfficeWorkflowOutputProfile OutputProfile,
        PdfAssemblyOptions Options,
        OfficeWorkflowLimits Limits,
        PdfLoadOptions PdfLoadOptions,
        PdfLoadOptions OutputPdfLoadOptions);
}
