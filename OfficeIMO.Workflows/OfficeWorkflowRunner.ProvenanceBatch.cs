using OfficeIMO.Provenance;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    /// <summary>Runs a bounded batch sequentially so providers and parsers share one predictable resource envelope.</summary>
    public async Task<IReadOnlyList<OfficeProvenanceWorkflowResult>> RunProvenanceBatchAsync(
        IEnumerable<OfficeProvenanceWorkflowRequest> requests,
        OfficeProvenanceWorkflowBatchOptions? options = null,
        IProgress<OfficeWorkflowProgress>? progress = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(requests);
        OfficeProvenanceWorkflowBatchOptions validatedOptions = (options ?? new OfficeProvenanceWorkflowBatchOptions()).CloneAndValidate();
        OfficeProvenanceWorkflowRequest[] materialized = MaterializeBatchRequests(
            requests,
            validatedOptions.MaximumRequests,
            cancellationToken);
        if (materialized.Length > validatedOptions.MaximumRequests) {
            throw new ArgumentException(
                $"The provenance batch exceeds the configured limit of {validatedOptions.MaximumRequests:N0} requests.",
                nameof(requests));
        }
        materialized = PrepareBatchRemovalPaths(materialized, cancellationToken);

        var results = new List<OfficeProvenanceWorkflowResult>(materialized.Length);
        for (int index = 0; index < materialized.Length; index++) {
            int batchIndex = index;
            var batchProgress = progress is null
                ? null
                : new InlineProgress<OfficeWorkflowProgress>(item => progress.Report(new OfficeWorkflowProgress(
                    item.RequestId,
                    item.Stage,
                    $"{batchIndex + 1} of {materialized.Length} · {item.Message}",
                    item.Fraction,
                    (batchIndex + item.Fraction) / Math.Max(1, materialized.Length))));
            OfficeProvenanceWorkflowResult result = await RunProvenanceAsync(
                materialized[index], batchProgress, cancellationToken).ConfigureAwait(false);
            results.Add(result);
            if (result.Status == OfficeWorkflowStatus.Cancelled) break;
            if (!validatedOptions.ContinueOnFailure && !result.Succeeded) break;
        }
        return results;
    }

    private static OfficeProvenanceWorkflowRequest[] MaterializeBatchRequests(
        IEnumerable<OfficeProvenanceWorkflowRequest> requests,
        int maximumRequests,
        CancellationToken cancellationToken) {
        if (cancellationToken.IsCancellationRequested) return Array.Empty<OfficeProvenanceWorkflowRequest>();
        var materialized = new List<OfficeProvenanceWorkflowRequest>(maximumRequests + 1);
        using IEnumerator<OfficeProvenanceWorkflowRequest> enumerator = requests.GetEnumerator();
        while (materialized.Count <= maximumRequests) {
            if (cancellationToken.IsCancellationRequested) break;
            if (!enumerator.MoveNext()) break;
            materialized.Add(enumerator.Current);
            if (cancellationToken.IsCancellationRequested) break;
        }
        return materialized.ToArray();
    }

    internal static OfficeProvenanceWorkflowRequest[] PrepareBatchRemovalPaths(
        IReadOnlyList<OfficeProvenanceWorkflowRequest> requests,
        CancellationToken cancellationToken = default) {
        var prepared = new OfficeProvenanceWorkflowRequest[requests.Count];
        for (int index = 0; index < requests.Count; index++) {
            OfficeProvenanceWorkflowRequest request = requests[index] ??
                throw new ArgumentException("Provenance batches cannot contain null requests.", nameof(requests));
            prepared[index] = CloneBatchRequest(request);
        }
        if (cancellationToken.IsCancellationRequested) return prepared;
        if (!prepared.Any(static request => request.Operation == OfficeProvenanceWorkflowOperation.Remove)) {
            return prepared;
        }

        var pathIndex = new BatchPathIndex(cancellationToken);
        var inputRequestIndexes = new Dictionary<string, List<int>>(StringComparer.Ordinal);
        for (int index = 0; index < prepared.Length; index++) {
            if (cancellationToken.IsCancellationRequested) return prepared;
            OfficeProvenanceWorkflowRequest request = prepared[index];
            string? identity = pathIndex.TryNormalize(request.InputPath);
            if (identity is null) continue;
            if (!inputRequestIndexes.TryGetValue(identity, out List<int>? indexes)) {
                indexes = new List<int>();
                inputRequestIndexes.Add(identity, indexes);
            }
            indexes.Add(index);
        }
        var inputIdentities = new SortedSet<string>(inputRequestIndexes.Keys, StringComparer.Ordinal);

        var removalOutputs = new BatchOutputPath?[prepared.Length];
        var outputIdentities = new SortedSet<string>(StringComparer.Ordinal);
        var outputPathsByIdentity = new Dictionary<string, string>(StringComparer.Ordinal);
        for (int index = 0; index < prepared.Length; index++) {
            if (cancellationToken.IsCancellationRequested) return prepared;
            OfficeProvenanceWorkflowRequest request = prepared[index];
            if (request.Operation != OfficeProvenanceWorkflowOperation.Remove) continue;
            string? outputPath = TryResolveBatchRemovalOutput(request);
            string? identity = pathIndex.TryNormalize(outputPath);
            if (outputPath is null || identity is null) continue;
            EnsureBatchOutputDoesNotOverlapInput(
                outputPath,
                identity,
                index,
                inputRequestIndexes,
                inputIdentities);
            if (!outputIdentities.Contains(identity) &&
                TryFindAncestorOrDescendant(identity, outputIdentities, out string? collisionIdentity)) {
                throw new ArgumentException(
                    $"Provenance removal outputs '{outputPathsByIdentity[collisionIdentity!]}' and '{outputPath}' have an ancestor/descendant path collision.",
                    "requests");
            }
            outputIdentities.Add(identity);
            outputPathsByIdentity.TryAdd(identity, outputPath);
            removalOutputs[index] = new BatchOutputPath(outputPath, identity);
        }

        var fixedOutputIdentities = new HashSet<string>(StringComparer.Ordinal);
        for (int index = 0; index < prepared.Length; index++) {
            if (cancellationToken.IsCancellationRequested) return prepared;
            OfficeProvenanceWorkflowRequest request = prepared[index];
            if (request.Operation != OfficeProvenanceWorkflowOperation.Remove ||
                request.ConflictPolicy == OfficeWorkflowConflictPolicy.Rename) continue;
            BatchOutputPath? output = removalOutputs[index];
            if (output is null) continue;
            EnsureBatchOutputIsUnique(output.Path, output.Identity, fixedOutputIdentities);
        }

        var reservedOutputIdentities = new HashSet<string>(fixedOutputIdentities, StringComparer.Ordinal);
        var nextSuffixByDestination = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int index = 0; index < prepared.Length; index++) {
            if (cancellationToken.IsCancellationRequested) return prepared;
            OfficeProvenanceWorkflowRequest request = prepared[index];
            if (request.Operation != OfficeProvenanceWorkflowOperation.Remove ||
                request.ConflictPolicy != OfficeWorkflowConflictPolicy.Rename) continue;
            BatchOutputPath? output = removalOutputs[index];
            if (output is null) continue;
            string requestedPath = output.Path;
            string requestedIdentity = output.Identity;
            nextSuffixByDestination.TryGetValue(requestedIdentity, out int nextSuffix);

            string? selectedPath = SelectBatchRenameOutput(
                requestedPath,
                inputRequestIndexes,
                inputIdentities,
                reservedOutputIdentities,
                pathIndex,
                ref nextSuffix,
                cancellationToken);
            if (selectedPath is null) continue;
            nextSuffixByDestination[requestedIdentity] = nextSuffix;
            string selectedIdentity = pathIndex.NormalizeCandidate(selectedPath);
            reservedOutputIdentities.Add(selectedIdentity);
            request.OutputPath = selectedPath;
        }
        var blockedOutputIdentities = new SortedSet<string>(inputIdentities, StringComparer.Ordinal);
        blockedOutputIdentities.UnionWith(reservedOutputIdentities);
        for (int index = 0; index < prepared.Length; index++) {
            if (cancellationToken.IsCancellationRequested) return prepared;
            OfficeProvenanceWorkflowRequest request = prepared[index];
            if (request.Operation != OfficeProvenanceWorkflowOperation.Remove ||
                request.ConflictPolicy != OfficeWorkflowConflictPolicy.Rename ||
                string.IsNullOrWhiteSpace(request.OutputPath)) continue;
            request.BatchBlockedOutputIdentities = blockedOutputIdentities;
            request.BatchOwnReservedOutputIdentity = pathIndex.NormalizeCandidate(request.OutputPath);
        }
        return prepared;
    }

    private static void EnsureBatchOutputDoesNotOverlapInput(
        string outputPath,
        string identity,
        int requestIndex,
        IReadOnlyDictionary<string, List<int>> inputRequestIndexes,
        SortedSet<string> inputIdentities) {
        if ((inputRequestIndexes.TryGetValue(identity, out List<int>? indexes) &&
             indexes.Any(index => index != requestIndex)) ||
            TryFindAncestorOrDescendant(identity, inputIdentities, out _)) {
            throw new ArgumentException(
                $"Provenance removal output '{outputPath}' overlaps another batch request's input path.",
                "requests");
        }
    }

    private static void EnsureBatchOutputIsUnique(
        string outputPath,
        string identity,
        ISet<string> outputIdentities) {
        if (!outputIdentities.Add(identity)) {
            throw new ArgumentException(
                $"Multiple provenance removal requests resolve to the same output path '{outputPath}'.",
                "requests");
        }
    }

    private static string? SelectBatchRenameOutput(
        string requestedPath,
        IReadOnlyDictionary<string, List<int>> inputRequestIndexes,
        SortedSet<string> inputIdentities,
        IReadOnlySet<string> reservedOutputIdentities,
        BatchPathIndex pathIndex,
        ref int nextSuffix,
        CancellationToken cancellationToken) {
        for (int attempts = 0; attempts < 10_000; attempts++) {
            if (cancellationToken.IsCancellationRequested) return null;
            int suffix = nextSuffix;
            if (nextSuffix == int.MaxValue) {
                throw new IOException("No available numbered provenance output path could be reserved for the batch.");
            }
            nextSuffix++;
            string candidate = suffix == 0 ? requestedPath : AddSuffix(requestedPath, suffix);
            string identity = pathIndex.NormalizeCandidate(candidate);
            if (inputRequestIndexes.ContainsKey(identity) ||
                TryFindAncestorOrDescendant(identity, inputIdentities, out _) ||
                reservedOutputIdentities.Contains(identity) ||
                pathIndex.CandidateExists(candidate)) continue;
            return candidate;
        }
        throw new IOException("No available numbered provenance output path could be reserved for the batch.");
    }

    private static OfficeProvenanceWorkflowRequest CloneBatchRequest(OfficeProvenanceWorkflowRequest source) => new() {
        Id = source.Id,
        Operation = source.Operation,
        InputPath = source.InputPath,
        OutputPath = source.OutputPath,
        ConflictPolicy = source.ConflictPolicy,
        Inspection = CloneBatchInspectionOptions(source.Inspection)!,
        Assessment = CloneBatchAssessmentOptions(source.Assessment)!,
        Removal = CloneBatchRemovalOptions(source.Removal)!,
        Limits = source.Limits is null
            ? null!
            : new OfficeWorkflowLimits {
                MaximumInputBytes = source.Limits.MaximumInputBytes,
                MaximumOutputBytes = source.Limits.MaximumOutputBytes
            }
    };

    private static OfficeProvenanceOptions? CloneBatchInspectionOptions(OfficeProvenanceOptions? source) {
        if (source is null) return null;
        var clone = new OfficeProvenanceOptions();
        CopyBatchInspectionOptions(source, clone);
        return clone;
    }

    private static OfficeProvenanceAssessmentOptions? CloneBatchAssessmentOptions(OfficeProvenanceAssessmentOptions? source) {
        if (source is null) return null;
        var clone = new OfficeProvenanceAssessmentOptions {
            InspectTextIntegrity = source.InspectTextIntegrity
        };
        CopyBatchInspectionOptions(source.Structural, clone.Structural);
        clone.TextIntegrity.MaxEncodedBytes = source.TextIntegrity.MaxEncodedBytes;
        clone.TextIntegrity.MaxCharacters = source.TextIntegrity.MaxCharacters;
        clone.TextIntegrity.MaxFindings = source.TextIntegrity.MaxFindings;
        clone.TextIntegrity.IgnoreLeadingByteOrderMark = source.TextIntegrity.IgnoreLeadingByteOrderMark;
        clone.TextIntegrity.IncludeTypographicSpaces = source.TextIntegrity.IncludeTypographicSpaces;
        clone.TextIntegrity.IncludeVariationSelectors = source.TextIntegrity.IncludeVariationSelectors;
        clone.Verification.Timeout = source.Verification.Timeout;
        clone.Verification.MaxReportBytes = source.Verification.MaxReportBytes;
        clone.Verification.AllowNetworkAccess = source.Verification.AllowNetworkAccess;
        clone.Verification.IncludeRawReport = source.Verification.IncludeRawReport;
        clone.Verification.TrustAnchorsPath = source.Verification.TrustAnchorsPath;
        clone.Verification.AllowedListPath = source.Verification.AllowedListPath;
        clone.Verification.TrustConfigurationPath = source.Verification.TrustConfigurationPath;
        return clone;
    }

    private static OfficeProvenanceRemovalOptions? CloneBatchRemovalOptions(OfficeProvenanceRemovalOptions? source) {
        if (source is null) return null;
        var clone = new OfficeProvenanceRemovalOptions {
            RemoveC2paManifests = source.RemoveC2paManifests,
            RemoveExternalC2paReferences = source.RemoveExternalC2paReferences,
            RemoveAiSourceMetadata = source.RemoveAiSourceMetadata,
            RequireStructurallyValidCarrier = source.RequireStructurallyValidCarrier,
            SignatureMutationPolicy = source.SignatureMutationPolicy,
            ProcessEmbeddedAssets = source.ProcessEmbeddedAssets,
            MaxEmbeddedAssets = source.MaxEmbeddedAssets,
            MaxOutputBytes = source.MaxOutputBytes
        };
        CopyBatchInspectionOptions(source.Limits, clone.Limits);
        return clone;
    }

    private static void CopyBatchInspectionOptions(OfficeProvenanceOptions source, OfficeProvenanceOptions destination) {
        destination.MaxAssetBytes = source.MaxAssetBytes;
        destination.MaxManifestBytes = source.MaxManifestBytes;
        destination.MaxCarriers = source.MaxCarriers;
        destination.MaxContainerEntries = source.MaxContainerEntries;
        destination.MaxExpandedContainerBytes = source.MaxExpandedContainerBytes;
        destination.ProcessEmbeddedAssets = source.ProcessEmbeddedAssets;
        destination.MaxEmbeddedAssets = source.MaxEmbeddedAssets;
    }

    private static bool TryFindAncestorOrDescendant(
        string identity,
        SortedSet<string> identities,
        out string? collisionIdentity) {
        string? parent = Path.GetDirectoryName(identity);
        while (!string.IsNullOrEmpty(parent)) {
            if (identities.Contains(parent)) {
                collisionIdentity = parent;
                return true;
            }
            string? next = Path.GetDirectoryName(parent);
            if (string.Equals(next, parent, StringComparison.Ordinal)) break;
            parent = next;
        }

        string prefix = identity.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) +
                        Path.DirectorySeparatorChar;
        foreach (string candidate in identities.GetViewBetween(prefix, prefix + '\uffff')) {
            collisionIdentity = candidate;
            return true;
        }
        collisionIdentity = null;
        return false;
    }

    private static string? TryResolveBatchRemovalOutput(OfficeProvenanceWorkflowRequest request) {
        try {
            if (string.IsNullOrWhiteSpace(request.InputPath)) return null;
            string inputPath = Path.GetFullPath(request.InputPath);
            return string.IsNullOrWhiteSpace(request.OutputPath)
                ? Path.Combine(
                    Path.GetDirectoryName(inputPath)!,
                    Path.GetFileNameWithoutExtension(inputPath) + ".provenance-cleaned" + Path.GetExtension(inputPath))
                : Path.GetFullPath(request.OutputPath);
        } catch (Exception exception) when (exception is ArgumentException or NotSupportedException or IOException or UnauthorizedAccessException) {
            // RunProvenanceAsync owns request validation and converts invalid paths into per-item results.
            return null;
        }
    }

    private sealed class BatchPathIndex(CancellationToken cancellationToken) {
        private readonly Dictionary<string, string?> _normalizedPaths = new(StringComparer.Ordinal);
        private readonly Dictionary<string, BatchDirectoryIndex> _lexicalDirectories = new(StringComparer.Ordinal);
        private readonly Dictionary<string, BatchDirectoryIndex> _physicalDirectories = new(StringComparer.Ordinal);

        internal string? TryNormalize(string? path) {
            if (string.IsNullOrWhiteSpace(path)) return null;
            string fullPath;
            try {
                fullPath = Path.GetFullPath(path);
            } catch (Exception exception) when (exception is ArgumentException or NotSupportedException) {
                return null;
            } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
                throw new IOException($"Unable to resolve provenance batch path '{path}'.", exception);
            }
            if (_normalizedPaths.TryGetValue(fullPath, out string? cached)) return cached;
            try {
                string identity = OfficeWorkflowPathIdentity.Normalize(fullPath);
                _normalizedPaths.Add(fullPath, identity);
                return identity;
            } catch (ArgumentException) {
                _normalizedPaths.Add(fullPath, null);
                return null;
            } catch (Exception exception) when (exception is NotSupportedException or IOException or UnauthorizedAccessException) {
                throw new IOException($"Unable to resolve provenance batch path identity '{path}'.", exception);
            }
        }

        internal string NormalizeCandidate(string path) {
            string fullPath = Path.GetFullPath(path);
            if (_normalizedPaths.TryGetValue(fullPath, out string? cached) && cached is not null) return cached;
            try {
                BatchDirectoryIndex directory = GetDirectory(fullPath);
                string identity = directory.NormalizeChild(Path.GetFileName(fullPath));
                _normalizedPaths[fullPath] = identity;
                return identity;
            } catch (Exception exception) when (exception is ArgumentException or NotSupportedException or IOException or UnauthorizedAccessException) {
                throw new IOException($"Unable to reserve provenance batch output '{path}'.", exception);
            }
        }

        internal bool CandidateExists(string path) {
            try {
                string fullPath = Path.GetFullPath(path);
                _ = GetDirectory(fullPath);
                _ = File.GetAttributes(fullPath);
                return true;
            } catch (FileNotFoundException) {
                return false;
            } catch (DirectoryNotFoundException) {
                return false;
            } catch (Exception exception) when (exception is ArgumentException or NotSupportedException or IOException or UnauthorizedAccessException) {
                throw new IOException($"Unable to inspect provenance batch output candidate '{path}'.", exception);
            }
        }

        private BatchDirectoryIndex GetDirectory(string fullChildPath) {
            string? directoryPath = Path.GetDirectoryName(fullChildPath);
            if (string.IsNullOrEmpty(directoryPath)) throw new IOException("The provenance batch output has no destination directory.");
            if (_lexicalDirectories.TryGetValue(directoryPath, out BatchDirectoryIndex? cached)) return cached;
            try {
                cancellationToken.ThrowIfCancellationRequested();
                string physicalDirectory = OfficeWorkflowPathIdentity.ResolvePhysicalPath(directoryPath);
                bool caseInsensitive = OfficeWorkflowPathIdentity.IsCaseInsensitiveFileSystem(physicalDirectory);
                string physicalIdentity = OfficeWorkflowPathIdentity.Normalize(physicalDirectory, caseInsensitive);
                if (_physicalDirectories.TryGetValue(physicalIdentity, out BatchDirectoryIndex? shared)) {
                    _lexicalDirectories.Add(directoryPath, shared);
                    return shared;
                }
                var index = new BatchDirectoryIndex(physicalDirectory, caseInsensitive);
                _physicalDirectories.Add(physicalIdentity, index);
                _lexicalDirectories.Add(directoryPath, index);
                return index;
            } catch (Exception exception) when (exception is ArgumentException or NotSupportedException or IOException or UnauthorizedAccessException) {
                throw new IOException($"Unable to index provenance batch destination directory '{directoryPath}'.", exception);
            }
        }
    }

    private sealed class BatchDirectoryIndex(
        string physicalDirectory,
        bool caseInsensitive) {
        internal string NormalizeChild(string fileName) => OfficeWorkflowPathIdentity.Normalize(
            Path.Combine(physicalDirectory, fileName),
            caseInsensitive);
    }

    private sealed record BatchOutputPath(string Path, string Identity);
}
