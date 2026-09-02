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
        OfficeProvenanceWorkflowRequest[] materialized = requests.Take(validatedOptions.MaximumRequests + 1).ToArray();
        if (materialized.Length > validatedOptions.MaximumRequests) {
            throw new ArgumentException(
                $"The provenance batch exceeds the configured limit of {validatedOptions.MaximumRequests:N0} requests.",
                nameof(requests));
        }
        if (!cancellationToken.IsCancellationRequested) {
            materialized = PrepareBatchRemovalPaths(materialized, cancellationToken);
        }
        if (materialized.Length == 0) cancellationToken.ThrowIfCancellationRequested();

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

    internal static OfficeProvenanceWorkflowRequest[] PrepareBatchRemovalPaths(
        IReadOnlyList<OfficeProvenanceWorkflowRequest> requests,
        CancellationToken cancellationToken = default) {
        var prepared = new OfficeProvenanceWorkflowRequest[requests.Count];
        for (int index = 0; index < requests.Count; index++) {
            prepared[index] = requests[index] ??
                throw new ArgumentException("Provenance batches cannot contain null requests.", nameof(requests));
        }
        if (cancellationToken.IsCancellationRequested) return prepared;

        var pathIndex = new BatchPathIndex(cancellationToken);
        var inputRequestIndexes = new Dictionary<string, List<int>>(StringComparer.Ordinal);
        for (int index = 0; index < requests.Count; index++) {
            if (cancellationToken.IsCancellationRequested) return prepared;
            OfficeProvenanceWorkflowRequest request = requests[index];
            string? identity = pathIndex.TryNormalize(request.InputPath);
            if (identity is null) continue;
            if (!inputRequestIndexes.TryGetValue(identity, out List<int>? indexes)) {
                indexes = new List<int>();
                inputRequestIndexes.Add(identity, indexes);
            }
            indexes.Add(index);
        }

        var fixedOutputIdentities = new HashSet<string>(StringComparer.Ordinal);
        for (int index = 0; index < requests.Count; index++) {
            if (cancellationToken.IsCancellationRequested) return prepared;
            OfficeProvenanceWorkflowRequest request = requests[index];
            if (request.Operation != OfficeProvenanceWorkflowOperation.Remove ||
                request.ConflictPolicy == OfficeWorkflowConflictPolicy.Rename) continue;
            string? outputPath = TryResolveBatchRemovalOutput(request);
            string? identity = pathIndex.TryNormalize(outputPath);
            if (outputPath is null || identity is null) continue;
            EnsureBatchOutputIsSafe(outputPath, identity, index, inputRequestIndexes, fixedOutputIdentities);
        }

        var reservedOutputIdentities = new HashSet<string>(fixedOutputIdentities, StringComparer.Ordinal);
        var nextSuffixByDestination = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int index = 0; index < requests.Count; index++) {
            if (cancellationToken.IsCancellationRequested) return prepared;
            OfficeProvenanceWorkflowRequest request = requests[index];
            if (request.Operation != OfficeProvenanceWorkflowOperation.Remove ||
                request.ConflictPolicy != OfficeWorkflowConflictPolicy.Rename) continue;
            string? requestedPath = TryResolveBatchRemovalOutput(request);
            if (requestedPath is null) continue;
            string? requestedIdentity = pathIndex.TryNormalize(requestedPath);
            if (requestedIdentity is null) continue;
            nextSuffixByDestination.TryGetValue(requestedIdentity, out int nextSuffix);

            string? selectedPath = SelectBatchRenameOutput(
                requestedPath,
                inputRequestIndexes,
                reservedOutputIdentities,
                pathIndex,
                ref nextSuffix,
                cancellationToken);
            if (selectedPath is null) continue;
            nextSuffixByDestination[requestedIdentity] = nextSuffix;
            string? selectedIdentity = pathIndex.TryNormalizeCandidate(selectedPath);
            if (selectedIdentity is null) continue;
            reservedOutputIdentities.Add(selectedIdentity);
            prepared[index] = CloneBatchRequestForReservedOutput(request, selectedPath);
        }
        return prepared;
    }

    private static void EnsureBatchOutputIsSafe(
        string outputPath,
        string identity,
        int requestIndex,
        IReadOnlyDictionary<string, List<int>> inputRequestIndexes,
        ISet<string> outputIdentities) {
        if (inputRequestIndexes.TryGetValue(identity, out List<int>? indexes) &&
            indexes.Any(index => index != requestIndex)) {
            throw new ArgumentException(
                $"Provenance removal output '{outputPath}' overlaps another batch request's input path.",
                "requests");
        }
        if (!outputIdentities.Add(identity)) {
            throw new ArgumentException(
                $"Multiple provenance removal requests resolve to the same output path '{outputPath}'.",
                "requests");
        }
    }

    private static string? SelectBatchRenameOutput(
        string requestedPath,
        IReadOnlyDictionary<string, List<int>> inputRequestIndexes,
        IReadOnlySet<string> reservedOutputIdentities,
        BatchPathIndex pathIndex,
        ref int nextSuffix,
        CancellationToken cancellationToken) {
        IReadOnlySet<string>? existingEntries = pathIndex.TryGetExistingEntries(requestedPath);
        if (existingEntries is null) return null;
        for (int attempts = 0; attempts < 10_000; attempts++) {
            if (cancellationToken.IsCancellationRequested) return null;
            int suffix = nextSuffix;
            if (nextSuffix == int.MaxValue) {
                throw new IOException("No available numbered provenance output path could be reserved for the batch.");
            }
            nextSuffix++;
            string candidate = suffix == 0 ? requestedPath : AddSuffix(requestedPath, suffix);
            string? identity = pathIndex.TryNormalizeCandidate(candidate);
            if (identity is null) return null;
            if (inputRequestIndexes.ContainsKey(identity) || reservedOutputIdentities.Contains(identity) ||
                existingEntries.Contains(identity)) continue;
            return candidate;
        }
        throw new IOException("No available numbered provenance output path could be reserved for the batch.");
    }

    private static OfficeProvenanceWorkflowRequest CloneBatchRequestForReservedOutput(
        OfficeProvenanceWorkflowRequest source,
        string outputPath) => new() {
            Id = source.Id,
            Operation = source.Operation,
            InputPath = source.InputPath,
            OutputPath = outputPath,
            ConflictPolicy = OfficeWorkflowConflictPolicy.Fail,
            Inspection = source.Inspection,
            Assessment = source.Assessment,
            Removal = source.Removal,
            Limits = source.Limits
        };

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
        private readonly Dictionary<string, BatchDirectoryIndex?> _lexicalDirectories = new(StringComparer.Ordinal);
        private readonly Dictionary<string, BatchDirectoryIndex> _physicalDirectories = new(StringComparer.Ordinal);

        internal string? TryNormalize(string? path) {
            if (string.IsNullOrWhiteSpace(path)) return null;
            string fullPath;
            try {
                fullPath = Path.GetFullPath(path);
            } catch (Exception exception) when (exception is ArgumentException or NotSupportedException or IOException or UnauthorizedAccessException) {
                return null;
            }
            if (_normalizedPaths.TryGetValue(fullPath, out string? cached)) return cached;
            try {
                string identity = OfficeWorkflowPathIdentity.Normalize(fullPath);
                _normalizedPaths.Add(fullPath, identity);
                return identity;
            } catch (Exception exception) when (exception is ArgumentException or NotSupportedException or IOException or UnauthorizedAccessException) {
                _normalizedPaths.Add(fullPath, null);
                return null;
            }
        }

        internal string? TryNormalizeCandidate(string path) {
            try {
                string fullPath = Path.GetFullPath(path);
                if (_normalizedPaths.TryGetValue(fullPath, out string? cached)) return cached;
                BatchDirectoryIndex? directory = TryGetDirectory(fullPath);
                if (directory is null) return null;
                string identity = directory.NormalizeChild(Path.GetFileName(fullPath));
                _normalizedPaths.Add(fullPath, identity);
                return identity;
            } catch (Exception exception) when (exception is ArgumentException or NotSupportedException or IOException or UnauthorizedAccessException) {
                return null;
            }
        }

        internal IReadOnlySet<string>? TryGetExistingEntries(string path) {
            try {
                string fullPath = Path.GetFullPath(path);
                return TryGetDirectory(fullPath)?.ExistingEntries;
            } catch (Exception exception) when (exception is ArgumentException or NotSupportedException or IOException or UnauthorizedAccessException) {
                return null;
            }
        }

        private BatchDirectoryIndex? TryGetDirectory(string fullChildPath) {
            string? directoryPath = Path.GetDirectoryName(fullChildPath);
            if (string.IsNullOrEmpty(directoryPath)) return null;
            if (_lexicalDirectories.TryGetValue(directoryPath, out BatchDirectoryIndex? cached)) return cached;
            try {
                if (cancellationToken.IsCancellationRequested) return null;
                string physicalDirectory = OfficeWorkflowPathIdentity.ResolvePhysicalPath(directoryPath);
                bool caseInsensitive = OfficeWorkflowPathIdentity.IsCaseInsensitiveFileSystem(physicalDirectory);
                string physicalIdentity = OfficeWorkflowPathIdentity.Normalize(physicalDirectory, caseInsensitive);
                if (_physicalDirectories.TryGetValue(physicalIdentity, out BatchDirectoryIndex? shared)) {
                    _lexicalDirectories.Add(directoryPath, shared);
                    return shared;
                }
                var existingEntries = new HashSet<string>(StringComparer.Ordinal);
                if (Directory.Exists(directoryPath)) {
                    foreach (string entry in Directory.EnumerateFileSystemEntries(directoryPath)) {
                        if (cancellationToken.IsCancellationRequested) return null;
                        existingEntries.Add(OfficeWorkflowPathIdentity.Normalize(
                            Path.Combine(physicalDirectory, Path.GetFileName(entry)),
                            caseInsensitive));
                    }
                }
                var index = new BatchDirectoryIndex(physicalDirectory, caseInsensitive, existingEntries);
                _physicalDirectories.Add(physicalIdentity, index);
                _lexicalDirectories.Add(directoryPath, index);
                return index;
            } catch (Exception exception) when (exception is ArgumentException or NotSupportedException or IOException or UnauthorizedAccessException) {
                _lexicalDirectories.Add(directoryPath, null);
                return null;
            }
        }
    }

    private sealed class BatchDirectoryIndex(
        string physicalDirectory,
        bool caseInsensitive,
        IReadOnlySet<string> existingEntries) {
        internal IReadOnlySet<string> ExistingEntries { get; } = existingEntries;

        internal string NormalizeChild(string fileName) => OfficeWorkflowPathIdentity.Normalize(
            Path.Combine(physicalDirectory, fileName),
            caseInsensitive);
    }
}
