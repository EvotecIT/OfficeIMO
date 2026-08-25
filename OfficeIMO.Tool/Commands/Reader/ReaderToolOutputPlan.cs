using System.Text;

namespace OfficeIMO.Tool.Commands.Reader;

/// <summary>
/// Validates a complete set of prospective file and directory outputs in linear path-depth time.
/// </summary>
internal sealed class ReaderToolOutputPlan {
    private readonly string _conflictMessage;
    private readonly List<string> _filePaths = new();
    private readonly List<string> _directoryPaths = new();

    internal ReaderToolOutputPlan(string conflictMessage) {
        _conflictMessage = conflictMessage;
    }

    internal void AddFile(string path) => _filePaths.Add(path);

    internal void AddDirectory(string path) => _directoryPaths.Add(path);

    internal void Validate(CancellationToken cancellationToken) {
        try {
            var root = new PathNode();
            var physicalFileIdentities = new HashSet<string>(StringComparer.Ordinal);
            var pathKeyNormalizer = new PathKeyNormalizer();
            foreach (string filePath in _filePaths) {
                cancellationToken.ThrowIfCancellationRequested();
                string fullPath = Path.GetFullPath(filePath);
                if (File.Exists(fullPath) &&
                    !physicalFileIdentities.Add(OfficeImoToolPathSafety.GetPhysicalIdentityKey(fullPath))) {
                    ThrowConflict(fullPath);
                }
                AddFile(root, pathKeyNormalizer.Normalize(fullPath), fullPath);
            }

            foreach (string directoryPath in _directoryPaths) {
                cancellationToken.ThrowIfCancellationRequested();
                string fullPath = Path.GetFullPath(directoryPath);
                EnsureDirectoryDoesNotDescendFromFile(
                    root,
                    pathKeyNormalizer.Normalize(fullPath),
                    fullPath);
            }
        } catch (ReaderToolOutputException) {
            throw;
        } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
            throw new ReaderToolOutputException("Could not validate planned output paths.", exception);
        }
    }

    private void AddFile(PathNode root, string normalizedPath, string originalPath) {
        var visited = new List<PathNode>();
        PathNode node = root;
        visited.Add(node);
        foreach (string segment in GetSegments(normalizedPath)) {
            if (node.FilePath != null) ThrowConflict(originalPath);
            if (!node.Children.TryGetValue(segment, out PathNode? child)) {
                child = new PathNode();
                node.Children.Add(segment, child);
            }
            node = child;
            visited.Add(node);
        }

        if (node.FirstFilePath != null) ThrowConflict(originalPath);
        node.FilePath = originalPath;
        foreach (PathNode visitedNode in visited) {
            visitedNode.FirstFilePath ??= originalPath;
        }
    }

    private void EnsureDirectoryDoesNotDescendFromFile(
        PathNode root,
        string normalizedPath,
        string originalPath) {
        PathNode node = root;
        foreach (string segment in GetSegments(normalizedPath)) {
            if (node.FilePath != null) ThrowDirectoryConflict(originalPath);
            if (!node.Children.TryGetValue(segment, out PathNode? child)) return;
            node = child;
        }
        if (node.FilePath != null) ThrowDirectoryConflict(originalPath);
    }

    private static IEnumerable<string> GetSegments(string path) =>
        path.Split(
            new[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar },
            StringSplitOptions.RemoveEmptyEntries);

    private void ThrowConflict(string path) =>
        throw new ReaderToolOutputException(_conflictMessage + Path.GetFullPath(path));

    private static void ThrowDirectoryConflict(string path) =>
        throw new ReaderToolOutputException(
            "A planned output directory conflicts with a planned output file: " + Path.GetFullPath(path));

    private sealed class PathNode {
        internal Dictionary<string, PathNode> Children { get; } = new(StringComparer.Ordinal);
        internal string? FilePath { get; set; }
        internal string? FirstFilePath { get; set; }
    }

    private sealed class PathKeyNormalizer {
        private readonly Dictionary<string, bool> _normalizationBehavior = new(StringComparer.Ordinal);

        internal string Normalize(string path) {
            string key = OfficeImoToolPathSafety.NormalizePath(path);
            string directory = FindExistingDirectory(path);
            string directoryKey = OfficeImoToolPathSafety.NormalizePath(directory);
            if (!_normalizationBehavior.TryGetValue(directoryKey, out bool normalizeUnicode)) {
                normalizeUnicode = DetectUnicodeNormalizationAliases(directory);
                _normalizationBehavior.Add(directoryKey, normalizeUnicode);
            }
            return normalizeUnicode ? key.Normalize(NormalizationForm.FormC) : key;
        }

        private static string FindExistingDirectory(string path) {
            string? candidate = Directory.Exists(path) ? path : Path.GetDirectoryName(path);
            while (!string.IsNullOrEmpty(candidate) && !Directory.Exists(candidate)) {
                string? parent = Path.GetDirectoryName(candidate);
                if (string.IsNullOrEmpty(parent) || string.Equals(parent, candidate, StringComparison.Ordinal)) break;
                candidate = parent;
            }
            if (string.IsNullOrEmpty(candidate) || !Directory.Exists(candidate)) {
                throw new DirectoryNotFoundException("No existing parent directory was found for '" + path + "'.");
            }
            return candidate;
        }

        private static bool DetectUnicodeNormalizationAliases(string directory) {
            string probeName = ".officeimo-normalization-plan-" + Guid.NewGuid().ToString("N") + "-\u00e9";
            string alternateName = probeName.Normalize(NormalizationForm.FormD);
            string probePath = Path.Combine(directory, probeName);
            string alternatePath = Path.Combine(directory, alternateName);
            using (new FileStream(
                       probePath,
                       FileMode.CreateNew,
                       FileAccess.Write,
                       FileShare.Delete,
                       bufferSize: 1,
                       FileOptions.DeleteOnClose)) {
                return File.Exists(alternatePath);
            }
        }
    }
}
