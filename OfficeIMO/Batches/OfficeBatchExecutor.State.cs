using OfficeIMO.Core.Internal;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization.Metadata;
using OfficeIMO.Internal;

namespace OfficeIMO;

internal static partial class OfficeBatchExecutor {
    private const string Schema = "officeimo.conversion.batch.v2";
    private const int MaximumStateBytes = 64 * 1024;

    private static string Hash(string text) => OfficeRuntimeCompatibility.ToHexString(OfficeRuntimeCompatibility.HashSha256(Encoding.UTF8.GetBytes(text)));

    private static async Task<string> HashFileAsync(string path, string root, long limit, CancellationToken token) {
        using FileStream source = OfficePathIdentity.OpenRegularFileForRead(path, root, 81920);
        if (source.Length > limit) throw new InvalidDataException("Batch file exceeds its configured byte limit.");
        using var hasher = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        var buffer = new byte[81920];
        long total = 0;
        int read;
        while ((read = await source.ReadAsync(buffer, 0, buffer.Length, token).ConfigureAwait(false)) != 0) {
            total = checked(total + read);
            if (total > limit) throw new InvalidDataException("Batch file exceeds its configured byte limit.");
            hasher.AppendData(buffer, 0, read);
        }
        return OfficeRuntimeCompatibility.ToHexString(hasher.GetHashAndReset());
    }

    private static T? ReadState<T>(string path, JsonTypeInfo<T> type) {
        if (!File.Exists(path)) return default;
        EnsureNoLinks(path);
        using var stream = File.OpenRead(path);
        if (stream.Length > MaximumStateBytes) throw new InvalidDataException("Batch checkpoint is oversized.");
        byte[] bytes = new byte[checked((int)stream.Length)];
        OfficeRuntimeCompatibility.ReadAtLeast(stream, bytes, bytes.Length);
        return JsonSerializer.Deserialize(bytes, type) ?? throw new InvalidDataException("Batch checkpoint is empty.");
    }

    private static void WriteState<T>(string path, T value, JsonTypeInfo<T> type) {
        EnsureNoLinks(path);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        byte[] bytes = JsonSerializer.SerializeToUtf8Bytes(value, type);
        if (bytes.Length > MaximumStateBytes) throw new InvalidDataException("Batch checkpoint is oversized.");
        string temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try {
            using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None)) {
                OfficeRuntimeCompatibility.WriteAll(stream, bytes);
                stream.Flush(flushToDisk: true);
            }
            EnsureNoLinks(path);
            OfficeFileCommit.CommitTemporaryFileAtomically(temporary, path);
        } finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }

    internal static void EnsureNoLinks(string path) {
        for (string? entry = Path.GetFullPath(path); entry != null; entry = Path.GetDirectoryName(entry))
            if ((File.Exists(entry) || Directory.Exists(entry)) && (File.GetAttributes(entry) & FileAttributes.ReparsePoint) != 0)
                throw new InvalidDataException("Batch paths must not traverse symbolic links or junctions.");
    }

    private static OfficeBatchOptions Snapshot(OfficeBatchOptions request, IReadOnlyList<OfficeConversionExecutionRoute> routes) {
        if (request == null) throw new ArgumentNullException(nameof(request));
        if (request.MaxParallelism is < 1 or > 32 || request.MaximumFiles < 1 || request.MaximumInputBytes < 1 || request.MaximumOutputBytes < 1 || request.MaximumXmlCharactersInPart < 1)
            throw new ArgumentException("Batch resource limits are invalid.", nameof(request));
        var copy = request with {
            InputDirectory = request.InputDirectory == null ? null : Path.GetFullPath(request.InputDirectory),
            OutputDirectory = Path.GetFullPath(request.OutputDirectory),
            ResumeFrom = request.ResumeFrom == null ? null : Path.GetFullPath(request.ResumeFrom),
            InputPaths = request.InputPaths?.Select(Path.GetFullPath).GroupBy(OfficePathIdentity.GetPathIdentityKey).Select(group => group.First()).ToArray(),
            Extensions = request.Extensions?.Select(NormalizeExtension).Distinct(StringComparer.OrdinalIgnoreCase).ToArray(),
            TargetExtension = NormalizeExtension(request.TargetExtension),
            ConversionOptions = (request.ConversionOptions ?? throw new ArgumentException("Conversion settings are required.")).Snapshot()
        };
        if (copy.InputDirectory == null && copy.InputPaths is not { Length: > 0 })
            throw new ArgumentException("Select an input directory or explicit input files.", nameof(request));
        if (copy.InputPaths?.Length > copy.MaximumFiles) throw new ArgumentException("Explicit selection exceeds its file limit.", nameof(request));
        if (!OfficeRuntimeCompatibility.IsDefined(copy.OutputProfile) || !OfficeRuntimeCompatibility.IsDefined(copy.ConflictPolicy)) throw new ArgumentException("Choose supported output settings.");
        if (copy.CheckpointDirectory != null && copy.ConflictPolicy != OfficePublicationPolicy.Fail)
            throw new ArgumentException("Durable jobs require the Fail conflict policy. Ordinary batches also support Rename and Replace.");
        if (!routes.Any(route => route.TargetExtension == copy.TargetExtension))
            throw new NotSupportedException("No executable conversion route produces the requested target format.");
        if (copy.ConversionRouteId != null && routes.FirstOrDefault(route =>
                string.Equals(route.Id, copy.ConversionRouteId, StringComparison.OrdinalIgnoreCase))?.TargetExtension != copy.TargetExtension)
            throw new ArgumentException("The explicit route must be executable and produce the requested target format.");
        if (copy.CheckpointDirectory != null && copy.ConversionRouteId != null && OfficeConverter.FindExecutionRoute(copy.ConversionRouteId)?.Checkpointable != true)
            throw new NotSupportedException("Registered conversion routes require an ordinary batch without checkpoints; their runtime configuration cannot be fingerprinted.");
        string[] roots = new[] { copy.InputDirectory, copy.OutputDirectory }.OfType<string>().ToArray();
        if (copy.ResumeFrom != null) {
            EnsureNoLinks(copy.ResumeFrom);
            EnsureNoLinks(copy.CheckpointDirectory!);
            if (!copy.ResumeFrom.EndsWith(".json", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Use a JSON checkpoint manifest path.", nameof(request));
            if (copy.InputDirectory != null && OfficePathIdentity.IsSameOrDescendant(copy.ResumeFrom, copy.InputDirectory))
                throw new ArgumentException("Checkpoint state must remain outside the source tree.", nameof(request));
        }
        foreach (string root in roots) EnsureNoLinks(root);
        for (int first = 0; first < roots.Length; first++)
            for (int second = 0; second < roots.Length; second++)
                if (first != second && OfficePathIdentity.IsSameOrDescendant(roots[first], roots[second]))
                    throw new ArgumentException("Batch source and output directories must be separate non-overlapping trees.", nameof(request));
        if (copy.InputDirectory != null && !Directory.Exists(copy.InputDirectory)) throw new DirectoryNotFoundException(copy.InputDirectory);
        foreach (string input in copy.InputPaths ?? []) {
            EnsureNoLinks(input);
            if (copy.InputDirectory != null && !OfficePathIdentity.IsSameOrDescendant(input, copy.InputDirectory))
                throw new ArgumentException("Explicit files must be inside the selected source root.");
            if (OfficePathIdentity.IsSameOrDescendant(input, copy.OutputDirectory) ||
                (copy.CheckpointDirectory != null && OfficePathIdentity.IsSameOrDescendant(input, copy.CheckpointDirectory)))
                throw new ArgumentException("Selected input files must be outside output and checkpoint trees.");
            if (copy.CheckpointDirectory != null && SelectRoute(copy, input, routes) is { } route && GetResourceRoot(copy, route.Id, input) is { } resourceRoot) {
                EnsureNoLinks(resourceRoot);
                foreach (string generatedRoot in new[] { copy.OutputDirectory, copy.CheckpointDirectory })
                    if (OfficePathIdentity.IsSameOrDescendant(generatedRoot, resourceRoot) ||
                        OfficePathIdentity.IsSameOrDescendant(resourceRoot, generatedRoot))
                        throw new ArgumentException("Checkpoint resource, output and state directories must be separate non-overlapping trees.", nameof(request));
            }
        }
        return copy;
    }

    private static string? GetResourceRoot(OfficeBatchOptions settings, string routeId, string input) {
        string? root = routeId == "html-pdf" ? Path.GetDirectoryName(input) :
            routeId == "markdown-pdf" && settings.ConversionOptions.Markdown?.ResourcePolicy.AllowLocalFileAccess == true
                ? settings.ConversionOptions.Markdown.BaseDirectory ?? Path.GetDirectoryName(input) : null;
        return root == null ? null : Path.GetFullPath(root);
    }

    private static IEnumerable<string> SelectInputs(OfficeBatchOptions settings, CancellationToken token) =>
        settings.InputPaths ?? Discover(settings.InputDirectory!, settings.Recursive, token);

    private static string NormalizeExtension(string extension) {
        string normalized = extension.Trim().ToLowerInvariant();
        if (!normalized.StartsWith(".", StringComparison.Ordinal)) normalized = "." + normalized;
        if (normalized.Length is < 2 or > 16 || normalized.Skip(1).Any(character => !(character >= 'a' && character <= 'z' || character >= 'A' && character <= 'Z' || character >= '0' && character <= '9')))
            throw new ArgumentException("Use a simple filename extension.");
        return normalized;
    }

    private static OfficeConversionExecutionRoute? SelectRoute(OfficeBatchOptions settings, string input, IReadOnlyList<OfficeConversionExecutionRoute> routes) {
        string extension = Path.GetExtension(input);
        if (settings.Extensions != null && !settings.Extensions.Contains(extension, StringComparer.OrdinalIgnoreCase)) return null;
        if (settings.ConversionRouteId is { } id) {
            var route = routes.First(item => string.Equals(item.Id, id, StringComparison.OrdinalIgnoreCase));
            return route.SourceExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase) ? route : null;
        }
        return OfficeConverter.FindExecutionRoute(extension, settings.TargetExtension, routes);
    }

    private static FileStream? OpenCheckpoint(OfficeBatchOptions settings) {
        if (settings.CheckpointDirectory == null) return null;
        Directory.CreateDirectory(settings.CheckpointDirectory);
        string path = Path.Combine(settings.CheckpointDirectory, "batch.lock");
        EnsureNoLinks(path);
        return new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    }

    private static IEnumerable<string> Discover(string directory, bool recursive, CancellationToken token, int depth = 0, bool rejectLinks = false) {
        if (depth > 64) throw new InvalidDataException("Batch directory depth exceeds 64.");
        foreach (string path in Directory.EnumerateFileSystemEntries(directory)) {
            token.ThrowIfCancellationRequested();
            FileAttributes attributes = File.GetAttributes(path);
            if ((attributes & FileAttributes.ReparsePoint) != 0) {
                if (rejectLinks) throw new InvalidDataException("Checkpoint resource trees must not contain symbolic links or junctions.");
                continue;
            }
            if ((attributes & FileAttributes.Directory) != 0) {
                if (recursive) foreach (string child in Discover(path, true, token, depth + 1, rejectLinks)) yield return child;
            } else yield return path;
        }
    }
}
