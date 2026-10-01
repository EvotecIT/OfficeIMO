using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization.Metadata;
using OfficeIMO.Internal;

namespace OfficeIMO.Workflows;

public static partial class OfficePdfArchiveWorkflow {
    private const string Schema = "officeimo.pdf.archive.v1";
    private const int MaximumStateBytes = 64 * 1024;

    private static string Hash(string text) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(text)));

    private static async Task<string> HashFileAsync(string path, string root, long limit, CancellationToken token) {
        using FileStream source = OfficePathIdentity.OpenRegularFileForRead(path, root, 81920);
        if (source.Length > limit) throw new InvalidDataException("Archive file exceeds its configured byte limit.");
        using var hasher = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        var buffer = new byte[81920];
        long total = 0;
        int read;
        while ((read = await source.ReadAsync(buffer, token).ConfigureAwait(false)) != 0) {
            total = checked(total + read);
            if (total > limit) throw new InvalidDataException("Archive file exceeds its configured byte limit.");
            hasher.AppendData(buffer, 0, read);
        }
        return Convert.ToHexString(hasher.GetHashAndReset());
    }

    private static T? ReadState<T>(string path, JsonTypeInfo<T> type) {
        if (!File.Exists(path)) return default;
        EnsureNoLinks(path);
        using var stream = File.OpenRead(path);
        if (stream.Length > MaximumStateBytes) throw new InvalidDataException("Archive checkpoint is oversized.");
        byte[] bytes = new byte[checked((int)stream.Length)];
        stream.ReadExactly(bytes);
        return JsonSerializer.Deserialize(bytes, type) ?? throw new InvalidDataException("Archive checkpoint is empty.");
    }

    private static void WriteState<T>(string path, T value, JsonTypeInfo<T> type) {
        EnsureNoLinks(path);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        byte[] bytes = JsonSerializer.SerializeToUtf8Bytes(value, type);
        if (bytes.Length > MaximumStateBytes) throw new InvalidDataException("Archive checkpoint is oversized.");
        string temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try {
            using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None)) {
                stream.Write(bytes);
                stream.Flush(flushToDisk: true);
            }
            EnsureNoLinks(path);
            File.Move(temporary, path, overwrite: true);
        } finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }

    private static void EnsureNoLinks(string path) {
        for (string? entry = Path.GetFullPath(path); entry != null; entry = Path.GetDirectoryName(entry))
            if ((File.Exists(entry) || Directory.Exists(entry)) && (File.GetAttributes(entry) & FileAttributes.ReparsePoint) != 0)
                throw new InvalidDataException("Archive paths must not traverse symbolic links or junctions.");
    }

    private static OfficePdfArchiveRequest Snapshot(OfficePdfArchiveRequest request) {
        ArgumentNullException.ThrowIfNull(request);
        if (request.MaximumConcurrency is < 1 or > 8 || request.MaximumFiles < 1 || request.MaximumInputBytes < 1 ||
            request.MaximumOutputBytes < 1 || request.TabSize is < 1 or > 32 || request.MaximumTextCharacters < 1 || request.MaximumTextPages < 1)
            throw new ArgumentException("Archive resource limits are invalid.", nameof(request));
        var copy = request with {
            InputDirectory = Path.GetFullPath(request.InputDirectory), OutputDirectory = Path.GetFullPath(request.OutputDirectory),
            CheckpointDirectory = Path.GetFullPath(request.CheckpointDirectory)
        };
        string[] roots = [copy.InputDirectory, copy.OutputDirectory, copy.CheckpointDirectory];
        foreach (string root in roots) EnsureNoLinks(root);
        for (int first = 0; first < roots.Length; first++)
            for (int second = 0; second < roots.Length; second++)
                if (first != second && OfficePathIdentity.IsSameOrDescendant(roots[first], roots[second]))
                    throw new ArgumentException("Archive source, output and checkpoint directories must be separate non-overlapping trees.", nameof(request));
        if (!Directory.Exists(copy.InputDirectory)) throw new DirectoryNotFoundException(copy.InputDirectory);
        return copy;
    }

    private static IEnumerable<string> Discover(string directory, CancellationToken token, int depth = 0) {
        if (depth > 64) throw new InvalidDataException("Archive directory depth exceeds 64.");
        foreach (string path in Directory.EnumerateFileSystemEntries(directory)) {
            token.ThrowIfCancellationRequested();
            FileAttributes attributes = File.GetAttributes(path);
            if ((attributes & FileAttributes.ReparsePoint) != 0) continue;
            if ((attributes & FileAttributes.Directory) != 0) {
                foreach (string child in Discover(path, token, depth + 1)) yield return child;
            } else if (Path.GetExtension(path).ToLowerInvariant() is ".doc" or ".docx" or ".txt") yield return path;
        }
    }
}
