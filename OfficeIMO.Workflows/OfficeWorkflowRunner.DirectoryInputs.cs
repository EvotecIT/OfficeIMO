using OfficeIMO.Core.Internal;
using OfficeIMO.Internal;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private sealed partial class WorkflowInputSnapshots {
        private readonly List<string> _directoryRoots = [];
        private readonly List<(OfficeWorkflowStreamInput Source, string Fingerprint, WorkflowSourceAccess Access)> _directoryFiles = [];
        private readonly List<DirectoryMembership> _directoryManifests = [];

        private async Task<(string Path, long Bytes, int Entries)> CaptureDirectoryAsync(
            OfficeWorkflowDirectoryInput directory, bool recursive, int maximumEntries, long maximumBytes, CancellationToken token) {
            if (maximumEntries < 1) throw new InvalidDataException("The provider folders exceed the workflow entry limit.");
            var options = new OfficeWorkflowDirectoryReadOptions(recursive, maximumEntries);
            IReadOnlyList<OfficeWorkflowDirectoryEntry> entries = await ReadDirectoryEntriesAsync(directory, options, token).ConfigureAwait(false);
            string root = OfficeTemporaryDirectory.Create("officeimo-folder-");
            _directoryRoots.Add(root);
            long bytes = 0;
            var capturedFiles = new Dictionary<string, (string Fingerprint, WorkflowSourceAccess Access)>(StringComparer.Ordinal);
            foreach (var entry in entries) {
                token.ThrowIfCancellationRequested();
                string path = Path.Combine(root, entry.RelativePath.Replace('/', Path.DirectorySeparatorChar));
                if (entry.Input is null) { Directory.CreateDirectory(path); continue; }
                Directory.CreateDirectory(Path.GetDirectoryName(path)!);
                long remaining = maximumBytes - bytes;
                if (remaining < 1) throw new InvalidDataException("The provider inputs exceed the workflow input limit.");
                var access = new WorkflowSourceAccess(entry.Location, entry.Input);
                OfficeWorkflowStreamInput input = access.CreateInput();
                // The original member name determines routing; the intermediate snapshot uses a safe extension.
                using var snapshot = await OfficeStreamFileSnapshot.CaptureAsync(input.OpenRead, ".data", remaining,
                    input.ExpectedSha256, token).ConfigureAwait(false);
                await using (var source = new FileStream(snapshot.FilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                await using (var output = OfficeTemporaryFile.CreateAtPath(path, 81920, FileOptions.Asynchronous)) {
                    await source.CopyToAsync(output, 81920, token).ConfigureAwait(false);
                }
                bytes = checked(bytes + snapshot.Length);
                _directoryFiles.Add((input, snapshot.Fingerprint, access));
                capturedFiles.Add(entry.RelativePath, (snapshot.Fingerprint, access));
            }
            _directoryManifests.Add(new(directory, options, DirectoryManifest(entries), capturedFiles, maximumBytes));
            return (root, bytes, entries.Count);
        }

        private void CleanupDirectories(ref List<Exception>? failures) {
            foreach (string root in _directoryRoots.ToArray()) {
                try { Directory.Delete(root, recursive: true); _directoryRoots.Remove(root); }
                catch (Exception error) when (error is IOException or UnauthorizedAccessException) { (failures ??= []).Add(error); }
            }
        }
    }

    private sealed record DirectoryMembership(OfficeWorkflowDirectoryInput Source,
        OfficeWorkflowDirectoryReadOptions Options, string[] Manifest,
        IReadOnlyDictionary<string, (string Fingerprint, WorkflowSourceAccess Access)> Files, long MaximumBytes);

    private sealed class DirectoryMembershipPublicationGuard(IOfficeWorkflowPublicationGuard? host,
        DirectoryMembership[] directories) : IOfficeWorkflowPublicationGuard {
        public async ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken token) {
            await VerifyAsync(token).ConfigureAwait(false);
            if (host is not null && !await host.CanPublishAsync(path, isDirectory, token).ConfigureAwait(false)) return false;
            await VerifyAsync(token).ConfigureAwait(false);
            return true;
        }

        private async Task VerifyAsync(CancellationToken token) {
            foreach (var directory in directories) {
                var entries = await ReadDirectoryEntriesAsync(directory.Source, directory.Options, token).ConfigureAwait(false);
                if (!directory.Manifest.SequenceEqual(DirectoryManifest(entries), StringComparer.Ordinal))
                    throw new IOException("The provider folder contents changed during execution. Select the folder again.");
                foreach (var entry in entries.Where(entry => entry.Input is not null)) {
                    var expected = directory.Files[entry.RelativePath];
                    // Re-enumeration can return a new item at the same URI while an old item still opens old bytes.
                    await OfficeStreamPublication.VerifyFingerprintAsync(async cancellation => {
                        Stream stream = await entry.Input!.OpenRead(cancellation).ConfigureAwait(false);
                        try { cancellation.ThrowIfCancellationRequested(); expected.Access.VerifyOpenedStream(stream); return stream; }
                        catch { await stream.DisposeAsync().ConfigureAwait(false); throw; }
                    }, expected.Fingerprint, directory.MaximumBytes, token).ConfigureAwait(false);
                }
            }
        }
    }

    private static string[] DirectoryManifest(IReadOnlyList<OfficeWorkflowDirectoryEntry> entries) => entries
        .Select(entry => (entry.Input is null ? "D" : "F") + "\0" + entry.RelativePath + "\0" + entry.Location)
        .OrderBy(value => value, StringComparer.Ordinal).ToArray();

    private static async Task<IReadOnlyList<OfficeWorkflowDirectoryEntry>> ReadDirectoryEntriesAsync(
        OfficeWorkflowDirectoryInput directory, OfficeWorkflowDirectoryReadOptions options, CancellationToken token) {
        var entries = new List<OfficeWorkflowDirectoryEntry>();
        var paths = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        var directories = new HashSet<string>(StringComparer.Ordinal);
        await foreach (var entry in directory.Enumerate(options, token).WithCancellation(token).ConfigureAwait(false)) {
            token.ThrowIfCancellationRequested();
            if (entries.Count >= options.MaximumEntries) throw new InvalidDataException("The provider folder exceeds the workflow entry limit.");
            if (entry is null) throw new InvalidDataException("The provider returned an empty folder entry.");
            string relative = ValidateDirectoryMemberPath(entry.RelativePath, options.MaximumDepth);
            if (!options.IncludeSubdirectories && relative.Contains('/'))
                throw new InvalidDataException("The provider returned nested entries when recursion was disabled.");
            string location = OfficeStorageIdentity.Normalize(entry.Location);
            if (location.Length > 4096 || location.Contains('\0')) throw new InvalidDataException("The provider member identity is invalid.");
            if (!paths.TryAdd(relative, entry.Input is null)) throw new InvalidDataException("The provider folder contains conflicting member names.");
            int slash = relative.LastIndexOf('/');
            if (slash >= 0 && !directories.Contains(relative[..slash]))
                throw new InvalidDataException("The provider must enumerate each parent directory before its members.");
            if (entry.Input is null) directories.Add(relative);
            entries.Add(entry with { RelativePath = relative, Location = location });
        }
        token.ThrowIfCancellationRequested();
        return entries;
    }

    private static string ValidateDirectoryMemberPath(string path, int maximumDepth) {
        if (string.IsNullOrWhiteSpace(path) || path.Length > 1024 || path.Contains('\\'))
            throw new InvalidDataException("The provider member path is invalid.");
        string[] segments = path.Split('/');
        if (segments.Length > maximumDepth) throw new InvalidDataException("The provider folder exceeds the supported depth.");
        foreach (string segment in segments) {
            if (segment.Length is 0 or > 255 || segment is "." or ".." || segment.EndsWith('.') || segment.EndsWith(' ') ||
                segment.Any(character => character < 32 || "<>:\"|?*".Contains(character)))
                throw new InvalidDataException("The provider member path cannot be staged safely.");
            string stem = segment.Split('.')[0];
            if (stem.Equals("CON", StringComparison.OrdinalIgnoreCase) || stem.Equals("PRN", StringComparison.OrdinalIgnoreCase) ||
                stem.Equals("AUX", StringComparison.OrdinalIgnoreCase) || stem.Equals("NUL", StringComparison.OrdinalIgnoreCase) ||
                stem.Length == 4 && (stem.StartsWith("COM", StringComparison.OrdinalIgnoreCase) || stem.StartsWith("LPT", StringComparison.OrdinalIgnoreCase)) &&
                    (char.IsDigit(stem[3]) || "¹²³".Contains(stem[3])))
                throw new InvalidDataException("The provider member path contains a reserved filename.");
        }
        return path;
    }
}
