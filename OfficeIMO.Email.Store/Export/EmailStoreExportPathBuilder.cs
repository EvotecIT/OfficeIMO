using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

internal sealed class EmailStoreExportPathBuilder {
    internal const int MaximumPortableComponentBytes = 255;
    internal const int AtomicTemporarySuffixBytes = 37;

    private readonly string _root;
    private readonly IReadOnlyDictionary<string, EmailStoreFolderInfo> _folders;
    private readonly bool _preserveHierarchy;
    private readonly Dictionary<string, string> _paths = new Dictionary<string, string>(StringComparer.Ordinal);

    internal EmailStoreExportPathBuilder(string root, IEnumerable<EmailStoreFolderInfo> folders,
        bool preserveHierarchy) {
        _root = root;
        _folders = folders.ToDictionary(item => item.Id, StringComparer.Ordinal);
        _preserveHierarchy = preserveHierarchy;
    }

    internal string GetItemPath(EmailStoreItemReference reference, string? subject, EmailFileFormat format) {
        return GetItemPath(reference, subject, GetExtension(format));
    }

    internal string GetItemPath(EmailStoreItemReference reference, string? subject, string extension) {
        string directory = GetItemDirectory(reference);
        string fixedSuffix = string.Concat(".", GetStableHash(reference.Id), extension);
        int availableBytes = MaximumPortableComponentBytes - AtomicTemporarySuffixBytes -
            Encoding.UTF8.GetByteCount("__") - Encoding.UTF8.GetByteCount(fixedSuffix);
        int stableIdBytes = Math.Min(96, availableBytes - Encoding.UTF8.GetByteCount("item"));
        string stableId = SanitizeSegmentByUtf8Bytes(reference.Id, 48, stableIdBytes, "id");
        int baseNameBytes = availableBytes - Encoding.UTF8.GetByteCount(stableId);
        string baseName = SanitizeSegmentByUtf8Bytes(subject, 96, baseNameBytes, "item");
        return Path.Combine(directory,
            string.Concat(baseName, "__", stableId, fixedSuffix));
    }

    internal string GetItemDirectory(EmailStoreItemReference reference) {
        string directory = _preserveHierarchy ? GetFolderPath(reference.FolderId) : _root;
        if (reference.IsAssociated) directory = Path.Combine(directory, "_associated");
        if (reference.IsOrphaned) directory = Path.Combine(directory, "_recovered");
        return directory;
    }

    internal string GetFolderPath(string folderId) {
        if (_paths.TryGetValue(folderId, out string? existing)) return existing;
        var segments = new Stack<string>();
        var visited = new HashSet<string>(StringComparer.Ordinal);
        string? currentId = folderId;
        string basePath = _root;
        while (currentId != null) {
            if (_paths.TryGetValue(currentId, out string? cachedPath)) {
                basePath = cachedPath;
                break;
            }
            if (!visited.Add(currentId)) {
                segments.Push(string.Concat("folder-cycle__", GetStableHash(currentId)));
                basePath = _root;
                break;
            }
            if (!_folders.TryGetValue(currentId, out EmailStoreFolderInfo? folder)) {
                segments.Push(string.Concat("unknown-folder__", GetStableHash(currentId)));
                basePath = _root;
                break;
            }
            string fixedSuffix = string.Concat("__", GetStableHash(folder.Id));
            int nameBytes = MaximumPortableComponentBytes - Encoding.UTF8.GetByteCount(fixedSuffix);
            segments.Push(string.Concat(
                SanitizeSegmentByUtf8Bytes(folder.Name, 80, nameBytes, "folder"),
                fixedSuffix));
            currentId = folder.ParentId;
        }
        while (segments.Count > 0) basePath = Path.Combine(basePath, segments.Pop());
        _paths[folderId] = basePath;
        return basePath;
    }

    private static string GetExtension(EmailFileFormat format) =>
        format == EmailFileFormat.Eml ? ".eml"
        : format == EmailFileFormat.OutlookMsg ? ".msg"
        : format == EmailFileFormat.OutlookTemplate ? ".oft"
        : ".tnef";

    internal static string SanitizeSegment(string? value, int maximumLength, string fallback) {
        if (string.IsNullOrWhiteSpace(value)) return fallback;
        var builder = new StringBuilder(Math.Min(value!.Length, maximumLength));
        bool previousReplacement = false;
        for (int index = 0; index < value.Length && builder.Length < maximumLength; index++) {
            char character = value[index];
            bool allowed = char.IsLetterOrDigit(character) || character == '-' || character == '_' || character == '.';
            if (allowed) {
                builder.Append(character);
                previousReplacement = false;
            } else if (!previousReplacement) {
                builder.Append('_');
                previousReplacement = true;
            }
        }
        string result = builder.ToString().Trim(' ', '.');
        return string.IsNullOrEmpty(result) ? fallback : result;
    }

    internal static string SanitizeSegmentByUtf8Bytes(string? value, int maximumLength,
        int maximumBytes, string fallback) {
        if (maximumBytes < 1) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
        if (string.IsNullOrWhiteSpace(value)) return fallback;
        var builder = new StringBuilder(Math.Min(value!.Length, maximumLength));
        int bytes = 0;
        bool previousReplacement = false;
        for (int index = 0; index < value.Length && builder.Length < maximumLength; index++) {
            char character = value[index];
            bool allowed = char.IsLetterOrDigit(character) || character == '-' || character == '_' || character == '.';
            char output;
            if (allowed) {
                output = character;
                previousReplacement = false;
            } else {
                if (previousReplacement) continue;
                output = '_';
                previousReplacement = true;
            }
            int encodedBytes = output <= 0x7F ? 1 : output <= 0x7FF ? 2 : 3;
            if (bytes > maximumBytes - encodedBytes) break;
            builder.Append(output);
            bytes += encodedBytes;
        }
        string result = builder.ToString().Trim(' ', '.');
        return string.IsNullOrEmpty(result) ? fallback : result;
    }

    internal static string GetStableHash(string value) {
        if (value == null) throw new ArgumentNullException(nameof(value));
        ulong hash = 14695981039346656037UL;
        foreach (byte item in Encoding.UTF8.GetBytes(value)) {
            hash ^= item;
            hash *= 1099511628211UL;
        }
        return hash.ToString("x16", CultureInfo.InvariantCulture);
    }
}
