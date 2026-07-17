using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

internal sealed class EmailStoreExportPathBuilder {
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
        string directory = _preserveHierarchy ? GetFolderPath(reference.FolderId) : _root;
        if (reference.IsAssociated) directory = Path.Combine(directory, "_associated");
        if (reference.IsOrphaned) directory = Path.Combine(directory, "_recovered");
        string baseName = SanitizeSegment(subject, 96, "item");
        string stableId = SanitizeSegment(reference.Id, 48, "id");
        return Path.Combine(directory,
            string.Concat(baseName, "__", stableId, extension));
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
                segments.Push(string.Concat("folder-cycle__", SanitizeSegment(currentId, 48, "id")));
                basePath = _root;
                break;
            }
            if (!_folders.TryGetValue(currentId, out EmailStoreFolderInfo? folder)) {
                segments.Push(string.Concat("unknown-folder__", SanitizeSegment(currentId, 48, "id")));
                basePath = _root;
                break;
            }
            segments.Push(string.Concat(
                SanitizeSegment(folder.Name, 80, "folder"),
                "__",
                SanitizeSegment(folder.Id, 48, "id")));
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
}
