using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>Reads Outlook for Mac archives through bounded ZIP and XML primitives.</summary>
internal sealed partial class OlmStoreReader {
    private readonly EmailStoreReaderOptions _options;
    private readonly List<EmailStoreDiagnostic> _diagnostics = new List<EmailStoreDiagnostic>();
    private readonly Dictionary<string, EmailStoreFolder> _folders =
        new Dictionary<string, EmailStoreFolder>(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, ZipArchiveEntry> _entries =
        new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);
    private EmailStore _store = null!;
    private CancellationToken _cancellationToken;
    private int _itemCount;
    private long _totalAttachmentBytes;
    private OlmDecodedArchiveBudget _decodedArchiveBudget = null!;

    internal OlmStoreReader(EmailStoreReaderOptions options) {
        _options = options ?? throw new ArgumentNullException(nameof(options));
    }

    internal EmailStoreReadResult Read(Stream stream, string? sourceName, CancellationToken cancellationToken) {
        _cancellationToken = cancellationToken;
        _decodedArchiveBudget = new OlmDecodedArchiveBudget(_options.MaxArchiveDecodedBytes);
        _store = new EmailStore {
            Format = EmailStoreFormat.Olm,
            DisplayName = GetDisplayName(sourceName)
        };

        stream.Position = 0;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true)) {
            IndexArchive(archive);
            foreach (ZipArchiveEntry entry in archive.Entries) {
                _cancellationToken.ThrowIfCancellationRequested();
                if (!IsXmlEntry(entry) || !IsIndexedEntry(entry)) continue;
                ReadXmlEntry(entry);
            }
        }

        return new EmailStoreReadResult(_store, _diagnostics.AsReadOnly(), stream.Length);
    }

    private void IndexArchive(ZipArchive archive) {
        if (archive.Entries.Count > _options.MaxArchiveEntries) {
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxArchiveEntries),
                archive.Entries.Count, _options.MaxArchiveEntries);
        }

        long decodedBytes = 0;
        foreach (ZipArchiveEntry entry in archive.Entries) {
            _cancellationToken.ThrowIfCancellationRequested();
            long length = entry.Length;
            if (length > _options.MaxArchiveEntryBytes) {
                throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxArchiveEntryBytes),
                    length, _options.MaxArchiveEntryBytes);
            }
            decodedBytes = AddBounded(decodedBytes, length,
                nameof(EmailStoreReaderOptions.MaxArchiveDecodedBytes), _options.MaxArchiveDecodedBytes);

            if (!TryNormalizeArchivePath(entry.FullName, out string normalized)) {
                _diagnostics.Add(new EmailStoreDiagnostic(
                    "EMAIL_STORE_OLM_UNSAFE_ENTRY_PATH",
                    "An archive entry with an unsafe path was ignored.",
                    EmailStoreDiagnosticSeverity.Warning,
                    entry.FullName));
                continue;
            }
            if (entry.Name.Length == 0) {
                IndexEmptyMessageFolder(normalized);
                continue;
            }
            if (_entries.ContainsKey(normalized)) {
                _diagnostics.Add(new EmailStoreDiagnostic(
                    "EMAIL_STORE_OLM_DUPLICATE_ENTRY",
                    "A duplicate archive entry path was ignored to keep attachment resolution deterministic.",
                    EmailStoreDiagnosticSeverity.Warning,
                    normalized));
            } else {
                _entries.Add(normalized, entry);
            }
        }
    }

    private void IndexEmptyMessageFolder(string normalizedPath) {
        string[] parts = normalizedPath.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
        int markerIndex = Array.FindIndex(parts, part =>
            string.Equals(part, "com.microsoft.__Messages", StringComparison.OrdinalIgnoreCase));
        if (markerIndex < 0 || parts.Any(part =>
                string.Equals(part, "com.microsoft.__Attachments", StringComparison.OrdinalIgnoreCase))) return;
        string[] visible = parts.Where((_, index) => index != markerIndex).ToArray();
        if (visible.Length > 0) GetOrCreateFolder(string.Join("/", visible));
    }

    private bool IsIndexedEntry(ZipArchiveEntry entry) {
        if (!TryNormalizeArchivePath(entry.FullName, out string normalized) ||
            !_entries.TryGetValue(normalized, out ZipArchiveEntry? indexed)) return false;
        return ReferenceEquals(indexed, entry);
    }

    private void ReadXmlEntry(ZipArchiveEntry entry) {
        string location = entry.FullName;
        try {
            XDocument xml = LoadXml(entry);
            XElement? root = xml.Root;
            if (root == null) return;
            string rootName = root.Name.LocalName;
            if (string.Equals(rootName, "emails", StringComparison.OrdinalIgnoreCase)) {
                ReadItems(entry, root, "email", OutlookItemKind.Message);
            } else if (string.Equals(rootName, "appointments", StringComparison.OrdinalIgnoreCase)) {
                ReadItems(entry, root, "appointment", OutlookItemKind.Appointment);
            } else if (string.Equals(rootName, "contacts", StringComparison.OrdinalIgnoreCase)) {
                ReadItems(entry, root, "contact", OutlookItemKind.Contact);
            } else if (string.Equals(rootName, "tasks", StringComparison.OrdinalIgnoreCase)) {
                ReadItems(entry, root, "task", OutlookItemKind.Task);
            } else if (string.Equals(rootName, "notes", StringComparison.OrdinalIgnoreCase)) {
                ReadItems(entry, root, "note", OutlookItemKind.Note);
            }
        } catch (EmailStoreLimitExceededException) {
            throw;
        } catch (Exception exception) when (exception is XmlException || exception is InvalidDataException ||
                                             exception is IOException) {
            _diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_OLM_XML_INVALID",
                exception.Message,
                EmailStoreDiagnosticSeverity.Error,
                location));
        }
    }

    private void ReadItems(ZipArchiveEntry entry, XElement root, string itemName, OutlookItemKind kind) {
        EmailStoreFolder folder = GetOrCreateFolder(GetFolderPath(entry.FullName));
        int index = 0;
        foreach (XElement item in root.Elements().Where(element =>
                     string.Equals(element.Name.LocalName, itemName, StringComparison.OrdinalIgnoreCase))) {
            _cancellationToken.ThrowIfCancellationRequested();
            _itemCount++;
            if (_itemCount > _options.MaxItemCount) {
                throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxItemCount),
                    _itemCount, _options.MaxItemCount);
            }

            string id = string.Concat("olm:item:", NormalizeSlashes(entry.FullName), "#", index.ToString(CultureInfo.InvariantCulture));
            string location = string.Concat(entry.FullName, "#", index.ToString(CultureInfo.InvariantCulture));
            EmailDocument document = ProjectItem(item, kind, id, folder.Id, location);
            folder.MutableItems.Add(new EmailStoreItem(
                id, folder.Id, document, format: EmailStoreFormat.Olm));
            index++;
        }
    }

    private XDocument LoadXml(ZipArchiveEntry entry) {
        var settings = new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = _options.MaxXmlCharactersPerItem,
            MaxCharactersFromEntities = 0,
            IgnoreComments = true
        };
        using (Stream stream = OpenDecodedEntry(entry, _options.MaxArchiveEntryBytes))
        using (XmlReader reader = XmlReader.Create(stream, settings)) {
            return XDocument.Load(reader, LoadOptions.None);
        }
    }

    private Stream OpenDecodedEntry(ZipArchiveEntry entry, long maximumBytes,
        string limitName = nameof(EmailStoreReaderOptions.MaxArchiveEntryBytes)) =>
        new OlmDecodedEntryStream(entry.Open(), maximumBytes, limitName, _decodedArchiveBudget);

    private EmailStoreFolder GetOrCreateFolder(string path) {
        string normalized = NormalizeSlashes(path).Trim('/');
        if (normalized.Length == 0) normalized = "Archive";
        string[] parts = normalized.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
        string currentPath = string.Empty;
        string? parentId = null;
        EmailStoreFolder? folder = null;
        foreach (string part in parts) {
            currentPath = currentPath.Length == 0 ? part : string.Concat(currentPath, "/", part);
            if (!_folders.TryGetValue(currentPath, out folder)) {
                if (_folders.Count >= _options.MaxFolderCount) {
                    throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxFolderCount),
                        _folders.Count + 1L, _options.MaxFolderCount);
                }
                string id = string.Concat("olm:folder:", currentPath);
                folder = new EmailStoreFolder(id, parentId, part);
                _folders.Add(currentPath, folder);
                _store.MutableFolders.Add(folder);
            }
            parentId = folder.Id;
        }
        return folder!;
    }

    private static string GetFolderPath(string entryPath) {
        string normalized = NormalizeSlashes(entryPath).Trim('/');
        int slash = normalized.LastIndexOf('/');
        if (slash < 0) return "Archive";
        string[] parts = normalized.Substring(0, slash)
            .Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
        string[] visible = parts.Where(part =>
                !string.Equals(part, "com.microsoft.__Messages", StringComparison.OrdinalIgnoreCase))
            .ToArray();
        return visible.Length == 0 ? "Archive" : string.Join("/", visible);
    }

    private static bool IsXmlEntry(ZipArchiveEntry entry) {
        return entry.Name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase);
    }

    private static string? GetDisplayName(string? sourceName) {
        if (string.IsNullOrWhiteSpace(sourceName)) return null;
        try {
            return Path.GetFileNameWithoutExtension(sourceName);
        } catch (Exception exception) when (exception is ArgumentException || exception is NotSupportedException) {
            return sourceName;
        }
    }

    private static bool TryNormalizeArchivePath(string path, out string normalized) {
        normalized = NormalizeSlashes(path).Trim().TrimEnd('/');
        if (normalized.Length == 0 || normalized[0] == '/' || normalized.IndexOf('\0') >= 0) return false;
        if (normalized.Any(char.IsControl)) return false;
        string[] parts = normalized.Split('/');
        for (int index = 0; index < parts.Length; index++) {
            string part = parts[index];
            if (part.Length == 0 || part == "." || part == "..") return false;
            if (index == 0 && part.Length == 2 && char.IsLetter(part[0]) && part[1] == ':') return false;
        }
        return true;
    }

    private static string NormalizeSlashes(string value) {
        return value.Replace('\\', '/');
    }

    private static long AddBounded(long current, long value, string limitName, long limit) {
        if (value < 0 || current > limit - value) {
            long actual = value > long.MaxValue - current ? long.MaxValue : current + value;
            throw new EmailStoreLimitExceededException(limitName, actual, limit);
        }
        return current + value;
    }
}
