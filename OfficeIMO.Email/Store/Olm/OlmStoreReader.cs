using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Email;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Email.Store;

/// <summary>Reads Outlook for Mac archives through bounded ZIP and XML primitives.</summary>
internal sealed partial class OlmStoreReader {
    private readonly EmailStoreReaderOptions _options;
    private readonly EmailStoreDiagnosticCollection _diagnostics = new EmailStoreDiagnosticCollection();
    private readonly Dictionary<string, EmailStoreFolder> _folders =
        new Dictionary<string, EmailStoreFolder>(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, ZipArchiveEntry> _entries =
        new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);
    private EmailStore _store = null!;
    private CancellationToken _cancellationToken;
    private int _itemCount;
    private long _totalAttachmentBytes;
    private OlmDecodedArchiveBudget _decodedArchiveBudget = null!;
    private Action<EmailStoreItem, string, int, OutlookItemKind>? _catalogItem;
    private EmailStoreItemReadOptions? _readOptions;
    private EmailReadWorkspace? _workspace;
    private string? _selectedEntryPath;
    private XDocument? _selectedXml;

    internal OlmStoreReader(EmailStoreReaderOptions options) {
        _options = options ?? throw new ArgumentNullException(nameof(options));
    }

    internal EmailStoreReadResult ReadArchive(ZipArchive archive, string? sourceName, long sourceLength,
        CancellationToken cancellationToken,
        Action<EmailStoreItem, string, int, OutlookItemKind>? catalogItem = null) {
        _cancellationToken = cancellationToken;
        _catalogItem = catalogItem;
        _decodedArchiveBudget = new OlmDecodedArchiveBudget(_options.MaxArchiveDecodedBytes);
        _store = new EmailStore { Format = EmailStoreFormat.Olm, DisplayName = GetDisplayName(sourceName) };
        IndexArchive(archive);
        foreach (ZipArchiveEntry entry in archive.Entries) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (IsXmlEntry(entry) && IsIndexedEntry(entry)) ReadXmlEntry(entry);
        }
        _catalogItem = null;
        return new EmailStoreReadResult(_store, _diagnostics.AsReadOnly(), sourceLength);
    }

    internal EmailStoreItem ReadSelected(string entryPath, int index, OutlookItemKind kind,
        string id, string folderId, EmailStoreItemReadOptions options,
        EmailReadWorkspace? workspace, CancellationToken cancellationToken) {
        _cancellationToken = cancellationToken;
        _readOptions = options;
        _workspace = workspace;
        _totalAttachmentBytes = 0;
        _decodedArchiveBudget = new OlmDecodedArchiveBudget(_options.MaxArchiveDecodedBytes);
        try {
            if (!TryNormalizeArchivePath(entryPath, out string normalized) || !_entries.TryGetValue(normalized, out ZipArchiveEntry? entry))
                throw new InvalidDataException("The indexed OLM entry is no longer available.");
            if (_selectedEntryPath != normalized || _selectedXml == null) {
                // One bounded XML entry is retained so multi-record Contacts/Calendar entries are not
                // decompressed once per selected record. The session requires a stable source.
                _selectedXml = null;
                _selectedEntryPath = null;
                _selectedXml = LoadXml(entry);
                _selectedEntryPath = normalized;
            }
            XElement? root = _selectedXml.Root;
            string elementName = GetItemElementName(kind);
            XElement? item = root?.Elements().Where(element =>
                string.Equals(element.Name.LocalName, elementName, StringComparison.OrdinalIgnoreCase)).Skip(index).FirstOrDefault();
            if (item == null) throw new InvalidDataException("The indexed OLM item is no longer available.");
            return CreateItem(item, kind, id, folderId, entryPath + "#" + index.ToString(CultureInfo.InvariantCulture));
        } finally {
            _readOptions = null;
            _workspace = null;
        }
    }

    internal IReadOnlyList<EmailStoreDiagnostic> Diagnostics => _diagnostics;

    private static string GetItemElementName(OutlookItemKind kind) => kind switch {
        OutlookItemKind.Appointment => "appointment", OutlookItemKind.Contact => "contact",
        OutlookItemKind.Task => "task", OutlookItemKind.Note => "note", _ => "email"
    };

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
            } else {
                _diagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_OLM_XML_UNSUPPORTED",
                    "This XML entry is outside the supported OLM item collections and was not projected.",
                    EmailStoreDiagnosticSeverity.Warning, location));
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
            EmailStoreItem projected = CreateItem(item, kind, id, folder.Id, location);
            if (_catalogItem != null) _catalogItem(projected, entry.FullName, index, kind);
            else folder.MutableItems.Add(projected);
            index++;
        }
    }

    private EmailStoreItem CreateItem(XElement item, OutlookItemKind kind, string id, string folderId, string location) {
        long decodedPropertyBytes = CountItemPropertyBytes(item);
        EmailDocument document = ProjectItem(item, kind, id, folderId, location);
        long maximum = Math.Min(_options.MaxDecodedPropertyBytesPerItem,
            _readOptions?.MaxDecodedPropertyBytes ?? long.MaxValue);
        if (document.Properties.TryGetValue("Olm:StructuredProperties", out object? structured) && structured is string[] fragments) {
            foreach (string fragment in fragments)
                decodedPropertyBytes = AddBounded(decodedPropertyBytes, Encoding.UTF8.GetByteCount(fragment),
                    nameof(EmailStoreReaderOptions.MaxDecodedPropertyBytesPerItem), maximum);
        }
        if (document.Properties.TryGetValue("Olm:ItemAttributes", out object? attributes) && attributes is string attributeXml)
            decodedPropertyBytes = AddBounded(decodedPropertyBytes, Encoding.UTF8.GetByteCount(attributeXml),
                nameof(EmailStoreReaderOptions.MaxDecodedPropertyBytesPerItem), maximum);
        EmailStoreItemReadParts parts = _readOptions?.Parts ?? EmailStoreItemReadParts.All;
        if (_catalogItem != null || (_readOptions == null && !_options.RetainAttachmentContent))
            parts &= ~EmailStoreItemReadParts.AttachmentContent;
        if ((parts & EmailStoreItemReadParts.Bodies) == 0) { document.Body.Text = null; document.Body.Html = null; document.Body.Rtf = null; }
        if ((parts & EmailStoreItemReadParts.Recipients) == 0) document.Recipients.Clear();
        if ((parts & EmailStoreItemReadParts.AttachmentMetadata) == 0) document.Attachments.Clear();
        return new EmailStoreItem(id, folderId, document, loadedParts: parts, format: EmailStoreFormat.Olm) {
            DecodedPropertyBytes = decodedPropertyBytes
        };
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
        using (XmlReader reader = XmlReader.Create(stream, settings))
        using (var bounded = new OfficeXmlLimitingReader(reader, "OLM XML", _options.MaxBTreeDepth,
            (int)Math.Min(int.MaxValue, (long)_options.MaxPropertiesPerItem * _options.MaxItemCount + 1),
            (int)Math.Min(int.MaxValue, (long)_options.MaxPropertiesPerItem * _options.MaxItemCount), _cancellationToken)) {
            return XDocument.Load(bounded, LoadOptions.None);
        }
    }

    private long CountItemPropertyBytes(XElement item) {
        long bytes = 0;
        long properties = item.Descendants().LongCount() + item.DescendantsAndSelf().Attributes().LongCount();
        if (properties > _options.MaxPropertiesPerItem)
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxPropertiesPerItem), properties, _options.MaxPropertiesPerItem);
        long maximum = Math.Min(_options.MaxDecodedPropertyBytesPerItem,
            _readOptions?.MaxDecodedPropertyBytes ?? long.MaxValue);
        foreach (string value in item.DescendantsAndSelf().Select(element => element.Name.ToString())
            .Concat(item.DescendantNodes().OfType<XText>().Select(node => node.Value))
            .Concat(item.DescendantsAndSelf().Attributes().Select(attribute => attribute.Name.ToString() + attribute.Value))) {
            _cancellationToken.ThrowIfCancellationRequested();
            bytes = checked(bytes + Encoding.UTF8.GetByteCount(value));
            if (bytes > maximum)
                throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxDecodedPropertyBytesPerItem), bytes, maximum);
        }
        return bytes;
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
