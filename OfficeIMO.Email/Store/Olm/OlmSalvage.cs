using System.Buffers.Binary;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;

namespace OfficeIMO.Email.Store;

/// <summary>
/// Recovers an OLM archive whose ZIP structure is damaged. Entries come from the central directory when it is
/// intact and from local entry headers otherwise; each entry is verified against its CRC-32 before it is copied to
/// a rebuilt archive inside the destination root, which the normal OLM reader then projects.
/// </summary>
internal sealed class OlmSalvage {
    private const uint LocalHeaderSignature = 0x04034b50;
    private const uint CentralHeaderSignature = 0x02014b50;
    private const uint EndOfCentralDirectorySignature = 0x06054b50;
    private const uint DataDescriptorSignature = 0x08074b50;
    private const int LocalHeaderLength = 30;
    private const int SearchBufferLength = 1 << 16;
    private const string ItemPrefix = "olm:item:";

    private readonly Stream _source;
    private readonly string _archiveName;
    private readonly EmailStoreSalvageOptions _options;
    private readonly EmailStoreReaderOptions _limits;
    private readonly EmailStoreSalvageWriter _output;
    private readonly CancellationToken _cancellationToken;
    private readonly long _length;
    private readonly List<EmailStoreDiagnostic> _diagnostics = new List<EmailStoreDiagnostic>();
    private readonly Dictionary<string, long> _entryOffsets = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, long> _damagedEntries = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, Repair> _repairedEntries = new Dictionary<string, Repair>(StringComparer.OrdinalIgnoreCase);
    private readonly List<EmailStoreSalvageEntry> _results = new List<EmailStoreSalvageEntry>();
    private long _decodedBytes;
    private int _itemsAdmitted;
    private readonly HashSet<string> _admittedXml = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    private bool _stoppedAtLimit;
    private bool _discoveryCompleted = true;

    internal OlmSalvage(Stream source, string? sourceName, EmailStoreSalvageOptions options,
        EmailStoreSalvageWriter output, CancellationToken cancellationToken) {
        _source = source;
        _options = options;
        _limits = options.ReaderOptions;
        _output = output;
        _cancellationToken = cancellationToken;
        _length = source.Length;
        string baseName = string.IsNullOrWhiteSpace(sourceName)
            ? "Archive" : Path.GetFileNameWithoutExtension(sourceName) ?? "Archive";
        _archiveName = (baseName.Length == 0 ? "Archive" : baseName) + ".olm";
    }

    internal EmailStoreSalvageReport Run() {
        List<ZipEntry> entries = ReadCentralDirectory() ?? ScanLocalHeaders();
        using (FileStream rebuilt = _output.CreateSpill(".olm", deleteOnClose: true)) {
            using (var archive = new ZipArchive(rebuilt, ZipArchiveMode.Create, leaveOpen: true)) {
                foreach (ZipEntry entry in entries) {
                    _cancellationToken.ThrowIfCancellationRequested();
                    CopyEntry(entry, archive);
                }
            }
            rebuilt.Position = 0;
            ExportItems(rebuilt);
        }
        // Unreadable entries are found before projection; report every outcome in archive order.
        EmailStoreSalvageEntry[] ordered = _results.OrderBy(entry => entry.SourceOffset ?? long.MaxValue).ToArray();
        string? manifest = _output.WriteManifest(ordered, _diagnostics);
        return new EmailStoreSalvageReport(EmailStoreFormat.Olm, _output.Root, manifest, _stoppedAtLimit,
            Array.AsReadOnly(ordered), _diagnostics.ToArray(), _discoveryCompleted);
    }

    private List<ZipEntry>? ReadCentralDirectory() {
        var records = new List<OfficeArchiveSafety.ZipCentralDirectoryRecord>();
        OfficeArchiveSafety.ZipCentralDirectoryScanResult scan;
        try {
            scan = OfficeArchiveSafety.ScanZipCentralDirectoryForExtraction(_source, _length,
                _limits.MaxArchiveEntries, records, _cancellationToken);
        } catch (Exception exception) when (exception is InvalidDataException || exception is IOException) {
            scan = OfficeArchiveSafety.ZipCentralDirectoryScanResult.Invalid(exception.Message);
        }
        if (!scan.IsValid || scan.LimitExceeded) {
            Add(scan.LimitExceeded ? "EMAIL_STORE_OLM_SALVAGE_DIRECTORY_LIMIT" : "EMAIL_STORE_OLM_SALVAGE_DIRECTORY_UNREADABLE",
                (scan.LimitExceeded
                    ? "The central directory lists more entries than " + nameof(EmailStoreReaderOptions.MaxArchiveEntries) + "."
                    : "The ZIP central directory is missing or damaged (" + (scan.Error ?? "unreadable") + ").") +
                " Entries are recovered from their local headers.",
                EmailStoreDiagnosticSeverity.Warning, null);
            return null;
        }
        var entries = new List<ZipEntry>(records.Count);
        foreach (OfficeArchiveSafety.ZipCentralDirectoryRecord record in records) {
            entries.Add(new ZipEntry(DecodeName(record.NameBytes), record.Flags, record.CompressionMethod,
                record.Crc32, record.CompressedSize, record.UncompressedSize, record.LocalHeaderOffset, dataOffset: null));
        }
        return entries;
    }

    private List<ZipEntry> ScanLocalHeaders() {
        var entries = new List<ZipEntry>();
        long position = 0;
        var header = new byte[LocalHeaderLength];
        while (position + LocalHeaderLength <= _length) {
            _cancellationToken.ThrowIfCancellationRequested();
            ReadAvailable(position, header, LocalHeaderLength);
            uint signature = UInt32(header, 0);
            if (signature == CentralHeaderSignature || signature == EndOfCentralDirectorySignature) break;
            if (signature != LocalHeaderSignature) {
                long resync = FindSignature(position + 1, LocalHeaderSignature);
                Add("EMAIL_STORE_OLM_SALVAGE_RESYNCHRONIZED",
                    resync < 0
                        ? "Bytes after this offset contain no further local entry header."
                        : "Skipped " + (resync - position).ToString(CultureInfo.InvariantCulture) +
                          " unreadable byte(s) before the resync local entry header.",
                    EmailStoreDiagnosticSeverity.Warning, position);
                if (resync < 0) break;
                position = resync;
                continue;
            }
            if (entries.Count >= _limits.MaxArchiveEntries) {
                _stoppedAtLimit = true;
                Add("EMAIL_STORE_OLM_SALVAGE_DIRECTORY_LIMIT", "Scanning stopped at " +
                    nameof(EmailStoreReaderOptions.MaxArchiveEntries) + ".", EmailStoreDiagnosticSeverity.Warning, position);
                break;
            }
            ushort flags = UInt16(header, 6);
            ushort method = UInt16(header, 8);
            uint crc = UInt32(header, 14);
            long compressed = UInt32(header, 18);
            long expanded = UInt32(header, 22);
            int nameLength = UInt16(header, 26);
            int extraLength = UInt16(header, 28);
            long dataOffset = position + LocalHeaderLength + nameLength + extraLength;
            if (dataOffset > _length) {
                Add("EMAIL_STORE_OLM_SALVAGE_ENTRY_TRUNCATED", "The file ends inside a local entry header.",
                    EmailStoreDiagnosticSeverity.Error, position);
                _discoveryCompleted = false;
                break;
            }
            var nameBytes = new byte[nameLength];
            ReadAvailable(position + LocalHeaderLength, nameBytes, nameLength);
            var extra = new byte[extraLength];
            ReadAvailable(position + LocalHeaderLength + nameLength, extra, extraLength);
            string name = DecodeName(nameBytes);
            if (compressed == uint.MaxValue || expanded == uint.MaxValue) ReadZip64Sizes(extra, ref expanded, ref compressed);

            bool zip64 = HasZip64Extra(extra);
            bool sized = (flags & 0x08) == 0 || compressed != 0 || expanded != 0;
            long next;
            if (sized) {
                next = dataOffset + compressed;
            } else if (TryFindDataDescriptor(dataOffset, zip64, out uint descriptorCrc, out long descriptorCompressed,
                           out long descriptorExpanded, out long descriptorEnd)) {
                crc = descriptorCrc;
                compressed = descriptorCompressed;
                expanded = descriptorExpanded;
                next = descriptorEnd;
            } else {
                // Without sizes or a descriptor the entry runs to the end of the file; it cannot be verified.
                entries.Add(new ZipEntry(name, flags, method, crc, _length - dataOffset, -1, position, dataOffset));
                long resume = FindSignature(dataOffset, LocalHeaderSignature);
                if (resume < 0) break;
                position = resume;
                continue;
            }
            entries.Add(new ZipEntry(name, flags, method, crc, compressed, expanded, position, dataOffset));
            if (next <= position) break;
            position = next;
        }
        return entries;
    }

    private void CopyEntry(ZipEntry entry, ZipArchive archive) {
        string location = entry.Name;
        if (!OlmStoreReader.TryNormalizeArchivePath(entry.Name, out string normalized)) {
            Add("EMAIL_STORE_OLM_UNSAFE_ENTRY_PATH", "An archive entry with an unsafe path was ignored.",
                EmailStoreDiagnosticSeverity.Warning, entry.LocalHeaderOffset, location);
            return;
        }
        if (!_entryOffsets.ContainsKey(normalized)) _entryOffsets.Add(normalized, entry.LocalHeaderOffset);
        bool isDirectory = entry.Name.EndsWith("/", StringComparison.Ordinal) || entry.Name.EndsWith("\\", StringComparison.Ordinal);
        if (isDirectory) {
            archive.CreateEntry(entry.Name.Replace('\\', '/'));
            return;
        }
        if ((entry.Flags & 0x41) != 0 || (entry.Method != 0 && entry.Method != 8)) {
            Damaged(normalized, entry, "EMAIL_STORE_OLM_SALVAGE_ENTRY_UNSUPPORTED",
                "The entry is encrypted or uses an unsupported compression method.");
            return;
        }
        long dataOffset;
        try {
            dataOffset = ResolveDataOffset(entry);
        } catch (InvalidDataException exception) {
            Damaged(normalized, entry, "EMAIL_STORE_OLM_SALVAGE_LOCAL_HEADER_INVALID", exception.Message);
            return;
        }
        if (entry.ExpandedSize > _limits.MaxArchiveEntryBytes) {
            Damaged(normalized, entry, "EMAIL_STORE_OLM_SALVAGE_ENTRY_LIMIT",
                "The entry is larger than " + nameof(EmailStoreReaderOptions.MaxArchiveEntryBytes) + ".");
            return;
        }

        bool isXml = normalized.EndsWith(".xml", StringComparison.OrdinalIgnoreCase);
        if (isXml && _admittedXml.Contains(normalized)) {
            // The reader keeps the first copy of a path, so a later copy is skipped before it costs any decoding.
            Add("EMAIL_STORE_OLM_DUPLICATE_ENTRY",
                "A later copy of an XML entry that was already recovered was ignored, as the OLM reader would ignore it.",
                EmailStoreDiagnosticSeverity.Warning, entry.LocalHeaderOffset, normalized);
            return;
        }
        // Every decoding attempt, including failed verification and repair, counts against the archive budget.
        long remaining = _limits.MaxArchiveDecodedBytes - _decodedBytes;
        if (entry.ExpandedSize > remaining) {
            DecodedLimit(entry);
            return;
        }
        string? failure = Verify(entry, dataOffset, out long decoded);
        _decodedBytes += decoded;
        int itemsRemaining = _limits.MaxItemCount - _itemsAdmitted;

        if (failure == null) {
            if (isXml) {
                int count = CountXmlItems(entry, dataOffset);
                if (count > itemsRemaining) {
                    // Keep the records that fit the source item bound; the rest of the entry is not examined.
                    if (itemsRemaining > 0 && TryRepairXml(entry, dataOffset, itemsRemaining, out FileStream? subset,
                            out int kept, out _, out _, out _)) {
                        using (subset) WriteSpill(archive, entry.Name, subset!);
                        _itemsAdmitted += kept;
                        _admittedXml.Add(normalized);
                    }
                    SourceItemLimit(entry);
                    return;
                }
                _itemsAdmitted += count;
                _admittedXml.Add(normalized);
            }
            ZipArchiveEntry target = archive.CreateEntry(entry.Name, CompressionLevel.NoCompression);
            using Stream input = OpenPayload(entry, dataOffset);
            using Stream output = target.Open();
            CopyVerified(input, output, decoded, entry.Crc);
            return;
        }

        if (isXml && itemsRemaining > 0 && _decodedBytes < _limits.MaxArchiveDecodedBytes &&
            TryRepairXml(entry, dataOffset, itemsRemaining, out FileStream? repaired, out int items, out long consumed,
                out string? stopReason, out bool limited)) {
            _decodedBytes += consumed;
            using (repaired) WriteSpill(archive, entry.Name, repaired!);
            _itemsAdmitted += items;
            _admittedXml.Add(normalized);
            _repairedEntries[normalized] = new Repair(entry.LocalHeaderOffset, failure);
            Add("EMAIL_STORE_OLM_SALVAGE_ENTRY_REPAIRED",
                failure + " " + items.ToString(CultureInfo.InvariantCulture) +
                " complete record(s) were kept; they cannot be verified against the entry checksum." +
                (stopReason == null ? string.Empty : " Reading stopped after " +
                    consumed.ToString(CultureInfo.InvariantCulture) + " decoded byte(s): " + stopReason),
                EmailStoreDiagnosticSeverity.Warning, entry.LocalHeaderOffset, location);
            if (limited) SourceItemLimit(entry);
            return;
        }
        Damaged(normalized, entry, isXml ? "EMAIL_STORE_OLM_SALVAGE_XML_ENTRY_UNREADABLE" : "EMAIL_STORE_OLM_SALVAGE_ENTRY_DAMAGED",
            failure);
        if (isXml) {
            TryAddResult(new EmailStoreSalvageEntry("olm:entry:" + normalized, EmailStoreSalvageStatus.NotRecovered,
                normalized, entry.LocalHeaderOffset, 0, entry.ExpandedSize >= 0 ? entry.ExpandedSize : null, null,
                entry.LocalHeaderOffset, "EMAIL_STORE_OLM_SALVAGE_XML_ENTRY_UNREADABLE", null, null, null, 0,
                new[] { new EmailStoreDiagnostic("EMAIL_STORE_OLM_SALVAGE_XML_ENTRY_UNREADABLE", failure,
                    EmailStoreDiagnosticSeverity.Error, normalized) }));
        }
    }

    private static void WriteSpill(ZipArchive archive, string name, FileStream spill) {
        ZipArchiveEntry target = archive.CreateEntry(name, CompressionLevel.NoCompression);
        spill.Position = 0;
        using Stream output = target.Open();
        spill.CopyTo(output);
    }

    /// <summary>The record element the OLM reader projects under each collection root.</summary>
    private static readonly Dictionary<string, string> ItemElements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
        ["emails"] = "email", ["appointments"] = "appointment", ["contacts"] = "contact",
        ["tasks"] = "task", ["notes"] = "note"
    };

    /// <summary>Whether a top-level element is a record the OLM reader projects; other roots project nothing.</summary>
    private static bool IsItemElement(string? rootName, string localName) =>
        rootName != null && ItemElements.TryGetValue(rootName, out string? item) &&
        string.Equals(item, localName, StringComparison.OrdinalIgnoreCase);

    /// <summary>Counts an XML entry's top-level records; malformed XML counts as none and is reported by the reader.</summary>
    private int CountXmlItems(ZipEntry entry, long dataOffset) {
        try {
            using Stream input = OpenPayload(entry, dataOffset);
            using XmlReader reader = XmlReader.Create(input, CreateXmlSettings());
            if (reader.MoveToContent() != XmlNodeType.Element || reader.IsEmptyElement) return 0;
            string rootName = reader.LocalName;
            int count = 0;
            reader.Read();
            while (!reader.EOF && reader.NodeType != XmlNodeType.EndElement) {
                _cancellationToken.ThrowIfCancellationRequested();
                if (reader.NodeType == XmlNodeType.Element) {
                    if (IsItemElement(rootName, reader.LocalName) && ++count == int.MaxValue) return count;
                    reader.Skip();
                } else {
                    reader.Read();
                }
            }
            return count;
        } catch (Exception exception) when (exception is XmlException || exception is InvalidDataException) {
            return 0;
        }
    }

    private XmlReaderSettings CreateXmlSettings() => new XmlReaderSettings {
        DtdProcessing = DtdProcessing.Prohibit,
        XmlResolver = null,
        MaxCharactersInDocument = _limits.MaxArchiveEntryBytes,
        MaxCharactersFromEntities = 0,
        IgnoreComments = true,
        CloseInput = false
    };

    private string? Verify(ZipEntry entry, long dataOffset, out long decoded) {
        decoded = 0;
        if (entry.ExpandedSize < 0) return "The entry has no recorded size or data descriptor, so it cannot be verified.";
        if (dataOffset + entry.CompressedSize > _length) return "The file ends inside the entry data.";
        uint crc = OfficePngCrc32.Begin();
        var buffer = new byte[81920];
        try {
            using Stream input = OpenPayload(entry, dataOffset);
            while (true) {
                _cancellationToken.ThrowIfCancellationRequested();
                // Reading at most one byte past the declared size detects overlong data without extra work.
                int read = input.Read(buffer, 0, (int)Math.Min(buffer.Length, entry.ExpandedSize - decoded + 1));
                if (read == 0) break;
                decoded += read;
                if (decoded > entry.ExpandedSize) return "The entry decodes to more bytes than it declares.";
                crc = OfficePngCrc32.Append(crc, buffer, 0, read);
            }
        } catch (InvalidDataException exception) {
            return "The entry data cannot be decompressed (" + exception.Message + ").";
        }
        if (decoded != entry.ExpandedSize) return "The entry decodes to fewer bytes than it declares.";
        if (OfficePngCrc32.Complete(crc) != entry.Crc) return "The entry data does not match its CRC-32.";
        return null;
    }

    private bool TryRepairXml(ZipEntry entry, long dataOffset, int maxItems, out FileStream? repaired, out int items,
        out long consumed, out string? stopReason, out bool limited) {
        repaired = null;
        items = 0;
        consumed = 0;
        stopReason = null;
        limited = false;
        FileStream spill = _output.CreateSpill(".xml", deleteOnClose: true);
        bool transferred = false;
        try {
            transferred = RepairXml(entry, dataOffset, maxItems, spill, ref items, ref consumed, ref stopReason, ref limited);
            if (transferred) repaired = spill;
            return transferred;
        } finally {
            if (!transferred) spill.Dispose();
        }
    }

    /// <summary>Copies up to <paramref name="maxItems"/> complete top-level records of an XML entry into a spill.</summary>
    private bool RepairXml(ZipEntry entry, long dataOffset, int maxItems, FileStream spill, ref int items,
        ref long consumed, ref string? stopReason, ref bool limited) {
        XmlReaderSettings settings = CreateXmlSettings();
        var counting = new CountingStream(OpenPayload(entry, dataOffset, tolerateTruncation: true),
            Math.Min(_limits.MaxArchiveEntryBytes, _limits.MaxArchiveDecodedBytes - _decodedBytes));
        try {
            using (counting)
            using (XmlReader reader = XmlReader.Create(counting, settings))
            using (XmlWriter writer = XmlWriter.Create(spill, new XmlWriterSettings {
                       Encoding = new UTF8Encoding(false), CloseOutput = false })) {
                bool started = false;
                try {
                    if (reader.MoveToContent() != XmlNodeType.Element) return false;
                    string rootName = reader.LocalName;
                    writer.WriteStartElement(reader.Prefix, reader.LocalName, reader.NamespaceURI);
                    started = true;
                    if (reader.HasAttributes) writer.WriteAttributes(reader, defattr: false);
                    if (!reader.IsEmptyElement) {
                        reader.Read();
                        while (!reader.EOF && reader.NodeType != XmlNodeType.EndElement) {
                            _cancellationToken.ThrowIfCancellationRequested();
                            if (reader.NodeType == XmlNodeType.Element) {
                                bool isItem = IsItemElement(rootName, reader.LocalName);
                                XNode node = XNode.ReadFrom(reader);
                                node.WriteTo(writer);
                                if (isItem && ++items >= maxItems) {
                                    limited = reader.MoveToContent() == XmlNodeType.Element;
                                    break;
                                }
                            } else {
                                reader.Read();
                            }
                        }
                    }
                } catch (Exception exception) when (exception is XmlException || exception is InvalidDataException ||
                                                     exception is IOException) {
                    stopReason = exception.Message;
                }
                if (started) writer.WriteEndElement();
            }
            consumed = counting.BytesRead;
        } catch (Exception exception) when (exception is XmlException || exception is InvalidDataException || exception is IOException) {
            stopReason = exception.Message;
            return false;
        }
        return items > 0;
    }

    private void ExportItems(Stream rebuilt) {
        EmailStoreSession session;
        try {
            session = EmailStoreSession.Open(rebuilt, _archiveName, _limits, leaveOpen: true, _cancellationToken);
        } catch (Exception exception) when (exception is InvalidDataException || exception is IOException ||
                                             exception is EmailStoreLimitExceededException) {
            Add("EMAIL_STORE_OLM_SALVAGE_REBUILD_UNREADABLE", exception.Message, EmailStoreDiagnosticSeverity.Error, null);
            _discoveryCompleted = false;
            return;
        }
        using (session) {
            var paths = new EmailStoreExportPathBuilder(_output.Root, session.Folders, preserveHierarchy: true);
            EmailFileFormat format = _options.Format ?? EmailFileFormat.OutlookMsg;
            foreach (EmailStoreItemReference reference in session.EnumerateItems(cancellationToken: _cancellationToken)) {
                _cancellationToken.ThrowIfCancellationRequested();
                if (AtItemLimit()) break;
                _results.Add(ExportItem(session, reference, paths, format));
            }
            foreach (EmailStoreDiagnostic diagnostic in session.Diagnostics) {
                if (diagnostic.Code == "EMAIL_STORE_OLM_XML_INVALID" && diagnostic.Location != null &&
                    OlmStoreReader.TryNormalizeArchivePath(diagnostic.Location, out string entryPath)) {
                    long? offset = _entryOffsets.TryGetValue(entryPath, out long value) ? value : null;
                    if (!TryAddResult(new EmailStoreSalvageEntry("olm:entry:" + entryPath, EmailStoreSalvageStatus.NotRecovered,
                            entryPath, offset, 0, null, null, offset, diagnostic.Code, null, null, null, 0, new[] { diagnostic })))
                        _diagnostics.Add(diagnostic);
                } else if (diagnostic.Code != "EMAIL_STORE_OLM_ATTACHMENT_MISSING") {
                    _diagnostics.Add(diagnostic);
                }
            }
        }
    }

    /// <summary>Applies the outcome bound; the stop diagnostic is reported once.</summary>
    private bool AtItemLimit() {
        if (_results.Count < _options.MaxRecoveredItems) return false;
        if (!_stoppedAtLimit || !_diagnostics.Any(diagnostic => diagnostic.Code == "EMAIL_STORE_SALVAGE_ITEM_LIMIT"))
            Add("EMAIL_STORE_SALVAGE_ITEM_LIMIT", "Recovery stopped at the configured item bound.",
                EmailStoreDiagnosticSeverity.Warning, null);
        _stoppedAtLimit = true;
        return true;
    }

    private bool TryAddResult(EmailStoreSalvageEntry entry) {
        if (AtItemLimit()) return false;
        _results.Add(entry);
        return true;
    }

    private EmailStoreSalvageEntry ExportItem(EmailStoreSession session, EmailStoreItemReference reference,
        EmailStoreExportPathBuilder paths, EmailFileFormat format) {
        var diagnostics = new List<EmailStoreDiagnostic>();
        string entryPath = GetEntryPath(reference.Id);
        string location = reference.Id.StartsWith(ItemPrefix, StringComparison.Ordinal)
            ? reference.Id.Substring(ItemPrefix.Length) : reference.Id;
        long? sourceOffset = _entryOffsets.TryGetValue(entryPath, out long offset) ? offset : null;
        long? damageOffset = null;
        string? damageCode = null;
        if (_repairedEntries.TryGetValue(entryPath, out Repair repair)) {
            damageOffset = repair.Offset;
            damageCode = "EMAIL_STORE_OLM_SALVAGE_ENTRY_REPAIRED";
            diagnostics.Add(new EmailStoreDiagnostic(damageCode,
                "The record came from a damaged XML entry and cannot be verified: " + repair.Reason,
                EmailStoreDiagnosticSeverity.Warning, location));
        }
        OutlookItemKind? kind = null;
        string? subject = null;
        try {
            EmailStoreItem item = session.ReadItem(reference, _cancellationToken);
            kind = item.Document.OutlookItemKind;
            subject = item.Document.Subject;
            foreach (EmailAttachment attachment in item.Document.Attachments.ToArray()) {
                if (attachment.Content != null || attachment.ContentSource != null ||
                    string.IsNullOrWhiteSpace(attachment.ContentLocation)) continue;
                // The artifact omits an attachment whose bytes were lost instead of writing an empty placeholder.
                item.Document.Attachments.Remove(attachment);
                string attachmentPath = attachment.ContentLocation!.Replace('\\', '/').Trim();
                bool damaged = _damagedEntries.TryGetValue(attachmentPath, out long attachmentOffset);
                damageOffset ??= damaged ? attachmentOffset : null;
                damageCode ??= "EMAIL_STORE_OLM_SALVAGE_ATTACHMENT_MISSING";
                diagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_OLM_SALVAGE_ATTACHMENT_MISSING",
                    damaged
                        ? "The attachment entry is damaged and was not recovered: " + attachmentPath
                        : "The attachment entry is missing from the archive: " + attachmentPath,
                    EmailStoreDiagnosticSeverity.Warning, location));
            }
            string path = _output.ResolvePath(paths, reference, subject, format);
            long written = _output.WriteDocument(item.Document, path, format, reference.Id, diagnostics);
            if (written < 0)
                return new EmailStoreSalvageEntry(reference.Id, EmailStoreSalvageStatus.NotRecovered, location,
                    sourceOffset, 0, null, null, damageOffset, damageCode ?? "EMAIL_STORE_SALVAGE_ITEM_WRITE_FAILED",
                    kind, subject, null, 0, diagnostics.AsReadOnly());
            return new EmailStoreSalvageEntry(reference.Id,
                damageCode == null ? EmailStoreSalvageStatus.Recovered : EmailStoreSalvageStatus.Partial,
                location, sourceOffset, item.DecodedPropertyBytes ?? 0, null, null, damageOffset, damageCode, kind, subject,
                path, written, diagnostics.AsReadOnly());
        } catch (Exception exception) when (exception is InvalidDataException || exception is IOException ||
                                             exception is UnauthorizedAccessException || exception is NotSupportedException ||
                                             exception is EmailStoreLimitExceededException || exception is EmailLimitExceededException) {
            diagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_SALVAGE_ITEM_WRITE_FAILED", exception.Message,
                EmailStoreDiagnosticSeverity.Error, location));
            return new EmailStoreSalvageEntry(reference.Id, EmailStoreSalvageStatus.NotRecovered, location,
                sourceOffset, 0, null, null, damageOffset, damageCode ?? "EMAIL_STORE_SALVAGE_ITEM_WRITE_FAILED",
                kind, subject, null, 0, diagnostics.AsReadOnly());
        }
    }

    private static string GetEntryPath(string itemId) {
        if (!itemId.StartsWith(ItemPrefix, StringComparison.Ordinal)) return itemId;
        string value = itemId.Substring(ItemPrefix.Length);
        int hash = value.LastIndexOf('#');
        return hash < 0 ? value : value.Substring(0, hash);
    }

    private long ResolveDataOffset(ZipEntry entry) {
        if (entry.DataOffset.HasValue) return entry.DataOffset.Value;
        var header = new byte[LocalHeaderLength];
        if (entry.LocalHeaderOffset < 0 || ReadAvailable(entry.LocalHeaderOffset, header, LocalHeaderLength) < LocalHeaderLength ||
            UInt32(header, 0) != LocalHeaderSignature)
            throw new InvalidDataException("The central directory points at a missing or damaged local entry header.");
        return entry.LocalHeaderOffset + LocalHeaderLength + UInt16(header, 26) + UInt16(header, 28);
    }

    private Stream OpenPayload(ZipEntry entry, long dataOffset, bool tolerateTruncation = false) {
        long length = Math.Max(0, Math.Min(entry.CompressedSize, _length - dataOffset));
        if (!tolerateTruncation && length != entry.CompressedSize)
            throw new InvalidDataException("The file ends inside the entry data.");
        Stream raw = new SliceStream(_source, dataOffset, length);
        return entry.Method == 8 ? new DeflateStream(raw, CompressionMode.Decompress) : raw;
    }

    private void DecodedLimit(ZipEntry entry) {
        _stoppedAtLimit = true;
        Add("EMAIL_STORE_OLM_SALVAGE_DECODED_LIMIT", "The archive reached " +
            nameof(EmailStoreReaderOptions.MaxArchiveDecodedBytes) + "; this entry was not recovered.",
            EmailStoreDiagnosticSeverity.Warning, entry.LocalHeaderOffset, entry.Name);
    }

    private void SourceItemLimit(ZipEntry entry) {
        _stoppedAtLimit = true;
        Add("EMAIL_STORE_OLM_SALVAGE_SOURCE_ITEM_LIMIT", "The archive reached " +
            nameof(EmailStoreReaderOptions.MaxItemCount) + "; later records in this entry were not examined.",
            EmailStoreDiagnosticSeverity.Warning, entry.LocalHeaderOffset, entry.Name);
    }

    /// <summary>Copies verified entry bytes and checks them again, so a changed source cannot slip in.</summary>
    private void CopyVerified(Stream input, Stream output, long length, uint expectedCrc) {
        var buffer = new byte[81920];
        uint crc = OfficePngCrc32.Begin();
        long remaining = length;
        while (remaining > 0) {
            _cancellationToken.ThrowIfCancellationRequested();
            int read = input.Read(buffer, 0, (int)Math.Min(buffer.Length, remaining));
            if (read == 0) throw new IOException("The OLM source changed during recovery.");
            crc = OfficePngCrc32.Append(crc, buffer, 0, read);
            output.Write(buffer, 0, read);
            remaining -= read;
        }
        if (OfficePngCrc32.Complete(crc) != expectedCrc) throw new IOException("The OLM source changed during recovery.");
    }

    private void Damaged(string normalized, ZipEntry entry, string code, string message) {
        _damagedEntries[normalized] = entry.LocalHeaderOffset;
        Add(code, message, EmailStoreDiagnosticSeverity.Error, entry.LocalHeaderOffset, normalized);
    }

    /// <summary>
    /// Finds the data descriptor that follows an unsized entry. Descriptors may omit their signature and use
    /// eight-byte sizes for ZIP64 entries, so candidates ending at the next header are tried in every layout
    /// (ZIP64 first when the local header carries a ZIP64 field) before falling back to a signed descriptor
    /// anywhere later. A candidate is accepted only when its compressed size equals its distance from the data;
    /// the CRC-32 is checked when the entry is verified.
    /// </summary>
    private bool TryFindDataDescriptor(long dataOffset, bool zip64, out uint crc, out long compressed,
        out long expanded, out long end) {
        long search = dataOffset;
        while (true) {
            long header = FindHeader(search);
            if (header < 0) break;
            if (TryDescriptorEndingAt(dataOffset, header, zip64, out crc, out compressed, out expanded)) {
                end = header;
                return true;
            }
            search = header + 1;
        }
        if (TryDescriptorEndingAt(dataOffset, _length, zip64, out crc, out compressed, out expanded)) {
            end = _length;
            return true;
        }
        search = dataOffset;
        var descriptor = new byte[24];
        while (true) {
            long candidate = FindSignature(search, DataDescriptorSignature);
            if (candidate < 0) break;
            int read = ReadAvailable(candidate, descriptor, descriptor.Length);
            long distance = candidate - dataOffset;
            foreach (bool wide in zip64 ? new[] { true, false } : new[] { false, true }) {
                int length = wide ? 24 : 16;
                if (read < length || ReadSize(descriptor, 8, wide) != distance) continue;
                crc = UInt32(descriptor, 4);
                compressed = distance;
                expanded = ReadSize(descriptor, wide ? 16 : 12, wide);
                end = candidate + length;
                return true;
            }
            search = candidate + 1;
        }
        crc = 0;
        compressed = expanded = end = 0;
        return false;
    }

    private bool TryDescriptorEndingAt(long dataOffset, long end, bool zip64, out uint crc, out long compressed,
        out long expanded) {
        var bytes = new byte[24];
        foreach (bool wide in zip64 ? new[] { true, false } : new[] { false, true }) {
            foreach (bool signed in new[] { true, false }) {
                int length = (wide ? 20 : 12) + (signed ? 4 : 0);
                long start = end - length;
                if (start < dataOffset || ReadAvailable(start, bytes, length) < length) continue;
                int field = signed ? 4 : 0;
                if (signed && UInt32(bytes, 0) != DataDescriptorSignature) continue;
                if (ReadSize(bytes, field + 4, wide) != start - dataOffset) continue;
                crc = UInt32(bytes, field);
                compressed = start - dataOffset;
                expanded = ReadSize(bytes, field + (wide ? 12 : 8), wide);
                return true;
            }
        }
        crc = 0;
        compressed = expanded = 0;
        return false;
    }

    private static long ReadSize(byte[] bytes, int offset, bool wide) =>
        wide ? (long)BinaryPrimitives.ReadUInt64LittleEndian(bytes.AsSpan(offset, 8)) : UInt32(bytes, offset);

    private long FindHeader(long start) {
        var buffer = new byte[SearchBufferLength];
        long position = start;
        while (position + 4 <= _length) {
            _cancellationToken.ThrowIfCancellationRequested();
            int read = ReadFull(position, buffer);
            for (int index = 0; index + 4 <= read; index++) {
                uint value = UInt32(buffer, index);
                if (value == LocalHeaderSignature || value == CentralHeaderSignature ||
                    value == EndOfCentralDirectorySignature) return position + index;
            }
            if (position + read >= _length) break;
            position += read - 3;
        }
        return -1;
    }

    private static bool HasZip64Extra(byte[] extra) {
        int offset = 0;
        while (offset + 4 <= extra.Length) {
            if (BinaryPrimitives.ReadUInt16LittleEndian(extra.AsSpan(offset, 2)) == 1) return true;
            offset += 4 + BinaryPrimitives.ReadUInt16LittleEndian(extra.AsSpan(offset + 2, 2));
        }
        return false;
    }

    private long FindSignature(long start, uint signature) {
        var buffer = new byte[SearchBufferLength];
        long position = start;
        while (position + 4 <= _length) {
            _cancellationToken.ThrowIfCancellationRequested();
            int read = ReadFull(position, buffer);
            for (int index = 0; index + 4 <= read; index++)
                if (UInt32(buffer, index) == signature) return position + index;
            if (position + read >= _length) break;
            position += read - 3;
        }
        return -1;
    }

    private static void ReadZip64Sizes(byte[] extra, ref long expanded, ref long compressed) {
        int offset = 0;
        while (offset + 4 <= extra.Length) {
            int id = BinaryPrimitives.ReadUInt16LittleEndian(extra.AsSpan(offset, 2));
            int size = BinaryPrimitives.ReadUInt16LittleEndian(extra.AsSpan(offset + 2, 2));
            if (offset + 4 + size > extra.Length) return;
            if (id == 1) {
                int field = offset + 4;
                if (expanded == uint.MaxValue && field + 8 <= offset + 4 + size) {
                    expanded = (long)BinaryPrimitives.ReadUInt64LittleEndian(extra.AsSpan(field, 8));
                    field += 8;
                }
                if (compressed == uint.MaxValue && field + 8 <= offset + 4 + size)
                    compressed = (long)BinaryPrimitives.ReadUInt64LittleEndian(extra.AsSpan(field, 8));
                return;
            }
            offset += 4 + size;
        }
    }

    private static string DecodeName(byte[] bytes) => Encoding.UTF8.GetString(bytes);

    /// <summary>Reads a search window; a short read means the source shrank, which would stall the search.</summary>
    private int ReadFull(long position, byte[] buffer) {
        int expected = (int)Math.Min(buffer.Length, _length - position);
        int read = ReadAvailable(position, buffer, expected);
        if (read != expected) throw new IOException("The OLM source changed length during recovery.");
        return read;
    }

    private int ReadAvailable(long offset, byte[] buffer, int count) {
        if (offset >= _length || offset < 0) return 0;
        count = (int)Math.Min(count, _length - offset);
        _source.Position = offset;
        int total = 0;
        while (total < count) {
            int read = _source.Read(buffer, total, count - total);
            if (read == 0) break;
            total += read;
        }
        return total;
    }

    private void Add(string code, string message, EmailStoreDiagnosticSeverity severity, long? offset,
        string? location = null) =>
        _diagnostics.Add(new EmailStoreDiagnostic(code, message, severity,
            location ?? (offset.HasValue ? "offset 0x" + offset.Value.ToString("x", CultureInfo.InvariantCulture) : null),
            "salvage", offset, null, null, null, EmailDiagnosticDisposition.Recovered,
            severity == EmailStoreDiagnosticSeverity.Error ? EmailDataLossRisk.Confirmed : EmailDataLossRisk.Possible, null));

    private static uint UInt32(byte[] bytes, int offset) =>
        BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(offset, 4));

    private static ushort UInt16(byte[] bytes, int offset) =>
        BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(offset, 2));

    private sealed class ZipEntry {
        internal ZipEntry(string name, ushort flags, ushort method, uint crc, long compressedSize, long expandedSize,
            long localHeaderOffset, long? dataOffset) {
            Name = name;
            Flags = flags;
            Method = method;
            Crc = crc;
            CompressedSize = compressedSize;
            ExpandedSize = expandedSize;
            LocalHeaderOffset = localHeaderOffset;
            DataOffset = dataOffset;
        }

        internal string Name { get; }
        internal ushort Flags { get; }
        internal ushort Method { get; }
        internal uint Crc { get; }
        internal long CompressedSize { get; }
        /// <summary>Declared decoded size, or -1 when no size could be established.</summary>
        internal long ExpandedSize { get; }
        internal long LocalHeaderOffset { get; }
        internal long? DataOffset { get; }
    }

    private readonly struct Repair {
        internal Repair(long offset, string reason) {
            Offset = offset;
            Reason = reason;
        }

        internal long Offset { get; }
        internal string Reason { get; }
    }

    /// <summary>Read-only view of one byte range of the shared source.</summary>
    private sealed class SliceStream : Stream {
        private readonly Stream _source;
        private readonly long _start;
        private readonly long _length;
        private long _position;

        internal SliceStream(Stream source, long start, long length) {
            _source = source;
            _start = start;
            _length = length;
        }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => _length;
        public override long Position { get => _position; set => throw new NotSupportedException(); }

        public override int Read(byte[] buffer, int offset, int count) {
            long remaining = _length - _position;
            if (remaining <= 0 || count == 0) return 0;
            _source.Position = _start + _position;
            int read = _source.Read(buffer, offset, (int)Math.Min(count, remaining));
            _position += read;
            return read;
        }

        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }

    /// <summary>Counts decoded bytes and stops a damaged entry at the configured entry bound.</summary>
    private sealed class CountingStream : Stream {
        private readonly Stream _inner;
        private readonly long _maximum;

        internal CountingStream(Stream inner, long maximum) {
            _inner = inner;
            _maximum = maximum;
        }

        internal long BytesRead { get; private set; }
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => BytesRead; set => throw new NotSupportedException(); }

        public override int Read(byte[] buffer, int offset, int count) {
            if (BytesRead >= _maximum) throw new InvalidDataException("The damaged entry exceeds MaxArchiveEntryBytes.");
            int read = _inner.Read(buffer, offset, (int)Math.Min(count, _maximum - BytesRead));
            BytesRead += read;
            return read;
        }

        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
