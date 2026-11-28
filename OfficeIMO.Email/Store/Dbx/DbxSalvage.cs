using System.Buffers.Binary;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Email.Store;

/// <summary>
/// Recovers messages from a damaged DBX message database by scanning for self-addressed message records and
/// fragments instead of trusting the active index. Each object in a DBX file starts with its own file offset, which
/// lets records and fragments be found even when the tree or the header is unreadable.
/// </summary>
internal sealed class DbxSalvage {
    private const int HeaderLength = DbxDatabase.FileHeaderLength;
    private const int TreeNodeLength = 0x27c;
    private const int TreeNodeHeaderLength = 0x18;
    private const int TreeEntryLength = 12;
    private const int RecordHeaderLength = 12;
    private const int FragmentHeaderLength = 16;
    private const int MaximumFragmentAllocation = 0x10000;
    private const int MaximumTreeDepth = 64;
    private const int ScanBufferLength = 1 << 20;
    private const string FolderId = "dbx:folder:root";
    private const string RecordPrefix = "dbx:item:";
    private const string ChainPrefix = "dbx:chain:";

    private readonly Stream _source;
    private readonly EmailStoreSalvageOptions _options;
    private readonly EmailStoreReaderOptions _limits;
    private readonly EmailStoreSalvageWriter _output;
    private readonly CancellationToken _cancellationToken;
    private readonly long _length;
    private readonly string _displayName;
    private readonly List<EmailStoreDiagnostic> _diagnostics = new List<EmailStoreDiagnostic>();
    private readonly List<uint> _records = new List<uint>();
    private readonly Dictionary<uint, Fragment> _fragments = new Dictionary<uint, Fragment>();
    private readonly HashSet<uint> _consumedFragments = new HashSet<uint>();
    // First fragments that an independent record or message header identifies; a link into one is a conflict.
    private readonly HashSet<uint> _bodyHeads = new HashSet<uint>();
    // Fragments whose payload opens with top-level message headers; they begin a message even when linked to.
    private readonly HashSet<uint> _messageStarts = new HashSet<uint>();
    private bool _stoppedAtLimit;
    private bool _itemLimitReached;
    private bool _scanStopped;

    internal DbxSalvage(Stream source, string? sourceName, EmailStoreSalvageOptions options,
        EmailStoreSalvageWriter output, CancellationToken cancellationToken) {
        _source = source;
        _options = options;
        _limits = options.ReaderOptions;
        _output = output;
        _cancellationToken = cancellationToken;
        _length = Math.Min(source.Length, uint.MaxValue);
        _displayName = string.IsNullOrWhiteSpace(sourceName)
            ? "Mailbox" : Path.GetFileNameWithoutExtension(sourceName) ?? "Mailbox";
    }

    internal EmailStoreSalvageReport Run() {
        ValidateHeader();
        HashSet<uint> active = ReadActiveIndex(out bool indexComplete);
        Scan();
        // An indexed record that the scan did not recognize is still reported, so its loss is visible.
        var scanned = new HashSet<uint>(_records);
        foreach (uint address in active) if (scanned.Add(address)) _records.Add(address);
        CollectBodyHeads(active);

        var paths = new EmailStoreExportPathBuilder(_output.Root,
            new[] { new EmailStoreFolderInfo(FolderId, null, _displayName, 0, 0) }, preserveHierarchy: false);
        var entries = new List<EmailStoreSalvageEntry>();
        // Bodies any record claims are excluded from the orphan pass; only committed bodies suppress duplicates.
        var firstFragments = new HashSet<uint>();
        var recoveredBodies = new HashSet<uint>();
        // Records the active index lists win when stale records point at the same payload.
        foreach (uint address in _records.OrderBy(address => active.Contains(address) ? 0 : 1).ThenBy(address => address)) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (AtItemLimit(entries.Count)) { StopAtItemLimit(); break; }
            bool? indexed = active.Contains(address) ? true : indexComplete ? false : null;
            EmailStoreSalvageEntry? entry = RecoverRecord(address, indexed, firstFragments, recoveredBodies, paths);
            if (entry != null) entries.Add(entry);
        }

        // A scan stopped at the node bound still recovers the chains it found; only the item bound ends the pass.
        if (_options.IncludeUnindexedFragmentChains && !_itemLimitReached) {
            HashSet<uint> referenced = ReferencedFragments();
            HashSet<uint> cyclic = FindCycleMembers();
            foreach (KeyValuePair<uint, Fragment> head in _fragments.OrderBy(pair => pair.Key)) {
                _cancellationToken.ThrowIfCancellationRequested();
                // A closed cycle has no unreferenced member, so its first header-bearing fragment is the head.
                // A linked fragment that opens with top-level message headers is reached through a damaged link.
                if ((referenced.Contains(head.Key) && !cyclic.Contains(head.Key) && !_messageStarts.Contains(head.Key)) ||
                    firstFragments.Contains(head.Key) ||
                    _consumedFragments.Contains(head.Key) || !StartsWithHeader(head.Key, head.Value)) continue;
                if (AtItemLimit(entries.Count)) { StopAtItemLimit(); break; }
                entries.Add(RecoverChain(head.Key, indexComplete ? false : null, paths));
            }
        }

        ReportUnattachedFragments();
        // Active records are processed first so they win duplicate payloads; outcomes are reported in file order.
        EmailStoreSalvageEntry[] ordered = entries.OrderBy(entry => entry.SourceOffset ?? long.MaxValue).ToArray();
        string? manifest = _output.WriteManifest(ordered, _diagnostics);
        return new EmailStoreSalvageReport(EmailStoreFormat.Dbx, _output.Root, manifest, _stoppedAtLimit,
            Array.AsReadOnly(ordered), _diagnostics.ToArray(), discoveryCompleted: !_scanStopped);
    }

    /// <summary>Fragments another fragment links to; a fragment that links to itself does not count.</summary>
    private HashSet<uint> ReferencedFragments() =>
        new HashSet<uint>(_fragments.Where(pair => pair.Value.Next != 0 && pair.Value.Next != pair.Key)
            .Select(pair => pair.Value.Next));

    /// <summary>Fragments that lie on a closed chain of links, found in one pass over the link graph.</summary>
    private HashSet<uint> FindCycleMembers() {
        var cyclic = new HashSet<uint>();
        var finished = new HashSet<uint>();
        var path = new List<uint>();
        var onPath = new Dictionary<uint, int>();
        foreach (uint start in _fragments.Keys) {
            if (finished.Contains(start)) continue;
            path.Clear();
            onPath.Clear();
            uint current = start;
            while (_fragments.TryGetValue(current, out Fragment fragment) && !finished.Contains(current)) {
                _cancellationToken.ThrowIfCancellationRequested();
                if (onPath.TryGetValue(current, out int first)) {
                    for (int index = first; index < path.Count; index++) cyclic.Add(path[index]);
                    break;
                }
                onPath.Add(current, path.Count);
                path.Add(current);
                if (fragment.Next == 0) break;
                current = fragment.Next;
            }
            foreach (uint member in path) finished.Add(member);
        }
        return cyclic;
    }

    /// <summary>
    /// Records the first fragment of every message that something independent identifies: an indexed record, a
    /// stale record whose body starts with a message header, or an unreferenced fragment that does. A chain that
    /// links into one of them stops there instead of splicing another message's bytes into its own.
    /// </summary>
    private void CollectBodyHeads(HashSet<uint> active) {
        var database = new DbxDatabase(_source, _limits);
        foreach (uint address in _records) {
            _cancellationToken.ThrowIfCancellationRequested();
            bool indexed = active.Contains(address);
            if (!TryReadRecordTable(address, requireAscending: !indexed, out byte[] table) || !HasProperty(table, 0x04)) continue;
            uint head;
            try {
                head = database.ReadIndexedRecord(address, _cancellationToken).ReadUInt32(0x04, _cancellationToken) ?? 0;
            } catch (Exception exception) when (exception is InvalidDataException || exception is EmailStoreLimitExceededException) {
                continue;
            }
            if (head != 0 && _fragments.TryGetValue(head, out Fragment fragment) &&
                (indexed || StartsWithHeader(head, fragment))) _bodyHeads.Add(head);
        }
        HashSet<uint> referenced = ReferencedFragments();
        foreach (KeyValuePair<uint, Fragment> pair in _fragments) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (IsMessageStart(pair.Key, pair.Value)) _messageStarts.Add(pair.Key);
            if (!referenced.Contains(pair.Key) && StartsWithHeader(pair.Key, pair.Value)) _bodyHeads.Add(pair.Key);
        }
    }

    private static readonly string[] TopLevelHeaders = {
        "from", "subject", "date", "message-id", "received", "return-path", "mime-version", "to", "reply-to"
    };

    /// <summary>
    /// Whether a fragment's payload opens with a header block that names a top-level message header. Quoted or
    /// forwarded text can look the same, so this is only a carving hint: it stops a chain that no record owns, and
    /// never breaks a chain that a message record points at.
    /// </summary>
    private bool IsMessageStart(uint address, Fragment fragment) {
        var bytes = new byte[Math.Min(1024, (int)fragment.Used)];
        int read = ReadAvailable(address + (long)FragmentHeaderLength, bytes, 0, bytes.Length);
        int line = 0;
        bool first = true;
        while (line < read) {
            int end = line;
            while (end < read && bytes[end] != (byte)'\n') end++;
            int length = end - line;
            if (length > 0 && bytes[line + length - 1] == (byte)'\r') length--;
            if (length == 0) return false;
            if (bytes[line] == (byte)' ' || bytes[line] == (byte)'\t') {
                if (first) return false;
            } else {
                int colon = 0;
                while (colon < length && bytes[line + colon] > 0x20 && bytes[line + colon] < 0x7f && bytes[line + colon] != (byte)':') colon++;
                if (colon == 0 || colon == length || bytes[line + colon] != (byte)':') return false;
                string name = Encoding.ASCII.GetString(bytes, line, colon);
                foreach (string header in TopLevelHeaders)
                    if (string.Equals(header, name, StringComparison.OrdinalIgnoreCase)) return true;
            }
            first = false;
            line = end + 1;
        }
        return false;
    }

    /// <summary>Both the outcome bound and the source item bound stop recovery.</summary>
    private bool AtItemLimit(int count) => count >= _options.MaxRecoveredItems || count >= _limits.MaxItemCount;

    private void ValidateHeader() {
        var signature = new byte[24];
        int read = ReadAvailable(0, signature, 0, signature.Length);
        if (DbxDatabase.HasSignature(signature, read, folderDatabase: true))
            throw DbxDatabase.Invalid("Folders.dbx is a folder catalog and holds no messages to recover.");
        if (!DbxDatabase.HasSignature(signature, read))
            Add("EMAIL_STORE_DBX_SALVAGE_SIGNATURE_DAMAGED",
                "The message database signature is damaged or missing; records are recovered by scanning.",
                EmailStoreDiagnosticSeverity.Warning, 0);
        if (_length < HeaderLength) _scanStopped = true;
        if (_length < HeaderLength)
            Add("EMAIL_STORE_DBX_SALVAGE_HEADER_TRUNCATED",
                "The database ends inside its header, so no message records follow it.",
                EmailStoreDiagnosticSeverity.Error, _length);
    }

    private HashSet<uint> ReadActiveIndex(out bool complete) {
        var records = new HashSet<uint>();
        complete = false;
        if (_length < HeaderLength) return records;
        var header = new byte[0xf0];
        if (ReadAvailable(0, header, 0, header.Length) < header.Length) return records;
        uint expected = UInt32(header, 0xc4);
        uint root = UInt32(header, 0xe4);
        try {
            if (root != 0) ReadTree(root, 0, 1, records, new HashSet<uint>());
            if (records.Count != expected) {
                Add("EMAIL_STORE_DBX_SALVAGE_INDEX_COUNT_MISMATCH",
                    "The active index lists " + records.Count.ToString(CultureInfo.InvariantCulture) +
                    " record(s) but the header declares " + expected.ToString(CultureInfo.InvariantCulture) + ".",
                    EmailStoreDiagnosticSeverity.Warning, 0xc4);
                return records;
            }
            complete = true;
        } catch (TreeDamageException damage) {
            Add("EMAIL_STORE_DBX_SALVAGE_INDEX_DAMAGED",
                "The active index is unreadable at this offset: " + damage.Message +
                " Message records are recovered by scanning.", EmailStoreDiagnosticSeverity.Warning, damage.Offset);
        }
        return records;
    }

    private void ReadTree(uint address, uint parent, int depth, HashSet<uint> records, HashSet<uint> nodes) {
        _cancellationToken.ThrowIfCancellationRequested();
        if (depth > Math.Min(_limits.MaxBTreeDepth, MaximumTreeDepth))
            throw new TreeDamageException("The tree is deeper than the configured bound.", address);
        if (nodes.Count >= _limits.MaxNodeCount)
            throw new TreeDamageException("Index traversal stopped at " + nameof(EmailStoreReaderOptions.MaxNodeCount) + ".", address);
        if (!nodes.Add(address)) throw new TreeDamageException("A tree node repeats.", address);
        var node = new byte[TreeNodeLength];
        if (address < HeaderLength || ReadAvailable(address, node, 0, node.Length) < node.Length)
            throw new TreeDamageException("A tree node lies outside the file.", address);
        if (UInt32(node, 0) != address) throw new TreeDamageException("A tree node marker does not match its address.", address);
        if (UInt32(node, 12) != parent) throw new TreeDamageException("A tree node has an invalid parent.", address);
        int entryCount = node[17];
        if (entryCount > (TreeNodeLength - TreeNodeHeaderLength) / TreeEntryLength)
            throw new TreeDamageException("A tree node declares too many entries.", address);
        uint firstChild = UInt32(node, 8);
        if (firstChild != 0) ReadTree(firstChild, address, depth + 1, records, nodes);
        for (int entry = 0; entry < entryCount; entry++) {
            int offset = TreeNodeHeaderLength + entry * TreeEntryLength;
            uint record = UInt32(node, offset);
            if (record < HeaderLength || record >= _length)
                throw new TreeDamageException("A tree entry points outside the file.", address);
            if (records.Count >= _limits.MaxItemCount)
                throw new TreeDamageException("The index exceeds the configured item bound.", address);
            records.Add(record);
            uint child = UInt32(node, offset + 4);
            if (child != 0) ReadTree(child, address, depth + 1, records, nodes);
        }
    }

    private void Scan() {
        if (_length < HeaderLength + FragmentHeaderLength) return;
        var buffer = new byte[ScanBufferLength];
        long position = HeaderLength;
        while (position + FragmentHeaderLength <= _length) {
            _cancellationToken.ThrowIfCancellationRequested();
            int count = (int)Math.Min(buffer.Length, _length - position);
            int expected = count;
            count = ReadAvailable(position, buffer, 0, count);
            if (count != expected) throw new IOException("The DBX source changed length during recovery.");
            bool last = position + count >= _length;
            int end = last ? count - FragmentHeaderLength : count - FragmentHeaderLength - 4;
            for (int index = 0; index <= end; index += 4) {
                uint address = (uint)(position + index);
                if (UInt32(buffer, index) != address) continue;
                uint second = UInt32(buffer, index + 4);
                uint third = UInt32(buffer, index + 8);
                uint fourth = UInt32(buffer, index + 12);
                bool fragmentShape = second >= 1 && second <= MaximumFragmentAllocation && third >= 1 && third <= second;
                bool linkValid = fourth != address && (fourth == 0 || (fourth >= HeaderLength && fourth % 4 == 0));
                bool recordShape = buffer[index + 10] > 0 && buffer[index + 10] * 4L <= second &&
                    second <= _limits.MaxDecodedPropertyBytesPerItem;
                // A fragment with a broken link still holds readable bytes; the chain walk reports the link. The
                // link only decides the shape when the header could also be a message record.
                if (fragmentShape && (linkValid || !recordShape)) {
                    if (_fragments.Count + _records.Count >= _limits.MaxNodeCount) { StopAtNodeLimit(address); return; }
                    _fragments[address] = new Fragment(second, third, fourth);
                } else if (recordShape) {
                    if (_fragments.Count + _records.Count >= _limits.MaxNodeCount) { StopAtNodeLimit(address); return; }
                    _records.Add(address);
                }
            }
            if (last) break;
            position += end + 4;
        }
    }

    private EmailStoreSalvageEntry? RecoverRecord(uint address, bool? indexed, HashSet<uint> firstFragments,
        HashSet<uint> recoveredBodies,
        EmailStoreExportPathBuilder paths) {
        string id = RecordPrefix + address.ToString("x8", CultureInfo.InvariantCulture);
        string location = "record 0x" + address.ToString("x", CultureInfo.InvariantCulture);
        var diagnostics = new List<EmailStoreDiagnostic>();
        if (!TryReadRecordTable(address, requireAscending: indexed != true, out byte[] table)) {
            // Candidates that do not hold a valid message property table are not message records.
            if (indexed == true)
                return NotRecovered(id, location, address, indexed, address, "EMAIL_STORE_DBX_SALVAGE_RECORD_INVALID",
                    "The indexed message record has no readable property table.", diagnostics);
            return null;
        }
        if (!HasProperty(table, 0x04)) {
            if (indexed == true)
                return NotRecovered(id, location, address, indexed, address, "EMAIL_STORE_DBX_SALVAGE_NO_PAYLOAD",
                    "The indexed message record has no message body property.", diagnostics);
            return null;
        }

        var database = new DbxDatabase(_source, _limits);
        uint? flags = null;
        uint firstFragment = 0;
        string? subject = null;
        long? declared = null;
        DateTimeOffset? received = null;
        try {
            DbxIndexedRecord record = database.ReadIndexedRecord(address, _cancellationToken);
            firstFragment = record.ReadUInt32(0x04, _cancellationToken) ?? 0;
            flags = TryRead(() => record.ReadUInt32(0x01, _cancellationToken), id, diagnostics);
            subject = TryRead(() => record.ReadString(0x05, _cancellationToken, _limits.MaxDecodedPropertyBytesPerItem), id, diagnostics);
            declared = TryRead(() => record.ReadUInt32(0x11, _cancellationToken), id, diagnostics);
            ulong? fileTime = TryRead(() => record.ReadUInt64(0x12, _cancellationToken), id, diagnostics);
            if (fileTime.HasValue && fileTime.Value != 0 && fileTime.Value <= (ulong)DateTimeOffset.MaxValue.ToFileTime())
                received = new DateTimeOffset(DateTime.FromFileTimeUtc((long)fileTime.Value));
        } catch (Exception exception) when (exception is InvalidDataException || exception is EmailStoreLimitExceededException) {
            return NotRecovered(id, location, address, indexed, address, "EMAIL_STORE_DBX_SALVAGE_RECORD_INVALID",
                exception.Message, diagnostics);
        }
        if (firstFragment == 0)
            return NotRecovered(id, location, address, indexed, address, "EMAIL_STORE_DBX_SALVAGE_NO_PAYLOAD",
                "The message record has no local message body.", diagnostics, subject, declared);
        firstFragments.Add(firstFragment);
        if (recoveredBodies.Contains(firstFragment)) {
            _diagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_DBX_SALVAGE_DUPLICATE_RECORD",
                "A stale message record points at a message body that another record already recovered.",
                EmailStoreDiagnosticSeverity.Information, location));
            return null;
        }
        EmailStoreSalvageEntry result = Recover(id, location, address, firstFragment, indexed, subject, declared,
            flags, received, isChain: false, paths, diagnostics);
        if (result.DestinationPath != null) recoveredBodies.Add(firstFragment);
        return result;
    }

    private EmailStoreSalvageEntry RecoverChain(uint head, bool? indexed, EmailStoreExportPathBuilder paths) {
        string id = ChainPrefix + head.ToString("x8", CultureInfo.InvariantCulture);
        string location = "fragment chain 0x" + head.ToString("x", CultureInfo.InvariantCulture);
        return Recover(id, location, head, head, indexed, subject: null, declared: null, flags: null,
            received: null, isChain: true, paths, new List<EmailStoreDiagnostic>());
    }

    private EmailStoreSalvageEntry Recover(string id, string location, uint sourceOffset, uint firstFragment,
        bool? indexed, string? subject, long? declared, uint? flags, DateTimeOffset? received, bool isChain,
        EmailStoreExportPathBuilder paths, List<EmailStoreDiagnostic> diagnostics) {
        var reference = new EmailStoreItemReference(id, FolderId, isAssociated: false, isOrphaned: isChain);
        string? spillPath = null;
        try {
            ChainResult chain;
            try {
                using FileStream spill = _output.CreateSpill(".eml", deleteOnClose: false);
                spillPath = spill.Name;
                chain = CopyChain(firstFragment, spill, carving: isChain, declared);
            } catch (Exception exception) when (exception is IOException || exception is UnauthorizedAccessException) {
                diagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_SALVAGE_ITEM_WRITE_FAILED", exception.Message,
                    EmailStoreDiagnosticSeverity.Error, id));
                return new EmailStoreSalvageEntry(id, EmailStoreSalvageStatus.NotRecovered, location, sourceOffset, 0,
                    declared, indexed, null, "EMAIL_STORE_SALVAGE_ITEM_WRITE_FAILED", OutlookItemKind.Message, subject,
                    null, 0, diagnostics.AsReadOnly());
            }
            if (chain.DamageCode == null && declared.HasValue && declared.Value > chain.Bytes && chain.Bytes > 0) {
                // The record declares more MIME bytes than the chain holds, so its end-of-chain link came too early.
                chain = ChainResult.Damaged(chain.Bytes, chain.LastFragment ?? sourceOffset,
                    "EMAIL_STORE_DBX_SALVAGE_CHAIN_SHORT",
                    "The message fragment chain ends after " + chain.Bytes.ToString(CultureInfo.InvariantCulture) +
                    " of the " + declared.Value.ToString(CultureInfo.InvariantCulture) + " bytes the record declares.");
            }
            if (chain.DamageCode != null) {
                diagnostics.Add(new EmailStoreDiagnostic(chain.DamageCode, chain.DamageMessage!,
                    chain.Bytes > 0 ? EmailStoreDiagnosticSeverity.Warning : EmailStoreDiagnosticSeverity.Error,
                    location, "salvage", chain.DamageOffset, null, null, null, EmailDiagnosticDisposition.Recovered,
                    EmailDataLossRisk.Confirmed, "Recover the remaining message content from another copy of the mailbox."));
            }
            if (chain.Bytes == 0)
                return NotRecovered(id, location, sourceOffset, indexed, chain.DamageOffset, chain.DamageCode,
                    chain.DamageMessage ?? "The message body is empty.", diagnostics, subject, declared);

            EmailFileFormat format = _options.Format ?? EmailFileFormat.Eml;
            long? damageOffset = chain.DamageOffset;
            string? damageCode = chain.DamageCode;
            string? destination = null;
            long written = 0;
            try {
                if (format == EmailFileFormat.Eml && chain.Bytes > _options.WriterOptions.MaxOutputBytes) {
                    diagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_SALVAGE_ITEM_LIMIT",
                        "The recovered message is larger than " + nameof(EmailWriterOptions.MaxOutputBytes) + ".",
                        EmailStoreDiagnosticSeverity.Error, id, "salvage", null, nameof(EmailWriterOptions.MaxOutputBytes),
                        chain.Bytes, _options.WriterOptions.MaxOutputBytes, EmailDiagnosticDisposition.Skipped,
                        EmailDataLossRisk.Confirmed, "Raise MaxOutputBytes to write this message."));
                } else if (format == EmailFileFormat.Eml) {
                    subject ??= ReadSubject(spillPath!, id, diagnostics);
                    string path = _output.ResolvePath(paths, reference, subject, format);
                    if (_output.CommitRaw(spillPath!, path, id, diagnostics)) {
                        destination = path;
                        written = chain.Bytes;
                    }
                } else {
                    using var input = new FileStream(spillPath!, FileMode.Open, FileAccess.Read, FileShare.Read);
                    using EmailReadResult result = EmailStoreMessageReader.Read(input, _limits, _cancellationToken,
                        includeAttachmentContent: true);
                    EmailDocument document = result.Document;
                    // A converted artifact carries what decoding lost; the raw EML path keeps the bytes unchanged.
                    foreach (EmailDiagnostic diagnostic in result.Diagnostics)
                        diagnostics.Add(EmailStoreSession.ConvertDiagnostic(diagnostic, id));
                    if (result.HasErrors) damageCode ??= "EMAIL_STORE_DBX_SALVAGE_MIME_DAMAGED";
                    if (document.Subject == null && subject != null) document.Subject = subject;
                    document.ReceivedDate ??= received;
                    if (flags.HasValue) {
                        document.MessageMetadata.IsRead = (flags.Value & 0x80) != 0;
                        document.MessageMetadata.IsDraft = (flags.Value & 0x08) != 0;
                    }
                    subject = document.Subject ?? subject;
                    string path = _output.ResolvePath(paths, reference, subject, format);
                    long bytes = _output.WriteDocument(document, path, format, id, diagnostics);
                    if (bytes >= 0) {
                        destination = path;
                        written = bytes;
                    }
                }
            } catch (Exception exception) when (exception is InvalidDataException || exception is IOException ||
                                                 exception is UnauthorizedAccessException || exception is NotSupportedException ||
                                                 exception is EmailStoreLimitExceededException || exception is EmailLimitExceededException) {
                diagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_SALVAGE_ITEM_WRITE_FAILED", exception.Message,
                    EmailStoreDiagnosticSeverity.Error, id));
            }
            if (destination == null)
                return new EmailStoreSalvageEntry(id, EmailStoreSalvageStatus.NotRecovered, location, sourceOffset,
                    chain.Bytes, declared, indexed, damageOffset, damageCode ?? "EMAIL_STORE_SALVAGE_ITEM_WRITE_FAILED",
                    OutlookItemKind.Message, subject, null, 0, diagnostics.AsReadOnly());
            return new EmailStoreSalvageEntry(id,
                damageCode == null ? EmailStoreSalvageStatus.Recovered : EmailStoreSalvageStatus.Partial,
                location, sourceOffset, chain.Bytes, declared, indexed, damageOffset, damageCode,
                OutlookItemKind.Message, subject, destination, written, diagnostics.AsReadOnly());
        } finally {
            OfficeFileCommit.DeleteIfExists(spillPath);
        }
    }

    /// <summary>
    /// Copies a fragment chain. A carved chain stops before any fragment that opens with message headers; a
    /// record's chain stops there only once it holds the size the record declares, because headers inside a
    /// message (a forwarded one, say) can begin a fragment too.
    /// </summary>
    private ChainResult CopyChain(uint first, Stream output, bool carving, long? declared) {
        var seen = new HashSet<uint>();
        var header = new byte[FragmentHeaderLength];
        var buffer = new byte[MaximumFragmentAllocation];
        long bytes = 0;
        uint address = first;
        while (true) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (address < HeaderLength || address % 4 != 0)
                return ChainResult.Damaged(bytes, address, "EMAIL_STORE_DBX_SALVAGE_FRAGMENT_POINTER_INVALID",
                    "A message fragment link points into the header or at an unaligned offset.");
            if (address + (long)FragmentHeaderLength > _length)
                return ChainResult.Damaged(bytes, address, "EMAIL_STORE_DBX_SALVAGE_FRAGMENT_TRUNCATED",
                    "A message fragment link points past the end of the file.");
            if (seen.Count >= _limits.MaxNodeCount)
                return ChainResult.Damaged(bytes, address, "EMAIL_STORE_DBX_SALVAGE_MESSAGE_LIMIT",
                    "The message fragment chain reached " + nameof(EmailStoreReaderOptions.MaxNodeCount) + ".");
            if (!seen.Add(address))
                return ChainResult.Damaged(bytes, address, "EMAIL_STORE_DBX_SALVAGE_FRAGMENT_CYCLE",
                    "A message fragment chain links back to an earlier fragment.");
            ReadExactly(address, header, FragmentHeaderLength);
            uint allocated = UInt32(header, 4);
            uint used = UInt32(header, 8);
            if (UInt32(header, 0) != address || allocated == 0 || allocated > MaximumFragmentAllocation ||
                used == 0 || used > allocated)
                return ChainResult.Damaged(bytes, address, "EMAIL_STORE_DBX_SALVAGE_FRAGMENT_INVALID",
                    "A message fragment link does not point at a valid fragment.");
            _consumedFragments.Add(address);
            long payload = address + (long)FragmentHeaderLength;
            int available = (int)Math.Min(used, _length - payload);
            if (bytes + available > _limits.MaxMessageBytes)
                return ChainResult.Damaged(bytes, address, "EMAIL_STORE_DBX_SALVAGE_MESSAGE_LIMIT",
                    "The recovered message reached " + nameof(EmailStoreReaderOptions.MaxMessageBytes) + ".");
            if (available > 0) {
                ReadExactly(payload, buffer, available);
                output.Write(buffer, 0, available);
                bytes += available;
            }
            if (available < used)
                return ChainResult.Damaged(bytes, payload + available, "EMAIL_STORE_DBX_SALVAGE_FRAGMENT_TRUNCATED",
                    "The file ends inside a message fragment.");
            uint next = UInt32(header, 12);
            if (next == 0) return new ChainResult(bytes, null, null, null, address);
            if (!seen.Contains(next) && (_bodyHeads.Contains(next) || _consumedFragments.Contains(next) ||
                                         ((carving || (declared.HasValue && bytes >= declared.Value)) &&
                                          _messageStarts.Contains(next))))
                return ChainResult.Damaged(bytes, next, "EMAIL_STORE_DBX_SALVAGE_FRAGMENT_CONFLICT",
                    "A message fragment links into another message's fragments, so the chain stops before them.");
            address = next;
        }
    }

    private bool StartsWithHeader(uint address, Fragment fragment) {
        var bytes = new byte[Math.Min(80, (int)fragment.Used)];
        int read = ReadAvailable(address + (long)FragmentHeaderLength, bytes, 0, bytes.Length);
        int index = 0;
        while (index < read && bytes[index] > 0x20 && bytes[index] < 0x7f && bytes[index] != (byte)':') index++;
        return index > 0 && index < read && bytes[index] == (byte)':';
    }

    private string? ReadSubject(string path, string id, List<EmailStoreDiagnostic> diagnostics) {
        try {
            using var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            var envelope = new List<EmailDiagnostic>();
            return MimeHeaderParser.ReadEnvelope(input, EmailStoreMessageReader.CreateOptions(_limits), envelope,
                _cancellationToken).Subject;
        } catch (Exception exception) when (exception is InvalidDataException || exception is EmailLimitExceededException) {
            diagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_DBX_SALVAGE_HEADER_UNREADABLE", exception.Message,
                EmailStoreDiagnosticSeverity.Warning, id));
            return null;
        }
    }

    /// <summary>
    /// Reads a record's property table. Scan-only candidates must list properties in ascending order, which
    /// rejects coincidental matches; index-listed records only need a table within bounds, and the record reader
    /// rejects an ambiguous body pointer when it selects it.
    /// </summary>
    private bool TryReadRecordTable(uint address, bool requireAscending, out byte[] table) {
        table = Array.Empty<byte>();
        var header = new byte[RecordHeaderLength];
        if (ReadAvailable(address, header, 0, header.Length) < header.Length || UInt32(header, 0) != address) return false;
        uint bodyLength = UInt32(header, 4);
        int count = header[10];
        if (count == 0 || count > _limits.MaxPropertiesPerItem || count * 4L > bodyLength ||
            address + (long)RecordHeaderLength + bodyLength > _length) return false;
        var bytes = new byte[count * 4];
        if (ReadAvailable(address + (long)RecordHeaderLength, bytes, 0, bytes.Length) < bytes.Length) return false;
        int previous = -1;
        for (int entry = 0; entry < bytes.Length; entry += 4) {
            int property = bytes[entry] & 0x7f;
            if (requireAscending && property <= previous) return false;
            previous = property;
            // Index-listed records only need a readable body pointer; a damaged optional value is reported when read.
            if ((bytes[entry] & 0x80) == 0 && (requireAscending || property == 0x04) &&
                (UInt32(bytes, entry) >> 8) >= bodyLength - bytes.Length) return false;
        }
        table = bytes;
        return true;
    }

    private static bool HasProperty(byte[] table, int property) {
        for (int entry = 0; entry < table.Length; entry += 4)
            if ((table[entry] & 0x7f) == property) return true;
        return false;
    }

    private T? TryRead<T>(Func<T?> read, string id, List<EmailStoreDiagnostic> diagnostics) {
        try {
            return read();
        } catch (Exception exception) when (exception is InvalidDataException || exception is EmailStoreLimitExceededException) {
            diagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_DBX_SALVAGE_PROPERTY_UNREADABLE", exception.Message,
                EmailStoreDiagnosticSeverity.Warning, id));
            return default;
        }
    }

    private void ReportUnattachedFragments() {
        int count = 0;
        uint first = 0;
        foreach (uint address in _fragments.Keys) {
            if (_consumedFragments.Contains(address)) continue;
            if (count == 0 || address < first) first = address;
            count++;
        }
        if (count == 0) return;
        Add("EMAIL_STORE_DBX_SALVAGE_UNATTACHED_FRAGMENTS",
            count.ToString(CultureInfo.InvariantCulture) +
            " message fragment(s) belong to no recovered message; they may be the rest of a broken chain or freed space.",
            EmailStoreDiagnosticSeverity.Warning, first);
    }

    private EmailStoreSalvageEntry NotRecovered(string id, string location, uint offset, bool? indexed,
        long? damageOffset, string? damageCode, string message, List<EmailStoreDiagnostic> diagnostics,
        string? subject = null, long? declared = null) {
        if (!diagnostics.Any(diagnostic => diagnostic.Code == damageCode))
            diagnostics.Add(new EmailStoreDiagnostic(damageCode ?? "EMAIL_STORE_SALVAGE_ITEM_NOT_RECOVERED", message,
                EmailStoreDiagnosticSeverity.Error, location));
        return new EmailStoreSalvageEntry(id, EmailStoreSalvageStatus.NotRecovered, location, offset, 0, declared,
            indexed, damageOffset, damageCode, OutlookItemKind.Message, subject, null, 0, diagnostics.AsReadOnly());
    }

    private void StopAtItemLimit() {
        _stoppedAtLimit = true;
        _itemLimitReached = true;
        Add("EMAIL_STORE_SALVAGE_ITEM_LIMIT", "Recovery stopped at the configured item bound.",
            EmailStoreDiagnosticSeverity.Warning, null);
    }

    private void StopAtNodeLimit(uint address) {
        _stoppedAtLimit = true;
        _scanStopped = true;
        Add("EMAIL_STORE_SALVAGE_NODE_LIMIT",
            "Scanning stopped at " + nameof(EmailStoreReaderOptions.MaxNodeCount) + "; later records were not examined.",
            EmailStoreDiagnosticSeverity.Warning, address);
    }

    private void Add(string code, string message, EmailStoreDiagnosticSeverity severity, long? offset) =>
        _diagnostics.Add(new EmailStoreDiagnostic(code, message, severity,
            offset.HasValue ? "offset 0x" + offset.Value.ToString("x", CultureInfo.InvariantCulture) : null,
            "salvage", offset, null, null, null, EmailDiagnosticDisposition.Recovered,
            severity == EmailStoreDiagnosticSeverity.Error ? EmailDataLossRisk.Confirmed : EmailDataLossRisk.Possible, null));

    private int ReadAvailable(long offset, byte[] buffer, int start, int count) {
        if (offset >= _length) return 0;
        count = (int)Math.Min(count, _length - offset);
        _source.Position = offset;
        int total = 0;
        while (total < count) {
            int read = _source.Read(buffer, start + total, count - total);
            if (read == 0) break;
            total += read;
        }
        return total;
    }

    private void ReadExactly(long offset, byte[] buffer, int count) {
        if (ReadAvailable(offset, buffer, 0, count) != count)
            throw new IOException("The DBX source changed length during recovery.");
    }

    private static uint UInt32(byte[] bytes, int offset) =>
        BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(offset, 4));

    private readonly struct Fragment {
        internal Fragment(uint allocated, uint used, uint next) {
            Allocated = allocated;
            Used = used;
            Next = next;
        }

        internal uint Allocated { get; }
        internal uint Used { get; }
        internal uint Next { get; }
    }

    private readonly struct ChainResult {
        internal ChainResult(long bytes, long? damageOffset, string? damageCode, string? damageMessage,
            long? lastFragment = null) {
            LastFragment = lastFragment;
            Bytes = bytes;
            DamageOffset = damageOffset;
            DamageCode = damageCode;
            DamageMessage = damageMessage;
        }

        internal long Bytes { get; }
        /// <summary>The fragment whose zero link ended a complete chain.</summary>
        internal long? LastFragment { get; }
        internal long? DamageOffset { get; }
        internal string? DamageCode { get; }
        internal string? DamageMessage { get; }

        internal static ChainResult Damaged(long bytes, long offset, string code, string message) =>
            new ChainResult(bytes, offset, code, message);
    }

    private sealed class TreeDamageException : Exception {
        internal TreeDamageException(string message, long offset) : base(message) { Offset = offset; }
        internal long Offset { get; }
    }
}
