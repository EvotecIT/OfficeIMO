using System.Threading;

namespace OfficeIMO.IWork.Internal;

internal sealed class IWorkObjectIndex {
    private readonly Dictionary<ulong, IWorkArchiveRecord> _objects;
    private readonly Dictionary<IWorkArchiveRecord, IWorkWireMessage> _messages;
    private readonly object _messageLock;
    private readonly IWorkReadOptions _options;
    private readonly CancellationToken _cancellationToken;

    internal IWorkObjectIndex(IReadOnlyList<IWorkArchiveRecord> records, IWorkReadOptions options,
        CancellationToken cancellationToken = default) {
        _options = options;
        _cancellationToken = cancellationToken;
        _messages = new();
        _messageLock = new();
        _objects = new Dictionary<ulong, IWorkArchiveRecord>();
        foreach (IWorkArchiveRecord record in records) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!record.IsPrimary) continue;
            if (_objects.ContainsKey(record.Identifier)) {
                throw new InvalidDataException(
                    $"More than one primary IWA record declares object identifier {record.Identifier}.");
            }
            _objects.Add(record.Identifier, record);
        }
    }

    private IWorkObjectIndex(IWorkObjectIndex source, CancellationToken cancellationToken) {
        _objects = source._objects;
        _messages = source._messages;
        _messageLock = source._messageLock;
        _options = source._options;
        _cancellationToken = cancellationToken;
    }

    internal IWorkObjectIndex WithCancellation(CancellationToken cancellationToken) =>
        new(this, cancellationToken);

    internal IEnumerable<IWorkArchiveRecord> PrimaryRecords => _objects.Values;

    internal IWorkWireMessage Message(IWorkArchiveRecord record) {
        _cancellationToken.ThrowIfCancellationRequested();
        lock (_messageLock) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (_messages.TryGetValue(record, out IWorkWireMessage? cached)) return cached;
            IWorkWireMessage parsed = IWorkProtobuf.Parse(record.Payload, _options);
            _messages.Add(record, parsed);
            return parsed;
        }
    }

    internal IWorkArchiveRecord? UniqueOfType(uint type, out bool duplicate) {
        IWorkArchiveRecord[] matches = _objects.Values
            .Where(record => record.MessageType == type)
            .Take(2)
            .ToArray();
        duplicate = matches.Length > 1;
        return matches.Length == 1 ? matches[0] : null;
    }

    internal IWorkArchiveRecord? Find(ulong identifier) =>
        _objects.TryGetValue(identifier, out IWorkArchiveRecord? record) ? record : null;

    internal IReadOnlyCollection<IWorkArchiveRecord> ReachableFrom(params IWorkArchiveRecord[] roots) {
        var result = new Dictionary<ulong, IWorkArchiveRecord>();
        var pending = new Stack<IWorkArchiveRecord>(roots);
        while (pending.Count > 0) {
            _cancellationToken.ThrowIfCancellationRequested();
            IWorkArchiveRecord record = pending.Pop();
            if (result.ContainsKey(record.Identifier)) continue;
            result.Add(record.Identifier, record);
            foreach (ulong reference in record.ObjectReferences) {
                if (_objects.TryGetValue(reference, out IWorkArchiveRecord? target)
                    && !result.ContainsKey(target.Identifier)) pending.Push(target);
            }
        }
        return result.Values.ToArray();
    }

    internal IWorkArchiveRecord? Dereference(IWorkWireMessage message, int field) {
        IWorkWireMessage? reference = TryGetMessage(message, field);
        if (reference == null || reference.FieldCount(1) != 1
            || reference.HasUnexpectedWireKind(1, IWorkWireKind.Varint)) return null;
        ulong? identifier = reference?.GetUnsigned(1);
        return identifier.HasValue && _objects.TryGetValue(identifier.Value, out IWorkArchiveRecord? record)
            ? record
            : null;
    }

    internal IReadOnlyList<IWorkArchiveRecord> DereferenceAll(IWorkWireMessage message, int field) {
        return DereferenceAll(message, field, out _);
    }

    internal IReadOnlyList<IWorkArchiveRecord> DereferenceAll(IWorkWireMessage message, int field,
        out int unresolvedReferenceCount) => DereferenceAll(message, field, out unresolvedReferenceCount, out _);

    internal IReadOnlyList<IWorkArchiveRecord> DereferenceAll(IWorkWireMessage message, int field,
        out int unresolvedReferenceCount, out bool rejectedReferenceSet, List<int>? resolvedPositions = null) {
        var result = new List<IWorkArchiveRecord>();
        resolvedPositions?.Clear();
        unresolvedReferenceCount = 0;
        IReadOnlyList<IWorkWireMessage> references = TryGetMessages(message, field, out bool malformed);
        rejectedReferenceSet = malformed;
        if (malformed) {
            unresolvedReferenceCount = message.FieldCount(field);
            return result;
        }
        int position = 0;
        foreach (IWorkWireMessage reference in references) {
            position++;
            if (reference.FieldCount(1) != 1
                || reference.HasUnexpectedWireKind(1, IWorkWireKind.Varint)) {
                unresolvedReferenceCount++;
                continue;
            }
            ulong? identifier = reference.GetUnsigned(1);
            if (identifier.HasValue && _objects.TryGetValue(identifier.Value, out IWorkArchiveRecord? record)) {
                result.Add(record);
                resolvedPositions?.Add(position);
            } else {
                unresolvedReferenceCount++;
            }
        }
        return result;
    }

    internal static IWorkWireMessage? TryGetMessage(IWorkWireMessage message, int field) =>
        TryGetMessage(message, field, out _);

    internal static IWorkWireMessage? TryGetMessage(IWorkWireMessage message, int field, out bool malformed) {
        try {
            malformed = message.FieldCount(field) > 1
                || message.HasUnexpectedWireKind(field, IWorkWireKind.Bytes);
            if (malformed) return null;
            return message.GetMessage(field);
        } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
            malformed = true;
            return null;
        }
    }

    internal static IReadOnlyList<IWorkWireMessage> TryGetMessages(IWorkWireMessage message, int field) {
        return TryGetMessages(message, field, out _);
    }

    internal static IReadOnlyList<IWorkWireMessage> TryGetMessages(IWorkWireMessage message, int field,
        out bool malformed) {
        try {
            malformed = message.HasUnexpectedWireKind(field, IWorkWireKind.Bytes);
            if (malformed) return Array.Empty<IWorkWireMessage>();
            return message.GetRepeatedMessages(field);
        } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
            malformed = true;
            return Array.Empty<IWorkWireMessage>();
        }
    }
}

internal static class IWorkArchiveParser {
    internal static IReadOnlyList<IWorkArchiveRecord> Parse(IReadOnlyList<IWorkPackageEntry> entries,
        IWorkReadOptions options, CancellationToken cancellationToken = default) {
        var records = new List<IWorkArchiveRecord>();
        long totalDecompressedBytes = 0;
        int archiveReferenceCount = 0;
        foreach (IWorkPackageEntry entry in entries.Where(candidate => IsIndexArchivePath(candidate.Path))) {
            cancellationToken.ThrowIfCancellationRequested();
            byte[] stream;
            try {
                long remaining = options.MaximumTotalDecompressedIwaBytes - totalDecompressedBytes;
                stream = IWorkSnappy.DecodeIwa(entry.Bytes, options, remaining, cancellationToken);
                totalDecompressedBytes = checked(totalDecompressedBytes + stream.LongLength);
                ParseStream(stream, entry.Path, records, options, ref archiveReferenceCount,
                    cancellationToken);
            } catch (Exception exception) when (exception is InvalidDataException or OverflowException) {
                throw new InvalidDataException($"Failed to read IWA entry {entry.Path}: {exception.Message}", exception);
            }
        }
        return records;
    }

    internal static bool IsIndexArchivePath(string path) =>
        path.StartsWith("Index/", StringComparison.OrdinalIgnoreCase)
        && path.EndsWith(".iwa", StringComparison.OrdinalIgnoreCase);

    private static void ParseStream(byte[] stream, string entryPath, List<IWorkArchiveRecord> records,
        IWorkReadOptions options, ref int archiveReferenceCount, CancellationToken cancellationToken) {
        int offset = 0;
        while (offset < stream.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            ulong rawInfoLength = IWorkProtobuf.ReadVarint(stream, ref offset);
            if (rawInfoLength > (ulong)options.MaximumArchiveInfoBytes || rawInfoLength > int.MaxValue) {
                throw new InvalidDataException($"ArchiveInfo length {rawInfoLength} exceeds the configured limit.");
            }
            int infoLength = (int)rawInfoLength;
            if (offset > stream.Length - infoLength) throw new InvalidDataException($"Truncated ArchiveInfo at offset {offset}.");
            byte[] infoBytes = Slice(stream, offset, infoLength);
            offset += infoLength;
            int messageCount = IWorkProtobuf.CountFields(
                infoBytes, 2, options.MaximumProtobufFieldCount);
            if (messageCount == 0) {
                throw new InvalidDataException("ArchiveInfo does not declare any payloads.");
            }
            if (messageCount > options.MaximumRecordCount - records.Count) {
                throw new InvalidDataException($"IWA record count exceeds the configured limit of {options.MaximumRecordCount}.");
            }
            IWorkWireMessage archiveInfo = IWorkProtobuf.Parse(infoBytes, options);
            ulong? identifier = archiveInfo.GetUnsigned(1);
            if (archiveInfo.FieldCount(1) != 1 || !identifier.HasValue
                || archiveInfo.HasUnexpectedWireKind(1, IWorkWireKind.Varint)) {
                throw new InvalidDataException("ArchiveInfo does not declare a valid object identifier.");
            }

            if (archiveInfo.HasUnexpectedWireKind(2, IWorkWireKind.Bytes)) {
                throw new InvalidDataException($"ArchiveInfo {identifier.Value} contains malformed MessageInfo entries.");
            }
            IReadOnlyList<IWorkWireMessage> messages = archiveInfo.GetRepeatedMessages(2);
            if (messages.Count != messageCount) {
                throw new InvalidDataException($"ArchiveInfo {identifier.Value} contains malformed MessageInfo entries.");
            }
            for (int payloadIndex = 0; payloadIndex < messages.Count; payloadIndex++) {
                cancellationToken.ThrowIfCancellationRequested();
                IWorkWireMessage messageInfo = messages[payloadIndex];
                if (messageInfo.FieldCount(1) != 1 || messageInfo.FieldCount(3) != 1
                    || messageInfo.HasUnexpectedWireKind(1, IWorkWireKind.Varint)
                    || messageInfo.HasUnexpectedWireKind(2, IWorkWireKind.Varint, IWorkWireKind.Bytes)
                    || messageInfo.HasUnexpectedWireKind(3, IWorkWireKind.Varint)
                    || messageInfo.HasUnexpectedWireKind(5, IWorkWireKind.Varint, IWorkWireKind.Bytes)
                    || messageInfo.HasUnexpectedWireKind(6, IWorkWireKind.Varint, IWorkWireKind.Bytes)) {
                    throw new InvalidDataException(
                        $"MessageInfo for object {identifier.Value} contains an invalid wire encoding.");
                }
                ulong? rawType = messageInfo.GetUnsigned(1);
                ulong? rawLength = messageInfo.GetUnsigned(3);
                if (!rawType.HasValue || rawType.Value > uint.MaxValue) {
                    throw new InvalidDataException($"MessageInfo for object {identifier.Value} has an invalid registry type.");
                }
                if (!rawLength.HasValue || rawLength.Value > (ulong)options.MaximumRecordBytes || rawLength.Value > int.MaxValue) {
                    throw new InvalidDataException($"MessageInfo for object {identifier.Value} has an invalid payload length.");
                }
                int payloadLength = (int)rawLength.Value;
                if (offset > stream.Length - payloadLength) {
                    throw new InvalidDataException($"Truncated payload for object {identifier.Value} at offset {offset}.");
                }
                int remainingArchiveReferences = options.MaximumArchiveReferenceCount
                    - archiveReferenceCount;
                IReadOnlyList<ulong> versions = messageInfo.GetRepeatedUnsigned(
                    2, packed: true, remainingArchiveReferences);
                archiveReferenceCount += versions.Count;
                remainingArchiveReferences = options.MaximumArchiveReferenceCount
                    - archiveReferenceCount;
                IReadOnlyList<ulong> objectReferences = messageInfo.GetRepeatedUnsigned(
                    5, packed: true, remainingArchiveReferences);
                archiveReferenceCount += objectReferences.Count;
                remainingArchiveReferences = options.MaximumArchiveReferenceCount
                    - archiveReferenceCount;
                IReadOnlyList<ulong> dataReferences = messageInfo.GetRepeatedUnsigned(
                    6, packed: true, remainingArchiveReferences);
                archiveReferenceCount += dataReferences.Count;
                records.Add(new IWorkArchiveRecord(
                    identifier.Value,
                    (uint)rawType.Value,
                    versions.Select(value => checked((uint)value)).ToArray(),
                    objectReferences,
                    dataReferences,
                    entryPath,
                    payloadIndex,
                    Slice(stream, offset, payloadLength)));
                offset += payloadLength;
            }
        }
    }

    private static byte[] Slice(byte[] source, int offset, int length) {
        var result = new byte[length];
        Buffer.BlockCopy(source, offset, result, 0, length);
        return result;
    }
}
