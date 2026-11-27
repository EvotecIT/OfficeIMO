namespace OfficeIMO.Archives.SevenZip;

/// <summary>Parses the 7z container and associates authored files with slices of checked folder output.</summary>
internal sealed partial class SevenZipArchive {
    private const int SignatureSize = 32;
    private static readonly byte[] Magic = { (byte)'7', (byte)'z', 0xbc, 0xaf, 0x27, 0x1c };
    private readonly Stream _source;
    private readonly ArchiveLimits _limits;
    private readonly ArchiveBlockBudget _budget;
    private readonly SevenZipFolder[] _folders;
    private readonly Member[] _members;

    private SevenZipArchive(Stream source, ArchiveLimits limits, ArchiveBlockBudget budget, SevenZipFolder[] folders, Member[] members) {
        _source = source;
        _limits = limits;
        _budget = budget;
        _folders = folders;
        _members = members;
    }

    internal static bool HasSignature(byte[] bytes, int count) {
        if (count < Magic.Length) return false;
        for (int i = 0; i < Magic.Length; i++) {
            if (bytes[i] != Magic[i]) return false;
        }
        return true;
    }

    internal static SevenZipArchive Open(Stream source, ArchiveLimits limits, CancellationToken token) {
        source.Position = 0;
        var signature = new byte[SignatureSize];
        SevenZipInput.Fill(source, signature, 0, Magic.Length, token);
        if (!HasSignature(signature, Magic.Length)) throw new NotSupportedException("A 7z archive is required.");
        SevenZipInput.Fill(source, signature, Magic.Length, SignatureSize - Magic.Length, token);
        var start = new SevenZipInput(signature);
        start.Bytes(Magic.Length);
        if (start.Byte() != 0) throw new NotSupportedException("The 7z archive major version is unsupported.");
        start.Byte(); // All supported layouts use the same fixed start header.
        uint startCrc = start.U32();
        SevenZipInput.VerifyChecksum(signature, 12, 20, startCrc, "start header", token);
        long offset = start.I64();
        long size = start.I64();
        uint headerCrc = start.U32();
        long available = source.Length - SignatureSize;
        if (offset < 0 || size < 0 || offset > available || size > available - offset)
            throw SevenZipInput.Invalid("The 7z next header lies outside the archive.");
        var budget = new ArchiveBlockBudget(limits);
        if (size == 0) {
            if (offset != 0 || headerCrc != 0) throw SevenZipInput.Invalid("The empty 7z next header has an inconsistent offset or checksum.");
            return new SevenZipArchive(source, limits, budget, Array.Empty<SevenZipFolder>(), Array.Empty<Member>());
        }

        long maximumMetadata = Math.Min(limits.Metadata, 64L * 1024 * 1024);
        ArchiveLimits.Check(size, maximumMetadata, nameof(ArchiveExtractionOptions.MaxMetadataBytes));
        budget.Charge(0, size);
        long payloadEnd = SignatureSize + offset;
        source.Position = payloadEnd;
        var header = new byte[(int)size];
        SevenZipInput.Fill(source, header, 0, header.Length, token);
        SevenZipInput.VerifyChecksum(header, 0, header.Length, headerCrc, "next header", token);
        var record = new SevenZipInput(header);
        byte type = record.Byte();
        if (type == 0x17) {
            SevenZipFolder[] encoded = ReadStreams(record, payloadEnd, maximumFolders: 1, limits.Entries, token);
            record.Finish();
            if (encoded.Length != 1)
                throw SevenZipInput.Invalid("The 7z encoded header must have one folder.");
            payloadEnd = encoded[0].Offset;
            record = new SevenZipInput(encoded[0].Decode(source, maximumMetadata - size, nameof(ArchiveExtractionOptions.MaxMetadataBytes), budget, token));
            type = record.Byte();
        }
        if (type != 1) throw SevenZipInput.Invalid("The 7z next header has an unknown record type.");
        type = record.Byte();
        if (type == 2) {
            while (record.Byte() != 0) {
                token.ThrowIfCancellationRequested();
                record.Property().SkipRemaining();
            }
            type = record.Byte();
        }
        if (type == 3) throw new NotSupportedException("External 7z metadata streams are unsupported.");
        SevenZipFolder[] folders = Array.Empty<SevenZipFolder>();
        if (type == 4) {
            folders = ReadStreams(record, payloadEnd, limits.Entries, limits.Entries, token);
            type = record.Byte();
        }
        Member[] members = Array.Empty<Member>();
        if (type == 5) {
            members = ReadMembers(record, folders, limits, token);
            type = record.Byte();
        }
        if (type != 0) throw SevenZipInput.Invalid("The 7z header has an unexpected section.");
        if (members.Length == 0 && folders.Any(folder => folder.Extents.Length != 0))
            throw SevenZipInput.Invalid("The 7z header contains unassigned member streams.");
        record.Finish();
        return new SevenZipArchive(source, limits, budget, folders, members);
    }

    internal IReadOnlyList<SevenZipArchiveEntry> List() =>
        Array.AsReadOnly(_members.Select(member => member.Entry).ToArray());

    internal void Read(Action<string, byte[], int, int> consume, CancellationToken token) {
        var groups = new List<Member>?[_folders.Length];
        var empty = new List<Member>();
        long selectedBytes = 0;
        foreach (Member member in _members) {
            token.ThrowIfCancellationRequested();
            if (member.Entry.Kind == ArchiveEntryKind.Directory || !_limits.Select(member.Entry.RelativePath)) continue;
            if (member.Entry.Kind != ArchiveEntryKind.File)
                throw new NotSupportedException("A selected 7z member is a link or special file: " + member.Entry.Name);
            ArchiveLimits.Add(ref selectedBytes, member.Entry.Length, _limits.Expanded,
                nameof(ArchiveExtractionOptions.MaxExpandedBytes));
            if (member.Folder < 0) empty.Add(member);
            else (groups[member.Folder] ??= new List<Member>()).Add(member);
        }

        token.ThrowIfCancellationRequested();
        long expanded = 0;
        for (int i = 0; i < groups.Length; i++) {
            if (groups[i] != null) ArchiveLimits.Add(ref expanded, _folders[i].Size, _limits.Expanded,
                nameof(ArchiveExtractionOptions.MaxExpandedBytes));
        }
        _limits.Operation.Select(selectedBytes);
        foreach (Member member in empty) {
            token.ThrowIfCancellationRequested();
            consume(member.Entry.RelativePath, Array.Empty<byte>(), 0, 0);
        }
        for (int i = 0; i < groups.Length; i++) {
            if (groups[i] is not { } group) continue;
            byte[] decoded = _folders[i].Decode(_source, _limits.Expanded, nameof(ArchiveExtractionOptions.MaxExpandedBytes), _budget, token);
            foreach (Member member in group) {
                token.ThrowIfCancellationRequested();
                consume(member.Entry.RelativePath, decoded, (int)member.Offset, (int)member.Entry.Length);
            }
        }
        token.ThrowIfCancellationRequested();
    }

    private readonly struct Member {
        internal Member(SevenZipArchiveEntry entry, int folder, long offset) { Entry = entry; Folder = folder; Offset = offset; }

        internal SevenZipArchiveEntry Entry { get; }
        internal int Folder { get; }
        internal long Offset { get; }
    }
}
