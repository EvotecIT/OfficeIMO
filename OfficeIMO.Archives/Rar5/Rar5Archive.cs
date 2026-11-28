namespace OfficeIMO.Archives.Rar5;

/// <summary>Reads the primary RAR5 headers in supplied volume order.</summary>
internal sealed partial class Rar5Archive {
    internal const int MaxVolumeCount = 256;
    private readonly ArchiveLimits _limits;
    private readonly ArchiveBlockBudget _budget;
    private readonly ArchivePaths _paths;
    private readonly List<Member> _members = new();
    private readonly Dictionary<string, Member> _fileSources = new(StringComparer.Ordinal);
    private string? _comment;
    private (int Method, long Dictionary)? _commentCompression;
    private long _metadataBytes;
    private Member? _split;

    private Rar5Archive(ArchiveLimits limits) {
        _limits = limits;
        _budget = new ArchiveBlockBudget(limits);
        _paths = new ArchivePaths(limits, limits.Metadata);
    }

    internal static Rar5Archive Open(Stream source, ArchiveLimits limits, CancellationToken token) =>
        Open(new[] { source }, limits, token);

    internal static Rar5Archive Open(IReadOnlyList<Stream> volumes, ArchiveLimits limits, CancellationToken token) {
        var archive = new Rar5Archive(limits);
        for (int i = 0; i < volumes.Count; i++) {
            bool more = archive.WalkVolume(volumes[i], i, volumes.Count > 1, token);
            if (more != (i + 1 < volumes.Count))
                throw Location(i, volumes[i].Position, more ? "The next RAR5 volume is required." : "Unexpected volume after the end of the RAR5 set.");
        }
        if (archive._split != null) throw Location(volumes.Count - 1, volumes[volumes.Count - 1].Position, "The final RAR5 member is incomplete.");
        return archive;
    }

    private bool WalkVolume(Stream source, int volume, bool suppliedSet, CancellationToken token) {
        if (!Rar5Signature.TryFind(source, token)) throw new NotSupportedException("A RAR 5 signature was not found within the first 1 MiB.");
        var prefix = new byte[7];
        bool mainSeen = false;
        bool isVolume = false;
        bool firstFile = true;
        while (source.Position < source.Length) {
            token.ThrowIfCancellationRequested();
            long headerOffset = source.Position;
            try {
                Rar5Block block = Rar5Block.Read(source, prefix, _limits, _budget, ref _metadataBytes, token);
                long dataOffset = source.Position;
                if (block.Type == 4) throw new Rar5EncryptionNotSupportedException(headerOffset, null);
                if (!mainSeen && block.Type != 1) throw new InvalidDataException("The RAR5 archive is missing its main header.");
                if ((block.Flags & 24) != 0 && (!isVolume || block.Type != 2))
                    throw new InvalidDataException("The RAR5 split flags are unexpected on this header.");
                switch (block.Type) {
                    case 1:
                        if (mainSeen) throw new InvalidDataException("The RAR5 volume has more than one main header.");
                        isVolume = ValidateMain(block, volume, suppliedSet, token);
                        mainSeen = true;
                        break;
                    case 2:
                        if ((block.Flags & 8) != 0 && !firstFile)
                            throw new InvalidDataException("A continued RAR5 member must be the first file of its volume.");
                        ReadFile(source, block, headerOffset, dataOffset, token);
                        firstFile = false;
                        break;
                    case 3:
                        ReadService(source, block, volume, headerOffset, token);
                        break;
                    case 5:
                        bool more = (block.Fields.Number() & 1) != 0;
                        if ((!isVolume && more) || (!more && _split != null))
                            throw new InvalidDataException("The RAR5 end header has inconsistent volume flags.");
                        ValidateControlBlock(block, token);
                        return more;
                    default:
                        if ((block.Flags & 4) == 0) throw new NotSupportedException($"RAR5 header type {block.Type} is unsupported at byte offset {headerOffset}.");
                        break;
                }
                source.Position = dataOffset + block.PackedLength;
            } catch (InvalidDataException error) {
                throw Location(volume, headerOffset, error.Message, error);
            } catch (OfficeInspectionResourceLimitException error) {
                throw new OfficeInspectionResourceLimitException("archives", error.LimitName, error.Maximum,
                    error, FormattableString.Invariant($"RAR 5 volume {volume + 1}, byte offset {headerOffset}."));
            }
        }
        throw Location(volume, source.Position, "The RAR5 archive is missing its end header.");
    }

    private static bool ValidateMain(Rar5Block block, int volume, bool suppliedSet, CancellationToken token) {
        ulong flags = block.Fields.Number();
        bool isVolume = (flags & 1) != 0;
        bool hasIndex = (flags & 2) != 0;
        ulong storedIndex = hasIndex ? block.Fields.Number() : 0;
        if ((isVolume && storedIndex != (ulong)volume) || (!isVolume && (suppliedSet || hasIndex)))
            throw new InvalidDataException("The RAR5 volume number does not match the supplied order.");
        ValidateControlBlock(block, token);
        return isVolume;
    }

    private static void ValidateControlBlock(Rar5Block block, CancellationToken token) {
        if (block.PackedLength != 0) throw new InvalidDataException("A RAR5 control header declares unexpected file data.");
        block.Fields.Finish();
        while (block.Extras.Remaining != 0) {
            token.ThrowIfCancellationRequested();
            // Locator payloads are never followed. Only the primary header chain supplies file names and ranges.
            block.Extras.Slice(block.Extras.Size()).Number();
        }
    }

    private void ReadFile(Stream source, Rar5Block block, long headerOffset, long dataOffset, CancellationToken token) {
        var fields = new FileFields(block.Fields, _limits.PathLength);
        var extras = new FileExtras(block.Extras, fields, headerOffset, _limits.PathLength, token);
        if ((fields.Flags & 8) != 0) throw new NotSupportedException("RAR5 members with unknown unpacked sizes are unsupported.");
        bool fromPrevious = (block.Flags & 8) != 0;
        bool toNext = (block.Flags & 16) != 0;
        if (fromPrevious != (_split != null)) throw new InvalidDataException("The RAR5 split member has no matching continuation.");

        Member member;
        if (_split is { } prior) {
            if (prior.Fields.Name != fields.Name || prior.Fields.Size != fields.Size || prior.Fields.Coder != fields.Coder ||
                prior.Fields.Host != fields.Host || extras.Redirection != null ||
                GetKind(fields.Host, fields.Attributes, (fields.Flags & 1) != 0) != ArchiveEntryKind.File)
                throw new InvalidDataException("The RAR5 split member headers disagree.");
            member = prior;
        } else {
            ArchiveLimits.Check(_members.Count + 1L, _limits.Entries, nameof(ArchiveExtractionOptions.MaxEntryCount));
            _limits.Operation.AdmitEntries(1);
            member = CreateMember(fields, extras, _members.Count, _paths, _limits, _budget);
            _members.Add(member);
        }

        if ((fromPrevious || toNext) && (member.Kind != ArchiveEntryKind.File || member.Extras.Redirection != null || block.PackedLength == 0))
            throw new InvalidDataException("The RAR5 split member must be a file with data.");
        long packed = member.PackedLength;
        ArchiveLimits.Add(ref packed, block.PackedLength, _limits.Input, nameof(ArchiveExtractionOptions.MaxInputBytes));
        member.PackedLength = packed;
        member.Parts.Add(new Part(source, dataOffset, block.PackedLength, new Integrity(fields.Crc, extras.Hash, extras.UnknownHash)));
        _split = toNext ? member : null;
        if (!toNext) FinishMember(member);
    }

    private void ReadService(Stream source, Rar5Block block, int volume, long headerOffset, CancellationToken token) {
        var fields = new FileFields(block.Fields, _limits.PathLength);
        var extras = new FileExtras(block.Extras, fields, headerOffset, _limits.PathLength, token);
        if (fields.Name != "CMT") return;
        if (volume != 0 || _commentCompression.HasValue || _members.Count != 0 || fields.Size > 256 * 1024 ||
            (fields.Method == 0 && fields.Size != block.PackedLength))
            throw new InvalidDataException("The RAR5 archive comment has invalid placement or size.");
        _commentCompression = (fields.Method, fields.Dictionary);
        if (fields.Method != 0) return;
        _budget.Charge(0, Math.Max(64, fields.Size) * 2);
        var bytes = new byte[(int)fields.Size];
        Rar5HeaderReader.Fill(source, bytes, bytes.Length, token);
        var check = new Integrity(fields.Crc, extras.Hash, extras.UnknownHash);
        if (check.UnknownHash) throw new NotSupportedException("The RAR5 comment declares an unsupported hash.");
        if (bytes.Length != 0 && !check.IsVerifiable)
            throw new InvalidDataException("The RAR5 comment has no data checksum.");
        Verify(bytes, 0, bytes.Length, check, "archive comment", headerOffset, token);
        _comment = Rar5HeaderReader.DecodeText(bytes);
    }

    internal Rar5ArchiveInfo Describe(int volumeCount) =>
        new(Array.AsReadOnly(_members.Select(member => member.Describe()).ToArray()), volumeCount, _comment, _commentCompression);

    private static InvalidDataException Location(int volume, long offset, string message, Exception? error = null) =>
        new(FormattableString.Invariant($"RAR 5 volume {volume + 1}, byte offset {offset}: {message}"), error);
}
