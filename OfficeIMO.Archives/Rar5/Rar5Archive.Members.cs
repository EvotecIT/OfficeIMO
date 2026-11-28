namespace OfficeIMO.Archives.Rar5;

internal sealed partial class Rar5Archive {
    private sealed class Member {
        internal Member(FileFields fields, FileExtras extras, string path, ArchiveEntryKind kind, int ordinal) {
            Fields = fields;
            Extras = extras;
            Path = path;
            Kind = kind;
            Ordinal = ordinal;
        }

        internal FileFields Fields { get; }
        internal FileExtras Extras { get; }
        internal string Path { get; }
        internal ArchiveEntryKind Kind { get; }
        internal int Ordinal { get; }
        internal List<Part> Parts { get; } = new();
        internal long PackedLength { get; set; }
        internal Member? CopySource { get; set; }

        internal string? UnsupportedReason {
            get {
                if (Fields.Method != 0) return "The entry uses compression.";
                if (Parts.Any(part => part.Integrity.UnknownHash)) return "The RAR5 member declares an unsupported file hash.";
                if (CopySource != null) return CopySource.UnsupportedReason;
                if (Kind != ArchiveEntryKind.File || Extras.Redirection != null)
                    return "The RAR5 reader cannot materialize this link or special file.";
                if (Fields.Size != 0 && !Parts[Parts.Count - 1].Integrity.IsVerifiable) return "The RAR5 member is missing a data checksum.";
                return null;
            }
        }

        internal Rar5ArchiveEntry Describe() =>
            new(Ordinal, Fields.Name, Path, Kind, Fields.Size, PackedLength, Fields.Method, Fields.Dictionary,
                Extras.Times, Extras.Owner, Extras.Redirection, UnsupportedReason == null);
    }

    private readonly struct Part {
        internal Part(Stream source, long offset, long length, Integrity integrity) {
            Source = source;
            Offset = offset;
            Length = length;
            Integrity = integrity;
        }
        internal Stream Source { get; }
        internal long Offset { get; }
        internal long Length { get; }
        internal Integrity Integrity { get; }
    }

    private readonly struct Integrity {
        internal Integrity(uint? crc, byte[]? hash, bool unknownHash) {
            Crc = crc;
            Hash = hash;
            UnknownHash = unknownHash;
        }
        internal uint? Crc { get; }
        internal byte[]? Hash { get; }
        internal bool UnknownHash { get; }
        internal bool IsVerifiable => Crc.HasValue || Hash != null;
    }

    private static Member CreateMember(FileFields fields, FileExtras extras, int ordinal, ArchivePaths paths,
        ArchiveLimits limits, ArchiveBlockBudget budget) {
        if (fields.Name.IndexOf('\\') >= 0 || fields.Name.IndexOf('\ufffe') >= 0)
            throw new InvalidDataException("The RAR5 member name is not portable: " + fields.Name);
        ArchiveEntryKind kind = extras.Redirection?.Kind switch {
            Rar5RedirectionKind.UnixSymbolicLink or Rar5RedirectionKind.WindowsSymbolicLink => ArchiveEntryKind.SymbolicLink,
            Rar5RedirectionKind.HardLink => ArchiveEntryKind.HardLink,
            Rar5RedirectionKind.FileCopy => ArchiveEntryKind.File,
            not null => ArchiveEntryKind.Special,
            _ => GetKind(fields.Host, fields.Attributes, (fields.Flags & 1) != 0)
        };
        limits.CheckEntry(fields.Size);
        ChargePathWork(fields.Name, kind == ArchiveEntryKind.Directory, budget);
        string path = kind == ArchiveEntryKind.Directory ? paths.AddDirectory(fields.Name) : paths.Add(fields.Name);
        return new Member(fields, extras, path, kind, ordinal);
    }

    private static void ChargePathWork(string name, bool directory, ArchiveBlockBudget budget) {
        long characters = directory ? 2L * name.Length : name.Length;
        for (int i = 0; i < name.Length; i++) {
            if (name[i] == '/') characters += i;
        }
        budget.Charge(0, 4 * characters);
    }

    private static ArchiveEntryKind GetKind(ulong host, ulong attributes, bool directory) {
        if (host == 1) {
            ulong mode = attributes & 0xf000;
            if (mode == 0xa000) return ArchiveEntryKind.SymbolicLink;
            if (mode is not (0 or 0x4000 or 0x8000)) return ArchiveEntryKind.Special;
            if (mode != 0 && directory != (mode == 0x4000))
                throw new InvalidDataException("The RAR5 directory flag disagrees with the Unix file type.");
        } else if (host != 0 || (attributes & 0x400) != 0) {
            return ArchiveEntryKind.Special;
        }
        return directory ? ArchiveEntryKind.Directory : ArchiveEntryKind.File;
    }
}
