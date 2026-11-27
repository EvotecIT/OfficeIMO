namespace OfficeIMO.Archives.Rar5;

internal sealed partial class Rar5Archive {
    private sealed class FileFields {
        internal FileFields(Rar5HeaderReader reader, int nameLimit) {
            Flags = reader.Number();
            Size = reader.Size();
            Attributes = reader.Number();
            if ((Flags & 2) != 0) Time = UnixTime(reader.UInt32(), 0);
            if ((Flags & 4) != 0) Crc = reader.UInt32();
            Coder = reader.Number();
            Host = reader.Number();
            Name = reader.ReadName(nameLimit);
            reader.Finish();
        }
        internal ulong Flags { get; }
        internal long Size { get; }
        internal ulong Attributes { get; }
        internal Rar5Timestamp? Time { get; }
        internal uint? Crc { get; }
        internal ulong Coder { get; }
        internal ulong Host { get; }
        internal string Name { get; }
        internal int Method => (int)((Coder >> 7) & 7);
        internal long Dictionary {
            get {
                if (Method == 0) return 0;
                long size = 131072L << (int)((Coder >> 10) & 31);
                return (Coder & 63) == 1 ? size + (size / 32) * (long)((Coder >> 15) & 31) : size;
            }
        }
    }

    private sealed class FileExtras {
        internal FileExtras(Rar5HeaderReader extras, FileFields fields, long blockOffset, int nameLimit, CancellationToken token) {
            bool sawTime = false;
            bool sawHash = false;
            Times = new Rar5FileTimes(fields.Time, null, null);
            while (extras.Remaining != 0) {
                token.ThrowIfCancellationRequested();
                var item = extras.Slice(extras.Size());
                switch (item.Number()) {
                    case 1:
                        throw new Rar5EncryptionNotSupportedException(blockOffset, fields.Name);
                    case 2:
                        EnsureUnique(sawHash);
                        sawHash = true;
                        if (item.Number() == 0) {
                            Hash = item.Bytes(32);
                            item.Finish();
                        } else UnknownHash = true;
                        break;
                    case 3:
                        EnsureUnique(sawTime);
                        sawTime = true;
                        Times = ReadTimes(item, fields.Time);
                        break;
                    case 5:
                        EnsureUnique(Redirection != null);
                        Redirection = ReadRedirection(item, nameLimit);
                        break;
                    case 6:
                        EnsureUnique(Owner != null);
                        Owner = ReadOwner(item);
                        break;
                }
            }
        }

        internal byte[]? Hash { get; }
        internal bool UnknownHash { get; }
        internal Rar5UnixOwner? Owner { get; }
        internal Rar5Redirection? Redirection { get; }
        internal Rar5FileTimes Times { get; }

        private static void EnsureUnique(bool alreadySeen) {
            if (alreadySeen) throw new InvalidDataException("The RAR5 member has duplicate extra metadata.");
        }
    }

    private static Rar5Redirection ReadRedirection(Rar5HeaderReader reader, int nameLimit) {
        ulong type = reader.Number();
        if (type > int.MaxValue) throw new InvalidDataException("The RAR5 redirection type is out of range.");
        ulong flags = reader.Number();
        string target = reader.ReadName(nameLimit);
        reader.Finish();
        return new Rar5Redirection((Rar5RedirectionKind)type, target, (flags & 1) != 0);
    }

    private static Rar5UnixOwner ReadOwner(Rar5HeaderReader reader) {
        ulong flags = reader.Number();
        byte[]? user = (flags & 1) != 0 ? reader.Bytes(reader.Size()) : null;
        byte[]? group = (flags & 2) != 0 ? reader.Bytes(reader.Size()) : null;
        ulong? userId = (flags & 4) != 0 ? reader.Number() : null;
        ulong? groupId = (flags & 8) != 0 ? reader.Number() : null;
        reader.Finish();
        return new Rar5UnixOwner(user, group, userId, groupId);
    }

    private static Rar5FileTimes ReadTimes(Rar5HeaderReader reader, Rar5Timestamp? baseTime) {
        ulong flags = reader.Number();
        bool unix = (flags & 1) != 0;
        ulong? modified = ReadTime(reader, flags, 2, unix);
        ulong? created = ReadTime(reader, flags, 4, unix);
        ulong? accessed = ReadTime(reader, flags, 8, unix);
        bool nanos = unix && (flags & 16) != 0;
        var times = new Rar5FileTimes(
            DecodeTime(reader, modified, unix, nanos) ?? baseTime,
            DecodeTime(reader, created, unix, nanos),
            DecodeTime(reader, accessed, unix, nanos));
        reader.Finish();
        return times;
    }

    private static ulong? ReadTime(Rar5HeaderReader reader, ulong flags, int bit, bool unix) =>
        (flags & (uint)bit) == 0 ? null : unix ? reader.UInt32() : reader.UInt64();

    private static Rar5Timestamp? DecodeTime(Rar5HeaderReader reader, ulong? value, bool unix, bool nanos) {
        if (value is not ulong time) return null;
        if (unix) return UnixTime((uint)time, nanos ? reader.UInt32() : 0);
        if (time > 2650467743999999999UL) throw new InvalidDataException("The RAR5 FILETIME timestamp is out of range.");
        var utc = new DateTimeOffset(DateTime.FromFileTimeUtc((long)time));
        return new Rar5Timestamp(utc, (int)(time % TimeSpan.TicksPerSecond) * 100);
    }

    private static Rar5Timestamp UnixTime(uint seconds, uint nanos) {
        if (nanos >= 1_000_000_000) throw new InvalidDataException("The RAR5 nanosecond timestamp is out of range.");
        return new Rar5Timestamp(DateTimeOffset.FromUnixTimeSeconds(seconds).AddTicks(nanos / 100), (int)nanos);
    }
}
