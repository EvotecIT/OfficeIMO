namespace OfficeIMO.Email.Store;

internal enum PstVariant {
    Ansi,
    Unicode,
    Unicode4K
}

internal sealed class PstHeader {
    private PstHeader() { }

    internal PstVariant Variant { get; private set; }
    internal int Version { get; private set; }
    internal int PageSize { get; private set; }
    internal int PageTrailerSize { get; private set; }
    internal int BlockTrailerSize { get; private set; }
    internal int BlockAlignment { get; private set; }
    internal int BTreeMetadataSize { get; private set; }
    internal ulong NbtRootBid { get; private set; }
    internal long NbtRootOffset { get; private set; }
    internal ulong BbtRootBid { get; private set; }
    internal long BbtRootOffset { get; private set; }
    internal byte CryptMethod { get; private set; }
    internal bool IsUnicode => Variant != PstVariant.Ansi;
    internal int BidSize => IsUnicode ? 8 : 4;

    internal static PstHeader Read(Stream stream, EmailStoreFormat expectedFormat) {
        byte[] bytes = PstBinary.ReadAt(stream, 0, checked((int)Math.Min(stream.Length, 564)));
        if (bytes.Length < 512 || PstBinary.UInt32(bytes, 0) != 0x4E444221) {
            throw new InvalidDataException("The PST/OST header signature is invalid.");
        }
        bool isOst = bytes[8] == 0x53 && bytes[9] == 0x4F;
        bool isPst = bytes[8] == 0x53 && bytes[9] == 0x4D;
        if (!isOst && !isPst) throw new InvalidDataException("The NDB client signature is not PST or OST.");
        if ((expectedFormat == EmailStoreFormat.Ost) != isOst) {
            throw new InvalidDataException("The detected store kind does not match the NDB client signature.");
        }

        int version = PstBinary.UInt16(bytes, 10);
        var header = new PstHeader { Version = version };
        if (version == 14 || version == 15) {
            header.Variant = PstVariant.Ansi;
            header.PageSize = 512;
            header.PageTrailerSize = 12;
            header.BlockTrailerSize = 12;
            header.BlockAlignment = 64;
            header.BTreeMetadataSize = 4;
            header.NbtRootBid = PstBinary.UInt32(bytes, 184);
            header.NbtRootOffset = PstBinary.UInt32(bytes, 188);
            header.BbtRootBid = PstBinary.UInt32(bytes, 192);
            header.BbtRootOffset = PstBinary.UInt32(bytes, 196);
            header.CryptMethod = bytes[461];
        } else if (version == 21 || version == 23) {
            header.Variant = PstVariant.Unicode;
            header.PageSize = 512;
            header.PageTrailerSize = 16;
            header.BlockTrailerSize = 16;
            header.BlockAlignment = 64;
            header.BTreeMetadataSize = 8;
            header.NbtRootBid = PstBinary.UInt64(bytes, 216);
            header.NbtRootOffset = checked((long)PstBinary.UInt64(bytes, 224));
            header.BbtRootBid = PstBinary.UInt64(bytes, 232);
            header.BbtRootOffset = checked((long)PstBinary.UInt64(bytes, 240));
            header.CryptMethod = bytes[513];
        } else if (version == 36) {
            header.Variant = PstVariant.Unicode4K;
            header.PageSize = 4096;
            header.PageTrailerSize = 24;
            header.BlockTrailerSize = 24;
            header.BlockAlignment = 512;
            header.BTreeMetadataSize = 16;
            header.NbtRootBid = PstBinary.UInt64(bytes, 216);
            header.NbtRootOffset = checked((long)PstBinary.UInt64(bytes, 224));
            header.BbtRootBid = PstBinary.UInt64(bytes, 232);
            header.BbtRootOffset = checked((long)PstBinary.UInt64(bytes, 240));
            header.CryptMethod = bytes[513];
        } else {
            throw new NotSupportedException(string.Concat("Unsupported PST/OST NDB version ",
                version.ToString(CultureInfo.InvariantCulture), "."));
        }

        if (header.NbtRootOffset <= 0 || header.BbtRootOffset <= 0 ||
            header.NbtRootOffset > stream.Length - header.PageSize ||
            header.BbtRootOffset > stream.Length - header.PageSize) {
            throw new InvalidDataException("The PST/OST root B-tree references are invalid.");
        }
        if (header.CryptMethod > 2) {
            throw new NotSupportedException(string.Concat("Unsupported PST/OST encryption method ",
                header.CryptMethod.ToString(CultureInfo.InvariantCulture), "."));
        }
        return header;
    }
}
