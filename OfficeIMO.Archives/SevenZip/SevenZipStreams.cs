using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;

namespace OfficeIMO.Archives.SevenZip;

internal sealed class SevenZipFolder {
    internal const int Copy = 0;
    internal const int Lzma = 0x030101;
    internal const int Lzma2 = 0x21;
    private const int Aes = 0x06f10701;
    private const long MaximumFolderBytes = 256L * 1024 * 1024;

    internal SevenZipFolder(int method, byte[] properties, long offset, long packedSize, uint? packedCrc) {
        Method = method;
        Properties = properties;
        Offset = offset;
        PackedSize = packedSize;
        PackedCrc = packedCrc;
    }

    internal int Method { get; }
    private byte[] Properties { get; }
    internal long Offset { get; }
    internal long PackedSize { get; }
    private uint? PackedCrc { get; }
    internal long Size { get; set; }
    internal uint? Crc { get; set; }
    internal SevenZipExtent[] Extents { get; set; } = Array.Empty<SevenZipExtent>();

    internal ArchiveCompression Compression => Method switch {
        Copy => ArchiveCompression.Store,
        Lzma => ArchiveCompression.Lzma,
        Lzma2 => ArchiveCompression.Lzma2,
        _ => ArchiveCompression.Unknown
    };

    internal void ValidateLayout() {
        switch (Method) {
            case Copy when PackedSize != Size || Properties.Length != 0:
                throw SevenZipInput.Invalid("The 7z Copy coder has inconsistent sizes or properties.");
            case Lzma when PackedSize < 5 || Properties.Length != 5 || Properties[0] >= 225:
                throw SevenZipInput.Invalid("The 7z LZMA coder has an invalid packed size or properties.");
            case Lzma2 when PackedSize == 0 || Properties.Length != 1 || Properties[0] > 40:
                throw SevenZipInput.Invalid("The 7z LZMA2 coder has an invalid packed size or dictionary property.");
        }
    }

    internal byte[] Decode(Stream source, long maximum, string limitName, ArchiveBlockBudget budget, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        ArchiveLimits.Check(Size, Math.Min(maximum, MaximumFolderBytes), limitName);
        if (Compression == ArchiveCompression.Unknown)
            throw new NotSupportedException($"7z compression method 0x{Method:x} is unsupported.");
        long work = Math.Max(8192, Math.Max(PackedSize, Size));
        if (PackedCrc.HasValue) work += Math.Min(PackedSize, long.MaxValue - work);
        budget.Charge(Size, work);
        var output = new byte[(int)Size];
        using var slice = new SeekableReadWindowStream(source, Offset, PackedSize, preserveSourcePosition: false);
        using var packed = new BufferedStream(slice, 8192);
        if (PackedCrc is { } checksum) VerifyPackedChecksum(packed, checksum, token);

        switch (Method) {
            case Copy:
                SevenZipInput.Fill(packed, output, 0, output.Length, token);
                break;
            case Lzma:
                LzmaFolderDecoder.DecodeLzma(packed, Properties, output, budget, token);
                break;
            case Lzma2:
                LzmaFolderDecoder.DecodeLzma2(packed, Properties, output, budget, token);
                break;
        }
        SevenZipInput.VerifyChecksum(output, 0, output.Length, Crc, "folder", token);
        foreach (SevenZipExtent entry in Extents) {
            token.ThrowIfCancellationRequested();
            SevenZipInput.VerifyChecksum(output, (int)entry.Offset, (int)entry.Length, entry.Crc, "member", token);
        }
        return output;
    }

    internal static (int Method, byte[] Properties) ReadCoder(SevenZipInput input) {
        int count = input.Count(16, "MaxSevenZipCodersPerFolder");
        if (count == 0) throw SevenZipInput.Invalid("A 7z folder declares no coders.");
        int method = 0;
        byte[] properties = Array.Empty<byte>();
        bool simple = count == 1;
        for (int i = 0; i < count; i++) {
            int flags = input.Byte();
            int length = flags & 15;
            if (length == 0 || (flags & 0xc0) != 0)
                throw new NotSupportedException("The 7z coder uses unsupported reserved or alternative flags.");
            ulong id = 0;
            for (int j = 0; j < length; j++) id = (id << 8) | input.Byte();
            if (length > 4 || id > int.MaxValue) method = -1;
            else method = (int)id;
            if (method == Aes) throw new NotSupportedException("Encrypted 7z archives are unsupported.");
            if ((flags & 0x10) != 0) {
                long inputs = input.Size();
                long outputs = input.Size();
                simple &= inputs == 1 && outputs == 1;
            }
            properties = (flags & 0x20) == 0 ? Array.Empty<byte>() : input.Bytes(input.Count(64, "MaxSevenZipCoderProperties"));
        }
        if (!simple) throw new NotSupportedException("7z folders with chained filters or multiple packed streams are unsupported.");
        return (method, properties);
    }

    private static void VerifyPackedChecksum(Stream source, uint expected, CancellationToken token) {
        var buffer = new byte[65536];
        uint crc = OfficePngCrc32.Begin();
        int read;
        while ((read = source.Read(buffer, 0, buffer.Length)) != 0) {
            token.ThrowIfCancellationRequested();
            crc = OfficePngCrc32.Append(crc, buffer, 0, read);
        }
        if (OfficePngCrc32.Complete(crc) != expected) throw SevenZipInput.Invalid("The 7z packed stream checksum does not match.");
        source.Position = 0;
    }
}

internal readonly struct SevenZipExtent {
    internal SevenZipExtent(long offset, long length, uint? crc) {
        Offset = offset;
        Length = length;
        Crc = crc;
    }

    internal long Offset { get; }
    internal long Length { get; }
    internal uint? Crc { get; }
}
