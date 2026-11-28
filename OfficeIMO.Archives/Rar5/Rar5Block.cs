using OfficeIMO.Drawing;

namespace OfficeIMO.Archives.Rar5;

/// <summary>One CRC-checked archive header and the bounded range of its separate data area.</summary>
internal sealed class Rar5Block {
    private Rar5Block(ulong type, ulong flags, long packedLength, Rar5HeaderReader fields, Rar5HeaderReader extras) {
        Type = type;
        Flags = flags;
        PackedLength = packedLength;
        Fields = fields;
        Extras = extras;
    }

    internal ulong Type { get; }
    internal ulong Flags { get; }
    internal long PackedLength { get; }
    internal Rar5HeaderReader Fields { get; }
    internal Rar5HeaderReader Extras { get; }

    internal static Rar5Block Read(Stream source, byte[] prefix, ArchiveLimits limits, ArchiveBlockBudget budget,
        ref long metadataBytes, CancellationToken token) {
        Rar5HeaderReader.Fill(source, prefix, 4, token);
        int size = 0;
        int encodedLength = 0;
        bool last = false;
        for (; encodedLength < 3; encodedLength++) {
            int value = source.ReadByte();
            if (value < 0) throw new InvalidDataException("A RAR5 header size is truncated.");
            prefix[4 + encodedLength] = (byte)value;
            size |= (value & 127) << (7 * encodedLength);
            if (value < 128) {
                last = true;
                encodedLength++;
                break;
            }
        }
        if (!last || size < 2 || size > source.Length - source.Position)
            throw new InvalidDataException("A RAR5 header has an invalid size.");
        long headerBytes = 4L + encodedLength + size;
        ArchiveLimits.Add(ref metadataBytes, headerBytes, limits.Metadata, nameof(ArchiveExtractionOptions.MaxMetadataBytes));
        // Account for parsing and allocations even for small service and skippable records.
        budget.Charge(0, Math.Max(headerBytes, 4096));

        var body = new byte[size];
        Rar5HeaderReader.Fill(source, body, body.Length, token);
        uint crc = OfficePngCrc32.Append(OfficePngCrc32.Begin(), prefix, 4, encodedLength);
        if (Rar5HeaderReader.Checksum(crc, body, token) != Rar5HeaderReader.UInt32At(prefix, 0))
            throw new InvalidDataException("The RAR5 header checksum does not match.");

        var header = new Rar5HeaderReader(body);
        ulong type = header.Number();
        ulong flags = header.Number();
        long extraSize = (flags & 1) != 0 ? header.Size() : 0;
        long packedLength = (flags & 2) != 0 ? header.Size() : 0;
        if (packedLength > source.Length - source.Position || extraSize > header.Remaining)
            throw new InvalidDataException("A RAR5 data or extra area extends beyond its bounds.");
        Rar5HeaderReader fields = header.Slice(header.Remaining - extraSize);
        return new Rar5Block(type, flags, packedLength, fields, header.Slice(extraSize));
    }
}
