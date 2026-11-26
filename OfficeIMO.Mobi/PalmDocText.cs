using static OfficeIMO.Mobi.PalmDatabase;

namespace OfficeIMO.Mobi;

internal static class PalmDocText {
    internal static byte[] Read(PalmDatabase database, MobiHeader header, CancellationToken token) {
        if (header.Compression == 17_480) throw new NotSupportedException("Mobipocket HUFF/CDIC compression is outside the supported profile.");
        if (header.Compression is not (1 or 2)) throw new NotSupportedException("This PalmDOC compression method is unsupported.");
        var raw = new byte[header.Length];
        var buffer = new byte[header.RecordSize];
        int written = 0;
        for (int i = 0; i < header.TextRecords; i++) {
            token.ThrowIfCancellationRequested();
            ReadOnlySpan<byte> record = StripTrailers(database.Record(header.Record + i + 1), header.ExtraFlags);
            int count = header.Compression == 2 ? ExpandRecord(record, buffer) : record.Length;
            if (count > buffer.Length || count > raw.Length - written)
                throw new InvalidDataException("A PalmDOC text record exceeds its declared decompressed size.");
            (header.Compression == 2 ? buffer.AsSpan(0, count) : record).CopyTo(raw.AsSpan(written));
            written += count;
        }
        if (written != raw.Length) throw new InvalidDataException("The PalmDOC text length does not match its header.");
        return raw;
    }

    private static ReadOnlySpan<byte> StripTrailers(ReadOnlySpan<byte> source, int flags) {
        for (int bit = 1; bit < 16; bit++) {
            if ((flags & (1 << bit)) == 0) continue;
            int length = 0, width = 0;
            bool end = false;
            while (width < 4 && width < source.Length) {
                byte next = source[source.Length - 1 - width];
                length |= (next & 0x7f) << (7 * width++);
                if ((next & 0x80) == 0) continue;
                end = true;
                break;
            }
            if (!end || length < width || length > source.Length)
                throw new InvalidDataException("A Mobipocket text record has an invalid trailer.");
            source = source.Slice(0, source.Length - length);
        }
        if ((flags & 1) == 0) return source;
        int overlap = (Slice(source, source.Length - 1, 1)[0] & 3) + 1;
        return Slice(source, 0, source.Length - overlap);
    }

    private static int ExpandRecord(ReadOnlySpan<byte> source, Span<byte> output) {
        int read = 0, written = 0;
        while (read < source.Length) {
            byte value = source[read++];
            if (value is >= 1 and <= 8) {
                EnsureSpace(output, written, value);
                Slice(source, read, value).CopyTo(output.Slice(written));
                read += value;
                written += value;
            } else if (value < 0x80) {
                EnsureSpace(output, written, 1);
                output[written++] = value;
            } else if (value >= 0xc0) {
                EnsureSpace(output, written, 2);
                output[written++] = 0x20;
                output[written++] = (byte)(value ^ 0x80);
            } else {
                int pair = U16(source, read - 1);
                read++;
                int distance = (pair >> 3) & 0x7ff;
                int count = (pair & 7) + 3;
                EnsureSpace(output, written, count);
                if (distance == 0 || distance > written)
                    throw new InvalidDataException("A PalmDOC back-reference precedes the current text record.");
                for (int i = 0; i < count; i++) {
                    output[written] = output[written - distance];
                    written++;
                }
            }
        }
        return written;
    }

    private static void EnsureSpace(Span<byte> output, int written, int count) {
        if (count > output.Length - written)
            throw new InvalidDataException("A PalmDOC record expands beyond its declared size.");
    }
}
