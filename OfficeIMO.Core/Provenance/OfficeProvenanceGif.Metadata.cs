using System;
using System.IO;
using System.Text;
using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.Provenance;

internal static partial class OfficeProvenanceGif {
    internal static byte[] RewriteMetadataProfiles(byte[] input, byte[]? xmp, byte[]? icc,
        OfficeImageMetadataProfileKinds replace, bool writeReplacement, CancellationToken token,
        out OfficeImageMetadataProfileKinds present, out byte[]? readXmp, out byte[]? readIcc, bool readOnly = false, long additionallyRetainedBytes = 0L) {
        const int maximumEntries = 100000;
        present = OfficeImageMetadataProfileKinds.None;
        readXmp = null;
        readIcc = null;
        int cursor = GetBodyOffset(input);
        int entryCount = 0;
        int nextRawTrailer = -2;
        long profiles = (xmp?.LongLength ?? 0L) + (icc?.LongLength ?? 0L);
        long retained = checked(input.LongLength + additionallyRetainedBytes + profiles * 3L + 65536L);
        int capacity = checked((int)Math.Min(OfficeRasterGuards.MaximumEncodedBytes, input.LongLength + profiles + profiles / 255L + 512L));
        using var output = readOnly ? null : new OfficeMetadataRewriteStream(retained, capacity, token);
        output?.Write(input, 0, cursor);
        if (writeReplacement) {
            if (xmp != null || icc != null) {
                // Application metadata uses the GIF89a extension grammar.
                output!.Position = 0;
                output.Write(Encoding.ASCII.GetBytes("GIF89a"), 0, 6);
                output.Position = cursor;
            }
            if (icc != null) WriteProfileApplication(output!, "ICCRGBG1012", icc, xmpTrailer: false);
            if (xmp != null) WriteProfileApplication(output!, "XMP DataXMP", xmp, xmpTrailer: true);
        }
        while (cursor < input.Length) {
            token.ThrowIfCancellationRequested();
            ReserveEntry(ref entryCount, maximumEntries);
            int start = cursor;
            byte introducer = input[cursor++];
            if (introducer == 0x3B) {
                output?.Write(input, start, input.Length - start);
                return output?.ToArray() ?? Array.Empty<byte>();
            }
            if (introducer == 0x2C) {
                if (input.Length - cursor < 9) throw new FormatException("Truncated GIF image descriptor.");
                byte flags = input[cursor + 8];
                cursor += 9;
                if ((flags & 128) != 0) cursor += 3 << ((flags & 7) + 1);
                if (cursor >= input.Length) throw new FormatException("Truncated GIF image palette or code size.");
                cursor++;
                cursor = SkipSubBlocks(input, cursor, OfficeRasterGuards.MaximumEncodedBytes, ref entryCount, maximumEntries, out _, token);
                output?.Write(input, start, cursor - start);
                continue;
            }
            if (introducer != 0x21 || cursor >= input.Length) throw new FormatException("Malformed GIF extension.");
            byte label = input[cursor++];
            OfficeImageMetadataProfileKinds kind = OfficeImageMetadataProfileKinds.None;
            if (label == 0xFF) {
                if (cursor >= input.Length) throw new FormatException("Truncated GIF application header.");
                int headerLength = input[cursor++];
                if (headerLength > input.Length - cursor) throw new FormatException("Truncated GIF application header.");
                bool isXmp = headerLength == 11 && OfficeProvenanceBinary.MatchesAscii(input, cursor, "XMP DataXMP");
                bool isIcc = headerLength == 11 && OfficeProvenanceBinary.MatchesAscii(input, cursor, "ICCRGBG1012");
                cursor += headerLength;
                int payload = cursor;
                if (isXmp) {
                    output?.CheckTransientBytes(2L * OfficeExifProfileCodec.MaximumProfileBytes);
                    if (!TryReadXmpApplicationData(input, payload, OfficeExifProfileCodec.MaximumProfileBytes,
                        ref entryCount, maximumEntries, token, ref nextRawTrailer, out byte[] packet,
                        out int extensionEnd, out _, out _)) throw new FormatException("Malformed GIF XMP application data.");
                    if (readXmp == null) { output?.AddRetainedBytes(packet.LongLength); readXmp = packet; }
                    cursor = extensionEnd;
                    kind = OfficeImageMetadataProfileKinds.Xmp;
                } else {
                    cursor = SkipSubBlocks(input, payload, isIcc ? OfficeExifProfileCodec.MaximumProfileBytes : OfficeRasterGuards.MaximumEncodedBytes,
                        ref entryCount, maximumEntries, out int length, token);
                    if (isIcc) {
                        output?.CheckTransientBytes(length);
                        if (readIcc == null) { output?.AddRetainedBytes(length); readIcc = CollectSubBlocks(input, payload, length, token); }
                        kind = OfficeImageMetadataProfileKinds.Icc;
                    }
                }
            } else {
                cursor = SkipSubBlocks(input, cursor, OfficeRasterGuards.MaximumEncodedBytes, ref entryCount, maximumEntries, out _, token);
            }
            present |= kind;
            if (kind == OfficeImageMetadataProfileKinds.None || (replace & kind) == 0) output?.Write(input, start, cursor - start);
        }
        throw new FormatException("GIF is missing its trailer.");
    }

    private static void WriteProfileApplication(Stream output, string identifier, byte[] profile, bool xmpTrailer) {
        output.WriteByte(0x21);
        output.WriteByte(0xFF);
        output.WriteByte(11);
        output.Write(Encoding.ASCII.GetBytes(identifier), 0, 11);
        WriteSubBlocks(output, profile);
        if (xmpTrailer) {
            output.WriteByte(1);
            for (int value = 255; value >= 0; value--) output.WriteByte((byte)value);
        }
        output.WriteByte(0);
    }
}
