using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.TestAssets;

internal static class JpegXrTestFixture {
    // Append a new first directory so encoded pixel packets and their offsets
    // remain unchanged. This creates metadata variants of independent fixtures.
    internal static byte[] WithField(byte[] source, int tag, int type, byte[] value) {
        int directory = Read32(source, 4), count = Read16(source, directory);
        var entries = new List<byte[]>();
        for (int i = 0; i < count; i++) {
            int offset = directory + 2 + i * 12;
            if (Read16(source, offset) != tag) entries.Add(source.Skip(offset).Take(12).ToArray());
        }
        int newDirectory = (source.Length + 1) & ~1;
        int payload = newDirectory + 2 + (entries.Count + 1) * 12 + 4;
        var field = new byte[12]; Write16(field, 0, tag); Write16(field, 2, type);
        Write32(field, 4, type == 3 ? value.Length / 2 : type == 4 || type == 11 ? value.Length / 4 : value.Length);
        if (value.Length <= 4) Buffer.BlockCopy(value, 0, field, 8, value.Length);
        else Write32(field, 8, payload);
        entries.Add(field); entries.Sort((a, b) => Read16(a, 0).CompareTo(Read16(b, 0)));
        var output = new byte[payload + (value.Length > 4 ? value.Length : 0)];
        Buffer.BlockCopy(source, 0, output, 0, source.Length); Write32(output, 4, newDirectory);
        Write16(output, newDirectory, entries.Count);
        for (int i = 0; i < entries.Count; i++) Buffer.BlockCopy(entries[i], 0, output, newDirectory + 2 + i * 12, 12);
        if (value.Length > 4) Buffer.BlockCopy(value, 0, output, payload, value.Length);
        return output;
    }

    private static int Read16(byte[] b, int p) => b[p] | b[p + 1] << 8;
    private static int Read32(byte[] b, int p) => Read16(b, p) | Read16(b, p + 2) << 16;
    private static void Write16(byte[] b, int p, int value) { b[p] = (byte)value; b[p + 1] = (byte)(value >> 8); }
    private static void Write32(byte[] b, int p, int value) { Write16(b, p, value); Write16(b, p + 2, value >> 16); }
}
