using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Reads bounded family-name metadata without loading outline or glyph tables.</summary>
internal static class OfficeFontNameMetadata {
    internal static bool MatchesFamily(byte[] data, string familyKey, CancellationToken cancellationToken) {
        using var stream = new MemoryStream(data, writable: false);
        return ReadFamilyKeys(stream, cancellationToken).Contains(familyKey);
    }

    internal static string FamilyKey(string value) {
        var key = new StringBuilder(value.Length);
        foreach (char c in value) if (char.IsLetterOrDigit(c)) key.Append(char.ToLowerInvariant(c));
        return key.ToString();
    }

    internal static HashSet<string> ReadFamilyKeys(string path, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete);
            return ReadFamilyKeys(stream, cancellationToken);
        } catch (Exception exception) when (exception is IOException || exception is UnauthorizedAccessException
            || exception is ArgumentException || exception is NotSupportedException) {
            return new HashSet<string>(StringComparer.Ordinal);
        }
    }

    internal static HashSet<string> ReadFamilyKeys(Stream stream, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var names = new HashSet<string>(StringComparer.Ordinal);
        try {
            byte[] header = Read(stream, 0, 12, cancellationToken);
            var faces = new List<uint>();
            if (UInt32(header, 0) == 0x74746366) {
                uint count = UInt32(header, 8);
                if (count == 0 || count > 64) return names;
                byte[] offsets = Read(stream, 12, (int)count * 4, cancellationToken);
                for (int index = 0; index < count; index++) faces.Add(UInt32(offsets, index * 4));
            } else faces.Add(0);
            int totalNameBytes = 0;
            foreach (uint face in faces) {
                cancellationToken.ThrowIfCancellationRequested();
                header = Read(stream, face, 12, cancellationToken);
                uint scaler = UInt32(header, 0);
                int count = UInt16(header, 4);
                if (scaler != 0x00010000 && scaler != 0x74727565 && scaler != 0x4F54544F
                    || count == 0 || count > 4096) continue;
                byte[] directory = Read(stream, (long)face + 12, count * 16, cancellationToken);
                for (int index = 0; index < count; index++) {
                    int record = index * 16;
                    if (UInt32(directory, record) != 0x6E616D65) continue;
                    uint offset = UInt32(directory, record + 8), length = UInt32(directory, record + 12);
                    if (length < 6 || length > 256 * 1024 || totalNameBytes > 4 * 1024 * 1024 - length) break;
                    totalNameBytes += (int)length;
                    AddNames(Read(stream, offset, (int)length, cancellationToken), names, cancellationToken);
                    break;
                }
            }
        } catch (Exception exception) when (exception is IOException || exception is InvalidDataException) {
            // A corrupt face or an updating cache is unavailable for discovery.
        }
        return names;
    }

    private static void AddNames(byte[] table, HashSet<string> names, CancellationToken cancellationToken) {
        int count = UInt16(table, 2), storage = UInt16(table, 4);
        if (count > 4096 || 6 + count * 12 > table.Length || storage < 6 + count * 12 || storage > table.Length) return;
        int decodedBytes = 0;
        for (int index = 0; index < count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            int record = 6 + index * 12, id = UInt16(table, record + 6);
            if (id != 1 && id != 16) continue;
            int platform = UInt16(table, record), length = UInt16(table, record + 8), offset = UInt16(table, record + 10);
            if (length == 0 || length > 512 || storage + offset > table.Length - length) continue;
            if (decodedBytes > 64 * 1024 - length) return;
            decodedBytes += length;
            string value;
            if (platform == 0 || platform == 3) {
                if ((length & 1) != 0) continue;
                value = Encoding.BigEndianUnicode.GetString(table, storage + offset, length);
            } else if (platform == 1) value = Encoding.ASCII.GetString(table, storage + offset, length);
            else continue;
            string key = FamilyKey(value.Trim('\0', ' '));
            if (key.Length > 0) names.Add(key);
            if (names.Count >= 128) return;
        }
    }

    private static byte[] Read(Stream stream, long offset, int length, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (offset < 0 || offset > stream.Length - length) throw new InvalidDataException("Font name metadata is truncated.");
        stream.Position = offset;
        var buffer = new byte[length];
        for (int read = 0; read < length;) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = stream.Read(buffer, read, Math.Min(65536, length - read));
            cancellationToken.ThrowIfCancellationRequested();
            if (count == 0) throw new EndOfStreamException();
            read += count;
        }
        return buffer;
    }

    private static int UInt16(byte[] data, int offset) => data[offset] * 256 + data[offset + 1];
    private static uint UInt32(byte[] data, int offset) => (uint)data[offset] << 24
        | (uint)data[offset + 1] << 16 | (uint)data[offset + 2] << 8 | data[offset + 3];
}
