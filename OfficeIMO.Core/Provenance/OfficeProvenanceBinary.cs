using System;
using System.IO;
using System.Text;

namespace OfficeIMO.Provenance;

internal static class OfficeProvenanceBinary {
    internal static byte[] ReadBounded(Stream stream, long maximumBytes) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (maximumBytes <= 0 || maximumBytes > int.MaxValue) {
            throw new ArgumentOutOfRangeException(nameof(maximumBytes), "The asset limit must be between 1 byte and Int32.MaxValue.");
        }

        if (stream.CanSeek) {
            long remaining = stream.Length - stream.Position;
            if (remaining < 0 || remaining > maximumBytes) {
                throw new InvalidDataException($"The asset exceeds the configured limit of {maximumBytes} bytes.");
            }
            byte[] data = new byte[(int)remaining];
            ReadExactly(stream, data, 0, data.Length);
            return data;
        }

        using var buffer = new MemoryStream();
        byte[] chunk = new byte[8192];
        while (true) {
            int read = stream.Read(chunk, 0, chunk.Length);
            if (read <= 0) break;
            if (buffer.Length > maximumBytes - read) {
                throw new InvalidDataException($"The asset exceeds the configured limit of {maximumBytes} bytes.");
            }
            buffer.Write(chunk, 0, read);
        }
        return buffer.ToArray();
    }

    internal static void ValidateLimits(OfficeProvenanceOptions options) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (options.MaxAssetBytes <= 0 || options.MaxAssetBytes > int.MaxValue) {
            throw new ArgumentOutOfRangeException(nameof(options.MaxAssetBytes));
        }
        if (options.MaxManifestBytes <= 0 || options.MaxManifestBytes > options.MaxAssetBytes) {
            throw new ArgumentOutOfRangeException(nameof(options.MaxManifestBytes));
        }
        if (options.MaxCarriers <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxCarriers));
        if (options.MaxContainerEntries <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxContainerEntries));
        if (options.MaxExpandedContainerBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxExpandedContainerBytes));
        if (options.MaxEmbeddedAssets <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxEmbeddedAssets));
    }

    internal static void ValidateRemovalOptions(OfficeProvenanceRemovalOptions options) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        ValidateLimits(options.Limits);
        if (options.MaxOutputBytes.HasValue &&
            (options.MaxOutputBytes.Value <= 0 || options.MaxOutputBytes.Value > int.MaxValue)) {
            throw new ArgumentOutOfRangeException(nameof(options.MaxOutputBytes));
        }
        if (options.MaxEmbeddedAssets <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxEmbeddedAssets));
        if (!Enum.IsDefined(typeof(OfficeIMO.OfficeSignatureMutationPolicy), options.SignatureMutationPolicy)) {
            throw new ArgumentOutOfRangeException(nameof(options.SignatureMutationPolicy));
        }
    }

    internal static void EnsureOutputWithinLimit(long outputBytes, long maximumOutputBytes) {
        if (maximumOutputBytes <= 0 || maximumOutputBytes > int.MaxValue) {
            throw new ArgumentOutOfRangeException(nameof(maximumOutputBytes));
        }
        if (outputBytes < 0 || outputBytes > maximumOutputBytes) {
            throw OfficeProvenanceLimitException.CreateOutput(
                $"The rewritten asset exceeds the configured output limit of {maximumOutputBytes} bytes.");
        }
    }

    internal static byte[] CloneForOutput(byte[] data, long maximumOutputBytes) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        EnsureOutputWithinLimit(data.LongLength, maximumOutputBytes);
        return (byte[])data.Clone();
    }

    internal static bool HasPrefix(byte[] data, params byte[] prefix) {
        if (data.Length < prefix.Length) return false;
        for (int index = 0; index < prefix.Length; index++) {
            if (data[index] != prefix[index]) return false;
        }
        return true;
    }

    internal static bool MatchesAscii(byte[] data, int offset, string expected) {
        if (offset < 0 || expected.Length > data.Length - offset) return false;
        for (int index = 0; index < expected.Length; index++) {
            if (data[offset + index] != (byte)expected[index]) return false;
        }
        return true;
    }

    internal static ushort ReadUInt16(byte[] data, int offset, bool littleEndian) {
        EnsureRange(data, offset, 2);
        return littleEndian
            ? (ushort)(data[offset] | (data[offset + 1] << 8))
            : (ushort)((data[offset] << 8) | data[offset + 1]);
    }

    internal static uint ReadUInt32(byte[] data, int offset, bool littleEndian) {
        EnsureRange(data, offset, 4);
        if (littleEndian) {
            return (uint)(data[offset] | (data[offset + 1] << 8) | (data[offset + 2] << 16) | (data[offset + 3] << 24));
        }
        return ((uint)data[offset] << 24) | ((uint)data[offset + 1] << 16) | ((uint)data[offset + 2] << 8) | data[offset + 3];
    }

    internal static ulong ReadUInt64(byte[] data, int offset, bool littleEndian) {
        EnsureRange(data, offset, 8);
        ulong result = 0;
        if (littleEndian) {
            for (int index = 7; index >= 0; index--) result = (result << 8) | data[offset + index];
        } else {
            for (int index = 0; index < 8; index++) result = (result << 8) | data[offset + index];
        }
        return result;
    }

    internal static void WriteUInt16(byte[] data, int offset, ushort value, bool littleEndian) {
        EnsureRange(data, offset, 2);
        if (littleEndian) {
            data[offset] = (byte)value;
            data[offset + 1] = (byte)(value >> 8);
        } else {
            data[offset] = (byte)(value >> 8);
            data[offset + 1] = (byte)value;
        }
    }

    internal static void WriteUInt32(byte[] data, int offset, uint value, bool littleEndian) {
        EnsureRange(data, offset, 4);
        if (littleEndian) {
            data[offset] = (byte)value;
            data[offset + 1] = (byte)(value >> 8);
            data[offset + 2] = (byte)(value >> 16);
            data[offset + 3] = (byte)(value >> 24);
        } else {
            data[offset] = (byte)(value >> 24);
            data[offset + 1] = (byte)(value >> 16);
            data[offset + 2] = (byte)(value >> 8);
            data[offset + 3] = (byte)value;
        }
    }

    internal static void WriteUInt64(byte[] data, int offset, ulong value, bool littleEndian) {
        EnsureRange(data, offset, 8);
        if (littleEndian) {
            for (int index = 0; index < 8; index++) data[offset + index] = (byte)(value >> (index * 8));
        } else {
            for (int index = 0; index < 8; index++) data[offset + index] = (byte)(value >> ((7 - index) * 8));
        }
    }

    internal static string DecodeUtf8(byte[] data) {
        var encoding = new UTF8Encoding(false, true);
        return encoding.GetString(data);
    }

    internal static string DecodeUtf8(byte[] data, int offset, int count) {
        EnsureRange(data, offset, count);
        var encoding = new UTF8Encoding(false, true);
        return encoding.GetString(data, offset, count);
    }

    internal static void ReadExactly(Stream stream, byte[] buffer, int offset, int count) {
        while (count > 0) {
            int read = stream.Read(buffer, offset, count);
            if (read <= 0) throw new EndOfStreamException("Unexpected end of provenance data.");
            offset += read;
            count -= read;
        }
    }

    internal static void EnsureRange(byte[] data, int offset, int count) {
        if (offset < 0 || count < 0 || offset > data.Length - count) {
            throw new InvalidDataException("A provenance carrier points outside the asset bounds.");
        }
    }
}
