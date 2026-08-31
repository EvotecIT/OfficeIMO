using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;

namespace OfficeIMO;

/// <summary>Shared bounded byte and text primitives used by legacy format adapters.</summary>
internal static class OfficeLegacyImportBuffer {
    internal static byte[] ReadAll(string path, OfficeLegacyImportLimits limits, CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        using FileStream stream = File.OpenRead(path);
        return ReadAll(stream, limits, cancellationToken);
    }

    internal static byte[] ReadAll(Stream stream, OfficeLegacyImportLimits limits, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Source stream must be readable.", nameof(stream));
        if (limits == null) throw new ArgumentNullException(nameof(limits));
        limits.Validate();

        if (stream.CanSeek && stream.Length - stream.Position > limits.MaxInputBytes) {
            throw new InvalidDataException($"Legacy input exceeds the configured {limits.MaxInputBytes} byte limit.");
        }

        using var output = new MemoryStream(Math.Min(limits.MaxInputBytes, 81920));
        var buffer = new byte[81920];
        int total = 0;
        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            int remaining = limits.MaxInputBytes - total;
            int requested = remaining >= buffer.Length ? buffer.Length : remaining + 1;
            int read = stream.Read(buffer, 0, requested);
            if (read == 0) break;
            if (read > remaining) {
                throw new InvalidDataException($"Legacy input exceeds the configured {limits.MaxInputBytes} byte limit.");
            }
            total += read;
            output.Write(buffer, 0, read);
        }
        return output.ToArray();
    }

    internal static bool StartsWith(byte[] data, params byte[] prefix) {
        if (data == null || prefix == null || data.Length < prefix.Length) return false;
        for (int index = 0; index < prefix.Length; index++) {
            if (data[index] != prefix[index]) return false;
        }
        return true;
    }

    internal static ushort ReadUInt16(byte[] data, int offset) {
        if (offset < 0 || offset + 2 > data.Length) throw new InvalidDataException("Truncated legacy 16-bit value.");
        return (ushort)(data[offset] | (data[offset + 1] << 8));
    }

    internal static int ReadInt16(byte[] data, int offset) => unchecked((short)ReadUInt16(data, offset));

    internal static int ReadInt32(byte[] data, int offset) {
        if (offset < 0 || offset + 4 > data.Length) throw new InvalidDataException("Truncated legacy 32-bit value.");
        return data[offset] | (data[offset + 1] << 8) | (data[offset + 2] << 16) | (data[offset + 3] << 24);
    }

    internal static string ExtractPrintableText(
        byte[] data,
        int offset,
        int length,
        int maxCharacters,
        bool stripHighBit = false,
        int minimumRunLength = 3,
        CancellationToken cancellationToken = default) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        if (offset < 0 || length < 0 || offset > data.Length - length) throw new ArgumentOutOfRangeException(nameof(offset));
        if (maxCharacters < 1) throw new ArgumentOutOfRangeException(nameof(maxCharacters));

        var output = new StringBuilder(Math.Min(length, maxCharacters));
        var run = new StringBuilder();
        int end = offset + length;
        for (int index = offset; index < end && output.Length + run.Length < maxCharacters; index++) {
            if ((index & 0xFFF) == 0) cancellationToken.ThrowIfCancellationRequested();
            byte source = data[index];
            byte value = stripHighBit ? (byte)(source & 0x7F) : source;
            if (value == 9 || value == 10 || value == 13 || value >= 32 && value <= 126) {
                if (value == 13) {
                    FlushRun(output, run, minimumRunLength, maxCharacters);
                    AppendBounded(output, "\n", maxCharacters);
                    if (index + 1 < end && (data[index + 1] & (stripHighBit ? 0x7F : 0xFF)) == 10) index++;
                } else if (value == 10) {
                    FlushRun(output, run, minimumRunLength, maxCharacters);
                    AppendBounded(output, "\n", maxCharacters);
                } else {
                    run.Append(value == 9 ? '\t' : (char)value);
                }
            } else {
                FlushRun(output, run, minimumRunLength, maxCharacters);
                if (output.Length > 0 && output[output.Length - 1] != '\n') AppendBounded(output, " ", maxCharacters);
            }
        }
        FlushRun(output, run, minimumRunLength, maxCharacters);
        return NormalizeWhitespace(output.ToString());
    }

    private static void FlushRun(StringBuilder output, StringBuilder run, int minimumRunLength, int maximum) {
        if (run.Length >= minimumRunLength) AppendBounded(output, run.ToString(), maximum);
        run.Clear();
    }

    private static void AppendBounded(StringBuilder output, string value, int maximum) {
        int available = maximum - output.Length;
        if (available <= 0) return;
        output.Append(value, 0, Math.Min(value.Length, available));
    }

    private static string NormalizeWhitespace(string value) {
        string normalized = value.Replace("\r\n", "\n").Replace('\r', '\n');
        while (normalized.IndexOf("\n\n\n", StringComparison.Ordinal) >= 0) normalized = normalized.Replace("\n\n\n", "\n\n");
        return normalized.Trim(' ', '\t', '\n');
    }
}
