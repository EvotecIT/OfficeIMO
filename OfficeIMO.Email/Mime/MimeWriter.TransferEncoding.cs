using OfficeIMO.Core.Internal;

namespace OfficeIMO.Email;

internal static partial class MimeWriter {
    private static void WriteQuotedPrintable(Stream output, byte[] data) {
        int lineLength = 0;
        bool endedWithCrlf = false;
        const string hex = "0123456789ABCDEF";
        var token = new byte[3];
        for (int index = 0; index < data.Length; index++) {
            byte value = data[index];
            if (value == '\r' && index + 1 < data.Length && data[index + 1] == '\n') {
                WriteLine(output, string.Empty);
                lineLength = 0;
                endedWithCrlf = true;
                index++;
                continue;
            }
            endedWithCrlf = false;
            int tokenLength;
            if (value >= 33 && value <= 126 && value != '=') {
                token[0] = value;
                tokenLength = 1;
            } else {
                token[0] = (byte)'=';
                token[1] = (byte)hex[value >> 4];
                token[2] = (byte)hex[value & 15];
                tokenLength = 3;
            }
            if (lineLength + tokenLength > 73) {
                WriteLine(output, "=");
                lineLength = 0;
            }
            output.Write(token, 0, tokenLength);
            lineLength += tokenLength;
        }
        if (!endedWithCrlf) WriteLine(output, string.Empty);
    }

    private static byte[] ReadBoundedBytes(Stream input, long maximumInputBytes) {
        using (var buffer = new MemoryStream()) {
            var chunk = new byte[81920];
            long total = 0;
            while (true) {
                int read = input.Read(chunk, 0, chunk.Length);
                if (read == 0) break;
                total = checked(total + read);
                if (total > maximumInputBytes) {
                    throw new EmailLimitExceededException(nameof(EmailWriterOptions.MaxOutputBytes), total, maximumInputBytes);
                }
                buffer.Write(chunk, 0, read);
            }
            return buffer.ToArray();
        }
    }

    private static void WriteBase64(Stream output, byte[] data, int lineLength) {
        if (data.Length == 0) {
            WriteLine(output, string.Empty);
            return;
        }
        const int maximumBufferedInputBytes = 1024 * 1024;
        if (data.Length <= maximumBufferedInputBytes) {
            int characters = checked((data.Length + 2) / 3 * 4);
            int lines = checked((characters + lineLength - 1) / lineLength);
            int length = checked(characters + lines * 2);
#if NET8_0_OR_GREATER
            byte[] encoded = System.Buffers.ArrayPool<byte>.Shared.Rent(length);
#else
            byte[] encoded = new byte[length];
#endif
            try {
                OfficeBase64.EncodeToUtf8(data, 0, data.Length, encoded, 0);
                // Move backwards so the complete runtime SIMD encode can share storage with line folding.
                int sourceEnd = characters;
                int targetEnd = length;
                int take = characters - (lines - 1) * lineLength;
                while (sourceEnd > 0) {
                    encoded[--targetEnd] = (byte)'\n';
                    encoded[--targetEnd] = (byte)'\r';
                    sourceEnd -= take;
                    targetEnd -= take;
#if NET8_0_OR_GREATER
                    encoded.AsSpan(sourceEnd, take).CopyTo(encoded.AsSpan(targetEnd, take));
#else
                    Buffer.BlockCopy(encoded, sourceEnd, encoded, targetEnd, take);
#endif
                    take = lineLength;
                }
                output.Write(encoded, 0, length);
            } finally {
#if NET8_0_OR_GREATER
                System.Buffers.ArrayPool<byte>.Shared.Return(encoded);
#endif
            }
            return;
        }
        using (var input = new MemoryStream(data, writable: false)) {
            WriteBase64(output, input, lineLength, data.LongLength);
        }
    }

    internal static void WriteBase64(Stream output, Stream input, int lineLength, long maximumInputBytes,
        CancellationToken cancellationToken = default) {
        int bytesPerLine = checked(lineLength / 4 * 3);
        var buffer = new byte[bytesPerLine];
        var encodedLine = new byte[checked(lineLength + 2)];
        long total = 0;
        bool wrote = false;
        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = 0;
            while (count < buffer.Length) {
                int read = input.Read(buffer, count, buffer.Length - count);
                if (read == 0) break;
                count += read;
                total = checked(total + read);
                if (total > maximumInputBytes) {
                    throw new EmailLimitExceededException(nameof(EmailWriterOptions.MaxOutputBytes),
                        total, maximumInputBytes);
                }
            }
            if (count == 0) break;
            int characterCount = OfficeBase64.EncodeToUtf8(buffer, 0, count, encodedLine, 0);
            encodedLine[characterCount] = (byte)'\r';
            encodedLine[characterCount + 1] = (byte)'\n';
            output.Write(encodedLine, 0, characterCount + 2);
            wrote = true;
            if (count < buffer.Length) break;
        }
        if (!wrote) {
            WriteLine(output, string.Empty);
        }
    }

}
