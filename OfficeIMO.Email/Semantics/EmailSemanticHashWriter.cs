using System.Collections;
using System.Security.Cryptography;

namespace OfficeIMO.Email;

internal sealed class EmailSemanticHashWriter : IDisposable {
    private readonly HashAlgorithm _hash;
    private bool _completed;

    internal EmailSemanticHashWriter(byte[]? key) {
        _hash = key == null ? SHA256.Create() : new HMACSHA256(key);
    }

    internal void WriteByte(byte value) {
        byte[] bytes = { value };
        WriteRaw(bytes, 0, bytes.Length);
    }

    internal void WriteInt32(int value) {
        byte[] bytes = new byte[4];
        bytes[0] = unchecked((byte)value);
        bytes[1] = unchecked((byte)(value >> 8));
        bytes[2] = unchecked((byte)(value >> 16));
        bytes[3] = unchecked((byte)(value >> 24));
        WriteRaw(bytes, 0, bytes.Length);
    }

    internal void WriteInt64(long value) {
        byte[] bytes = new byte[8];
        for (int index = 0; index < bytes.Length; index++) {
            bytes[index] = unchecked((byte)(value >> (index * 8)));
        }
        WriteRaw(bytes, 0, bytes.Length);
    }

    internal void WriteString(string? value) {
        if (value == null) {
            WriteInt32(-1);
            return;
        }
        byte[] bytes = Encoding.UTF8.GetBytes(value);
        WriteInt32(bytes.Length);
        WriteRaw(bytes, 0, bytes.Length);
    }

    internal void WriteBytes(byte[]? value) {
        if (value == null) {
            WriteInt64(-1);
            return;
        }
        WriteInt64(value.LongLength);
        WriteRaw(value, 0, value.Length);
    }

    internal void WriteRaw(byte[] buffer, int offset, int count) {
        if (_completed) throw new InvalidOperationException("The semantic digest has already completed.");
        if (count == 0) return;
        _hash.TransformBlock(buffer, offset, count, buffer, offset);
    }

    internal byte[] Complete() {
        if (!_completed) {
            _hash.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
            _completed = true;
        }
        return _hash.Hash == null ? Array.Empty<byte>() : (byte[])_hash.Hash.Clone();
    }

    public void Dispose() => _hash.Dispose();
}

internal static class EmailSemanticValueDigest {
    internal static byte[] Compute(object? value, byte[]? key) {
        using (var writer = new EmailSemanticHashWriter(key)) {
            WriteValue(writer, value, 0);
            return writer.Complete();
        }
    }

    internal static byte[] ComputeStream(Stream stream, byte[]? key,
        long maximumBytes, CancellationToken cancellationToken, out long bytesRead) {
        using (var writer = new EmailSemanticHashWriter(key)) {
            writer.WriteString("stream-v1");
            bytesRead = 0;
            byte[] buffer = new byte[81920];
            while (true) {
                cancellationToken.ThrowIfCancellationRequested();
                int read = stream.Read(buffer, 0, buffer.Length);
                if (read == 0) break;
                bytesRead = checked(bytesRead + read);
                if (bytesRead > maximumBytes) {
                    throw new EmailLimitExceededException(
                        nameof(EmailSemanticComparisonOptions.MaxAttachmentBytes), bytesRead, maximumBytes);
                }
                writer.WriteRaw(buffer, 0, read);
            }
            writer.WriteInt64(bytesRead);
            return writer.Complete();
        }
    }

    internal static async Task<EmailSemanticStreamDigest> ComputeStreamAsync(Stream stream,
        byte[]? key, long maximumBytes, CancellationToken cancellationToken) {
        using (var writer = new EmailSemanticHashWriter(key)) {
            writer.WriteString("stream-v1");
            long bytesRead = 0;
            byte[] buffer = new byte[81920];
            while (true) {
                cancellationToken.ThrowIfCancellationRequested();
                int read = await stream.ReadAsync(buffer, 0, buffer.Length, cancellationToken)
                    .ConfigureAwait(false);
                if (read == 0) break;
                bytesRead = checked(bytesRead + read);
                if (bytesRead > maximumBytes) {
                    throw new EmailLimitExceededException(
                        nameof(EmailSemanticComparisonOptions.MaxAttachmentBytes), bytesRead, maximumBytes);
                }
                writer.WriteRaw(buffer, 0, read);
            }
            writer.WriteInt64(bytesRead);
            return new EmailSemanticStreamDigest(writer.Complete(), bytesRead);
        }
    }

    private static void WriteValue(EmailSemanticHashWriter writer, object? value, int depth) {
        if (depth > 64) throw new InvalidDataException("A semantic value exceeds the supported nesting depth.");
        if (value == null) { writer.WriteByte(0); return; }
        if (value is string text) { writer.WriteByte(1); writer.WriteString(text); return; }
        if (value is byte[] bytes) { writer.WriteByte(2); writer.WriteBytes(bytes); return; }
        if (value is bool boolean) { writer.WriteByte(3); writer.WriteByte(boolean ? (byte)1 : (byte)0); return; }
        if (value is short int16) { writer.WriteByte(4); writer.WriteInt32(int16); return; }
        if (value is ushort uint16) { writer.WriteByte(5); writer.WriteInt32(uint16); return; }
        if (value is int int32) { writer.WriteByte(6); writer.WriteInt32(int32); return; }
        if (value is uint uint32) { writer.WriteByte(7); writer.WriteInt64(uint32); return; }
        if (value is long int64) { writer.WriteByte(8); writer.WriteInt64(int64); return; }
        if (value is ulong uint64) { writer.WriteByte(9); writer.WriteString(uint64.ToString(CultureInfo.InvariantCulture)); return; }
        if (value is float single) { writer.WriteByte(10); writer.WriteString(single.ToString("R", CultureInfo.InvariantCulture)); return; }
        if (value is double number) { writer.WriteByte(11); writer.WriteString(number.ToString("R", CultureInfo.InvariantCulture)); return; }
        if (value is decimal decimalNumber) { writer.WriteByte(12); writer.WriteString(decimalNumber.ToString(CultureInfo.InvariantCulture)); return; }
        if (value is DateTimeOffset offset) {
            writer.WriteByte(13); writer.WriteInt64(offset.UtcDateTime.Ticks); writer.WriteInt64(offset.Offset.Ticks); return;
        }
        if (value is DateTime date) { writer.WriteByte(14); writer.WriteInt64(date.ToBinary()); return; }
        if (value is TimeSpan duration) { writer.WriteByte(15); writer.WriteInt64(duration.Ticks); return; }
        if (value is Guid guid) { writer.WriteByte(16); writer.WriteBytes(guid.ToByteArray()); return; }
        if (value is char character) { writer.WriteByte(17); writer.WriteInt32(character); return; }
        Type type = value.GetType();
        if (type.IsEnum) {
            writer.WriteByte(18);
            writer.WriteString(type.FullName);
            writer.WriteString(Convert.ToString(value, CultureInfo.InvariantCulture));
            return;
        }
        if (value is IDictionary dictionary) {
            writer.WriteByte(19);
            var entries = new List<DictionaryEntry>();
            foreach (DictionaryEntry entry in dictionary) entries.Add(entry);
            entries.Sort((left, right) => StringComparer.Ordinal.Compare(
                Convert.ToString(left.Key, CultureInfo.InvariantCulture),
                Convert.ToString(right.Key, CultureInfo.InvariantCulture)));
            writer.WriteInt32(entries.Count);
            foreach (DictionaryEntry entry in entries) {
                WriteValue(writer, entry.Key, depth + 1);
                WriteValue(writer, entry.Value, depth + 1);
            }
            return;
        }
        if (value is IEnumerable sequence) {
            writer.WriteByte(20);
            var items = new List<object?>();
            foreach (object? item in sequence) items.Add(item);
            writer.WriteInt32(items.Count);
            foreach (object? item in items) WriteValue(writer, item, depth + 1);
            return;
        }

        writer.WriteByte(21);
        writer.WriteString(type.FullName ?? type.Name);
        writer.WriteString(Convert.ToString(value, CultureInfo.InvariantCulture));
    }
}

internal readonly struct EmailSemanticStreamDigest {
    internal EmailSemanticStreamDigest(byte[] digest, long length) {
        Digest = digest;
        Length = length;
    }
    internal byte[] Digest { get; }
    internal long Length { get; }
}
