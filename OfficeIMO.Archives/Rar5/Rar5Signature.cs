namespace OfficeIMO.Archives.Rar5;

internal static class Rar5Signature {
    private const int MaxStubBytes = 1024 * 1024;
    private static readonly byte[] Prefix = { (byte)'R', (byte)'a', (byte)'r', (byte)'!', 0x1a, 7 };

    internal static bool IsRar5(byte[] buffer, int length, int offset = 0) =>
        length - offset >= 8 && buffer[offset + 6] == 1 && buffer[offset + 7] == 0 && MatchesPrefix(buffer, offset);

    internal static bool IsRar4(byte[] buffer, int length, int offset = 0) =>
        length - offset >= 7 && buffer[offset + 6] == 0 && MatchesPrefix(buffer, offset);

    private static bool MatchesPrefix(byte[] buffer, int offset) {
        for (int i = 0; i < Prefix.Length; i++) {
            if (buffer[offset + i] != Prefix[i]) return false;
        }
        return true;
    }

    internal static NotSupportedException Legacy(long offset) =>
        new(FormattableString.Invariant($"RAR 4 not supported (signature at byte offset {offset}). This reader accepts RAR 5."));

    internal static bool TryFind(Stream source, CancellationToken token) {
        const int searchChunk = 8192;
        var buffer = new byte[searchChunk + 7];
        source.Position = 0;
        int carry = 0;
        int position = 0;
        while (position <= MaxStubBytes) {
            int count = carry;
            int capacity = Math.Min(buffer.Length, MaxStubBytes + 8 - position);
            while (count < capacity) {
                token.ThrowIfCancellationRequested();
                int read = source.Read(buffer, count, capacity - count);
                if (read == 0) break;
                count += read;
            }
            for (int i = 0; i < count && position + i <= MaxStubBytes; i++) {
                if (buffer[i] != Prefix[0]) continue;
                if (IsRar4(buffer, count, i)) throw Legacy(position + i);
                if (IsRar5(buffer, count, i)) {
                    source.Position = position + i + 8;
                    return true;
                }
            }
            if (count < capacity || position + count >= MaxStubBytes + 8) break;
            position += count - 7;
            carry = 7;
            Buffer.BlockCopy(buffer, count - carry, buffer, 0, carry);
        }
        return false;
    }
}
