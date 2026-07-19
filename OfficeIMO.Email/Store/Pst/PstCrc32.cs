namespace OfficeIMO.Email.Store;

internal static class PstCrc32 {
    internal static uint Compute(byte[] bytes) =>
        Compute(bytes, 0, bytes?.Length ?? 0);

    internal static uint Compute(byte[] bytes, int count) => Compute(bytes, 0, count);

    internal static uint Compute(byte[] bytes, int offset, int count) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        if (offset < 0 || count < 0 || offset > bytes.Length - count) {
            throw new ArgumentOutOfRangeException(nameof(count));
        }
        uint checksum = 0;
        for (int index = offset; index < offset + count; index++) {
            checksum ^= bytes[index];
            for (int bit = 0; bit < 8; bit++) {
                checksum = (checksum & 1) != 0
                    ? (checksum >> 1) ^ 0xEDB88320U
                    : checksum >> 1;
            }
        }
        return checksum;
    }
}
