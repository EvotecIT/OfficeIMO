namespace OfficeIMO.OneNote;

/// <summary>Computes the four-byte longitudinal parity checksum defined by MS-CAB.</summary>
internal static class OneNoteCabinetChecksum {
    internal static uint Compute(byte[] data, int offset, int count) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        if (offset < 0 || count < 0 || offset > data.Length - count) throw new ArgumentOutOfRangeException(nameof(offset));

        uint checksum = 0;
        int end = offset + count;
        while (offset <= end - 4) {
            checksum ^= (uint)(data[offset] |
                (data[offset + 1] << 8) |
                (data[offset + 2] << 16) |
                (data[offset + 3] << 24));
            offset += 4;
        }

        uint remainder = 0;
        while (offset < end) remainder = (remainder << 8) | data[offset++];
        return checksum ^ remainder;
    }
}
