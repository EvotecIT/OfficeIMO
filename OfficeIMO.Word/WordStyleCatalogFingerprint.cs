using System.Security.Cryptography;

namespace OfficeIMO.Word;

internal readonly struct WordStyleCatalogFingerprint : IEquatable<WordStyleCatalogFingerprint> {
    private const uint EndOfCentralDirectorySignature = 0x06054B50;
    private const uint CentralDirectoryHeaderSignature = 0x02014B50;
    private const uint LocalFileHeaderSignature = 0x04034B50;
    private const string StylesEntryName = "word/styles.xml";

    private WordStyleCatalogFingerprint(ulong first, ulong second, ulong third, ulong fourth) {
        First = first;
        Second = second;
        Third = third;
        Fourth = fourth;
    }

    private ulong First { get; }
    private ulong Second { get; }
    private ulong Third { get; }
    private ulong Fourth { get; }

    internal static WordStyleCatalogFingerprint? TryCreate(byte[] packageBytes) {
        int end = FindEndOfCentralDirectory(packageBytes);
        if (end < 0) return null;
        uint centralOffset = ReadUInt32(packageBytes, end + 16);
        ushort entryCount = ReadUInt16(packageBytes, end + 10);
        if (centralOffset == uint.MaxValue || entryCount == ushort.MaxValue || centralOffset > int.MaxValue) return null;

        int position = (int)centralOffset;
        for (int index = 0; index < entryCount; index++) {
            if (!HasBytes(packageBytes, position, 46) || ReadUInt32(packageBytes, position) != CentralDirectoryHeaderSignature) {
                return null;
            }
            ushort nameLength = ReadUInt16(packageBytes, position + 28);
            ushort extraLength = ReadUInt16(packageBytes, position + 30);
            ushort commentLength = ReadUInt16(packageBytes, position + 32);
            if (!HasBytes(packageBytes, position + 46, nameLength)) return null;

            if (NameEquals(packageBytes, position + 46, nameLength, StylesEntryName)) {
                uint compressedSize = ReadUInt32(packageBytes, position + 20);
                uint localOffset = ReadUInt32(packageBytes, position + 42);
                if (compressedSize == uint.MaxValue || localOffset == uint.MaxValue ||
                    compressedSize > int.MaxValue || localOffset > int.MaxValue) return null;
                int local = (int)localOffset;
                if (!HasBytes(packageBytes, local, 30) || ReadUInt32(packageBytes, local) != LocalFileHeaderSignature) {
                    return null;
                }
                long dataOffsetValue = local + 30L + ReadUInt16(packageBytes, local + 26) + ReadUInt16(packageBytes, local + 28);
                if (dataOffsetValue > int.MaxValue) return null;
                int dataOffset = (int)dataOffsetValue;
                if (!HasBytes(packageBytes, dataOffset, (int)compressedSize)) return null;
                using SHA256 sha256 = SHA256.Create();
                byte[] hash = sha256.ComputeHash(packageBytes, dataOffset, (int)compressedSize);
                return new WordStyleCatalogFingerprint(
                    BitConverter.ToUInt64(hash, 0),
                    BitConverter.ToUInt64(hash, 8),
                    BitConverter.ToUInt64(hash, 16),
                    BitConverter.ToUInt64(hash, 24));
            }

            long nextPosition = position + 46L + nameLength + extraLength + commentLength;
            if (nextPosition > packageBytes.Length) return null;
            position = (int)nextPosition;
        }
        return null;
    }

    private static int FindEndOfCentralDirectory(byte[] bytes) {
        int minimum = Math.Max(0, bytes.Length - 65_557);
        for (int index = bytes.Length - 22; index >= minimum; index--) {
            if (ReadUInt32(bytes, index) == EndOfCentralDirectorySignature &&
                index + 22L + ReadUInt16(bytes, index + 20) == bytes.Length) return index;
        }
        return -1;
    }

    private static bool NameEquals(byte[] bytes, int offset, int length, string expected) {
        if (length != expected.Length) return false;
        for (int index = 0; index < length; index++) {
            if (bytes[offset + index] != (byte)expected[index]) return false;
        }
        return true;
    }

    private static bool HasBytes(byte[] bytes, int offset, int count) =>
        offset >= 0 && count >= 0 && offset <= bytes.Length - count;

    private static ushort ReadUInt16(byte[] bytes, int offset) =>
        HasBytes(bytes, offset, 2)
            ? (ushort)(bytes[offset] | bytes[offset + 1] << 8)
            : (ushort)0;

    private static uint ReadUInt32(byte[] bytes, int offset) =>
        HasBytes(bytes, offset, 4)
            ? (uint)(bytes[offset] | bytes[offset + 1] << 8 | bytes[offset + 2] << 16 | bytes[offset + 3] << 24)
            : 0;

    public bool Equals(WordStyleCatalogFingerprint other) =>
        First == other.First && Second == other.Second && Third == other.Third && Fourth == other.Fourth;

    public override bool Equals(object? obj) => obj is WordStyleCatalogFingerprint other && Equals(other);

    public override int GetHashCode() {
        unchecked {
            int hash = First.GetHashCode();
            hash = (hash * 397) ^ Second.GetHashCode();
            hash = (hash * 397) ^ Third.GetHashCode();
            hash = (hash * 397) ^ Fourth.GetHashCode();
            return hash;
        }
    }
}
