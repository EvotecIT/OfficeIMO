namespace OfficeIMO.OneNote;

internal static class OneNoteExtendedGuidReader {
    public static OneNoteExtendedGuid Read(byte[] data, int offset) {
        OneNoteBinary.EnsureRange(data, offset, 1);
        byte first = data[offset];
        if (first == 0) {
            return new OneNoteExtendedGuid(Guid.Empty, 0, 1);
        }

        if ((first & 0x07) == 0x04) {
            OneNoteBinary.EnsureRange(data, offset, 17);
            return new OneNoteExtendedGuid(
                OneNoteBinary.ReadGuid(data, offset + 1),
                (uint)(first >> 3),
                17);
        }

        ushort firstTwo = OneNoteBinary.ReadUInt16(data, offset);
        if ((firstTwo & 0x3F) == 0x20) {
            OneNoteBinary.EnsureRange(data, offset, 18);
            return new OneNoteExtendedGuid(
                OneNoteBinary.ReadGuid(data, offset + 2),
                (uint)(firstTwo >> 6),
                18);
        }

        OneNoteBinary.EnsureRange(data, offset, 3);
        uint firstThree = (uint)(data[offset] | (data[offset + 1] << 8) | (data[offset + 2] << 16));
        if ((firstThree & 0x7F) == 0x40) {
            OneNoteBinary.EnsureRange(data, offset, 19);
            return new OneNoteExtendedGuid(
                OneNoteBinary.ReadGuid(data, offset + 3),
                firstThree >> 7,
                19);
        }

        if (first == 0x80) {
            OneNoteBinary.EnsureRange(data, offset, 21);
            return new OneNoteExtendedGuid(
                OneNoteBinary.ReadGuid(data, offset + 5),
                OneNoteBinary.ReadUInt32(data, offset + 1),
                21);
        }

        throw new OneNoteFormatException(
            "ONENOTE_FSSHTTP_EXTENDED_GUID",
            "The package contains an invalid MS-FSSHTTPB extended GUID encoding.",
            offset);
    }
}
