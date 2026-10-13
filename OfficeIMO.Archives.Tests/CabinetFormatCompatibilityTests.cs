using OfficeIMO.Core.Internal.Cabinet;

namespace OfficeIMO.Archives.Tests;

public sealed class CabinetFormatCompatibilityTests {
    // MS-CAB section 3.1 checks the payload and then the header/reserve separately.
    // This seven-byte payload gives 0x40302010 XOR 0x00506070 = 0x40604060.
    // The four-byte length header contributes 0x00070007. Reserve bytes A1 B2 C3
    // contribute A1, A1B2 or A1B2C3, independently of the payload's alignment.
    private static readonly byte[] Payload = { 0x10, 0x20, 0x30, 0x40, 0x50, 0x60, 0x70 };

    [Theory]
    [InlineData(1, 0x406740C6u)]
    [InlineData(2, 0x4067E1D5u)]
    [InlineData(3, 0x40C6F2A4u)]
    public void ReadsChecksumsWithOddLengthDataReserves(int reserveLength, uint checksum) {
        byte[] cabinet = Cabinet(Encoding.ASCII.GetBytes("templates.admx"), reserveLength: reserveLength, checksum: checksum);

        OfficeCabinetEntry entry = Assert.Single(Read(cabinet));

        Assert.Equal("templates.admx", entry.Name);
        Assert.Equal(Payload, entry.Data);
    }

    [Theory]
    [InlineData(1, 0x406740C6u, false)]
    [InlineData(2, 0x4067E1D5u, false)]
    [InlineData(3, 0x40C6F2A4u, false)]
    [InlineData(1, 0x406740C6u, true)]
    [InlineData(2, 0x4067E1D5u, true)]
    [InlineData(3, 0x40C6F2A4u, true)]
    public void ChecksumsCoverBothDataReserveAndPayload(int reserveLength, uint checksum, bool corruptPayload) {
        byte[] cabinet = Cabinet(Encoding.ASCII.GetBytes("templates.admx"), reserveLength: reserveLength, checksum: checksum);
        int mutationOffset = corruptPayload ? cabinet.Length - 1 : cabinet.Length - Payload.Length - reserveLength;
        cabinet[mutationOffset] ^= 1;

        Assert.Equal("CAB_CHECKSUM", Assert.Throws<OfficeCabinetException>(() => Read(cabinet)).Code);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ReadsTheFull256EncodedByteMemberName(bool utf8) {
        string name = BoundaryName(utf8, extraBytes: 0);
        byte[] encodedName = Encoding.UTF8.GetBytes(name);
        Assert.Equal(256, encodedName.Length);
        byte[] cabinet = Cabinet(encodedName, utf8);

        OfficeCabinetEntry entry = Assert.Single(Read(cabinet));

        Assert.Equal(name, entry.Name);
        Assert.Equal(Payload, entry.Data);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RejectsMemberNamesAbove256EncodedBytes(bool utf8) {
        byte[] encodedName = Encoding.UTF8.GetBytes(BoundaryName(utf8, extraBytes: 1));
        Assert.Equal(257, encodedName.Length);

        Assert.Equal("CAB_STRING", Assert.Throws<OfficeCabinetException>(() => Read(Cabinet(encodedName, utf8))).Code);
    }

    [Fact]
    public void RejectsLegacyNamesWhoseMeaningDependsOnTheLocale() {
        // B9 means Polish a-ogonek in Windows-1250, but superscript one in Latin-1.
        // Without _A_NAME_IS_UTF there is no portable way to choose between them.
        byte[] name = Encoding.ASCII.GetBytes("templates-").Concat(new byte[] { 0xB9 })
            .Concat(Encoding.ASCII.GetBytes(".admx")).ToArray();

        OfficeCabinetException exception = Assert.Throws<OfficeCabinetException>(() => Read(Cabinet(name)));

        Assert.Equal("CAB_NAME_ENCODING", exception.Code);
    }

    [Theory]
    [InlineData(0, 0x40674067u)]
    [InlineData(1, 0x406740C6u)]
    public void IgnoresReservedHeaderFlagsAndStillReadsDefinedReserveFlag(int reserveLength, uint checksum) {
        byte[] cabinet = Cabinet(Encoding.ASCII.GetBytes("templates.admx"), reserveLength: reserveLength,
            checksum: checksum, headerFlags: 0xFFF8);

        Assert.Equal(Payload, Assert.Single(Read(cabinet)).Data);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    public void ReservedHeaderFlagsDoNotHideUnsupportedCabinetContinuation(ushort continuationFlag) {
        byte[] cabinet = Cabinet(Encoding.ASCII.GetBytes("templates.admx"), headerFlags: (ushort)(0xFFF8 | continuationFlag));

        Assert.Equal("CAB_MULTIPART", Assert.Throws<OfficeCabinetException>(() => Read(cabinet)).Code);
    }

    [Theory]
    [InlineData(0xFFFD)]
    [InlineData(0xFFFE)]
    [InlineData(0xFFFF)]
    public void RecognizesContinuedFileFolderMarkers(ushort folderIndex) {
        byte[] cabinet = Cabinet(Encoding.ASCII.GetBytes("templates.admx"));
        UInt16(cabinet, 44 + 8, folderIndex);

        Assert.Equal("CAB_MULTIPART", Assert.Throws<OfficeCabinetException>(() => Read(cabinet)).Code);
    }

    [Fact]
    public void MissingOrdinaryFolderRemainsMalformed() {
        byte[] cabinet = Cabinet(Encoding.ASCII.GetBytes("templates.admx"));
        UInt16(cabinet, 44 + 8, 1);

        Assert.Equal("CAB_FOLDER", Assert.Throws<OfficeCabinetException>(() => Read(cabinet)).Code);
    }

    private static string BoundaryName(bool utf8, int extraBytes) =>
        new string('a', 120) + "/" + new string('b', (utf8 ? 128 : 130) + extraBytes) + (utf8 ? "\u0105" : "") + ".admx";

    private static IReadOnlyList<OfficeCabinetEntry> Read(byte[] cabinet) =>
        OfficeCabinetArchiveReader.Read(cabinet, 1024, 1024, 10);

    private static byte[] Cabinet(byte[] name, bool utf8 = false, int reserveLength = 0,
        uint checksum = 0x40674067, ushort headerFlags = 0) {
        bool hasReserve = reserveLength != 0;
        int folderOffset = hasReserve ? 40 : 36;
        int filesOffset = folderOffset + 8;
        int dataOffset = filesOffset + 16 + name.Length + 1;
        var cabinet = new byte[dataOffset + 8 + reserveLength + Payload.Length];
        cabinet[0] = (byte)'M'; cabinet[1] = (byte)'S'; cabinet[2] = (byte)'C'; cabinet[3] = (byte)'F';
        UInt32(cabinet, 8, (uint)cabinet.Length);
        UInt32(cabinet, 16, (uint)filesOffset);
        cabinet[24] = 3; cabinet[25] = 1;
        UInt16(cabinet, 26, 1);
        UInt16(cabinet, 28, 1);
        UInt16(cabinet, 30, (ushort)(headerFlags | (hasReserve ? 4 : 0)));
        if (hasReserve) cabinet[39] = (byte)reserveLength;
        UInt32(cabinet, folderOffset, (uint)dataOffset);
        UInt16(cabinet, folderOffset + 4, 1);
        UInt32(cabinet, filesOffset, (uint)Payload.Length);
        if (utf8) UInt16(cabinet, filesOffset + 14, 0x80);
        Buffer.BlockCopy(name, 0, cabinet, filesOffset + 16, name.Length);
        UInt32(cabinet, dataOffset, checksum);
        UInt16(cabinet, dataOffset + 4, (ushort)Payload.Length);
        UInt16(cabinet, dataOffset + 6, (ushort)Payload.Length);
        byte[] reserve = { 0xA1, 0xB2, 0xC3 };
        Buffer.BlockCopy(reserve, 0, cabinet, dataOffset + 8, reserveLength);
        Buffer.BlockCopy(Payload, 0, cabinet, dataOffset + 8 + reserveLength, Payload.Length);
        return cabinet;
    }

    private static void UInt16(byte[] data, int offset, ushort value) {
        data[offset] = (byte)value; data[offset + 1] = (byte)(value >> 8);
    }

    private static void UInt32(byte[] data, int offset, uint value) {
        data[offset] = (byte)value; data[offset + 1] = (byte)(value >> 8);
        data[offset + 2] = (byte)(value >> 16); data[offset + 3] = (byte)(value >> 24);
    }
}
