using OfficeIMO.Drawing.Internal;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingCompoundFileTests {
    private const int SectorSize = 4096;
    private const uint FreeSect = 0xffffffff;
    private const uint EndOfChain = 0xfffffffe;
    private const uint FatSect = 0xfffffffd;

    [Fact]
    public void ReaderFindsVersion4SectorsAfterThePaddedHeaderSector() {
        byte[] compound = CreateVersion4RootOnlyCompoundFile();

        bool success = OfficeCompoundFileReader.TryRead(compound, out OfficeCompoundFile? file, out string? error);

        Assert.True(success, error);
        Assert.NotNull(file);
        Assert.Empty(file!.Streams);
        Assert.Contains(file.Entries, entry => entry.IsStorage && entry.Name == "Root Entry");
    }

    private static byte[] CreateVersion4RootOnlyCompoundFile() {
        byte[] compound = new byte[SectorSize * 3];
        byte[] signature = { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 };
        Buffer.BlockCopy(signature, 0, compound, 0, signature.Length);
        WriteUInt16(compound, 24, 0x003e);
        WriteUInt16(compound, 26, 0x0004);
        WriteUInt16(compound, 28, 0xfffe);
        WriteUInt16(compound, 30, 0x000c);
        WriteUInt16(compound, 32, 0x0006);
        WriteUInt32(compound, 40, 1);
        WriteUInt32(compound, 44, 1);
        WriteUInt32(compound, 48, 0);
        WriteUInt32(compound, 56, 4096);
        WriteUInt32(compound, 60, EndOfChain);
        WriteUInt32(compound, 68, EndOfChain);
        for (int index = 0; index < 109; index++) {
            WriteUInt32(compound, 76 + index * 4, index == 0 ? 1U : FreeSect);
        }

        int directoryOffset = SectorSize;
        byte[] rootName = Encoding.Unicode.GetBytes("Root Entry\0");
        Buffer.BlockCopy(rootName, 0, compound, directoryOffset, rootName.Length);
        WriteUInt16(compound, directoryOffset + 64, checked((ushort)rootName.Length));
        compound[directoryOffset + 66] = 5;
        compound[directoryOffset + 67] = 1;
        WriteUInt32(compound, directoryOffset + 68, FreeSect);
        WriteUInt32(compound, directoryOffset + 72, FreeSect);
        WriteUInt32(compound, directoryOffset + 76, FreeSect);
        WriteUInt32(compound, directoryOffset + 116, EndOfChain);

        int fatOffset = SectorSize * 2;
        WriteUInt32(compound, fatOffset, EndOfChain);
        WriteUInt32(compound, fatOffset + 4, FatSect);
        for (int index = 2; index < SectorSize / 4; index++) {
            WriteUInt32(compound, fatOffset + index * 4, FreeSect);
        }

        return compound;
    }

    private static void WriteUInt16(byte[] bytes, int offset, ushort value) {
        bytes[offset] = (byte)value;
        bytes[offset + 1] = (byte)(value >> 8);
    }

    private static void WriteUInt32(byte[] bytes, int offset, uint value) {
        bytes[offset] = (byte)value;
        bytes[offset + 1] = (byte)(value >> 8);
        bytes[offset + 2] = (byte)(value >> 16);
        bytes[offset + 3] = (byte)(value >> 24);
    }
}
