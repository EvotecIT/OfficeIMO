using OfficeIMO.Core.Internal;

namespace OfficeIMO.Archives.Tests;

public sealed class MsiMetadataBoundaryTests {
    [Theory]
    [InlineData(68, "self")]
    [InlineData(72, "self")]
    [InlineData(68, "shared")]
    [InlineData(72, "shared")]
    [InlineData(68, "root")]
    [InlineData(72, "root")]
    [InlineData(68, "outside")]
    [InlineData(72, "outside")]
    [InlineData(68, "unused")]
    [InlineData(72, "unused")]
    [InlineData(68, "end-of-chain")]
    [InlineData(72, "end-of-chain")]
    public void MalformedDirectorySiblingLinksCannotPassReachability(int linkOffset, string target) {
        byte[] package = ArchiveFixtures.Msi();
        int[] slots = DirectorySlots(package);
        int leaf = Array.FindIndex(slots, offset => package[offset + 66] == 2
            && BitConverter.ToUInt32(package, offset + 68) == uint.MaxValue
            && BitConverter.ToUInt32(package, offset + 72) == uint.MaxValue);
        Assert.True(leaf > 0);
        uint id = target switch {
            "self" => (uint)leaf,
            "shared" => (uint)Array.FindIndex(slots, offset => package[offset + 66] == 2 && offset != slots[leaf]),
            "root" => 0,
            "outside" => (uint)slots.Length,
            "unused" => (uint)Array.FindIndex(slots, offset => package[offset + 66] == 0),
            _ => 0xfffffffeU
        };
        Buffer.BlockCopy(BitConverter.GetBytes(id), 0, package, slots[leaf] + linkOffset, 4);
        AssertInvalid(package);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ReachableEntriesCannotDisappearFromTheStrictDirectoryIndex(bool unknownType) {
        byte[] package = Rewrite(ArchiveFixtures.Msi(), streams => streams.Add("Extra", Array.Empty<byte>()));
        int slot = DirectorySlots(package).Single(offset => Name(package, offset) == "Extra");
        if (unknownType) package[slot + 66] = 3;
        else Buffer.BlockCopy(BitConverter.GetBytes((ushort)0), 0, package, slot + 64, 2);
        AssertInvalid(package);
    }

    [Theory]
    [InlineData(true, 68)]
    [InlineData(true, 72)]
    [InlineData(false, 76)]
    public void DirectoryLinksMustRespectRootAndStreamRoles(bool root, int linkOffset) {
        byte[] package = ArchiveFixtures.Msi();
        int[] slots = DirectorySlots(package);
        int entry = root ? 0 : Array.FindIndex(slots, offset => package[offset + 66] == 2);
        Assert.True(entry >= 0);
        Buffer.BlockCopy(BitConverter.GetBytes((uint)entry), 0, package, slots[entry] + linkOffset, 4);
        AssertInvalid(package);
    }

    [Fact]
    public void ActiveOrphanStorageCannotHideBehindAReachableStreamName() {
        byte[] package = ArchiveFixtures.Msi();
        int slot = DirectorySlots(package).First(offset => package[offset + 66] == 0);
        WriteEntry(package, slot, MsiPackageReader.EncodeStreamName("payload.cab", false), 1, Guid.Empty);
        AssertInvalid(package);
    }

    [Fact]
    public void AdditionalRootCannotSubstituteItsInstallerClassForTheAllocationRoot() {
        byte[] package = ArchiveFixtures.Msi();
        int[] slots = DirectorySlots(package);
        Buffer.BlockCopy(Guid.Empty.ToByteArray(), 0, package, slots[0] + 80, 16);
        int spare = slots.First(offset => package[offset + 66] == 0);
        WriteEntry(package, spare, "A", 5, new Guid("000c1084-0000-0000-c000-000000000046"));
        AssertInvalid(package);
    }

    [Theory]
    [InlineData('/')]
    [InlineData('\\')]
    [InlineData(':')]
    [InlineData('!')]
    public void IllegalRawStreamNamesAreMalformedRatherThanEmbeddedStorage(char separator) {
        byte[] package = Rewrite(ArchiveFixtures.Msi(), streams => streams.Add("Extra", Array.Empty<byte>()));
        int slot = DirectorySlots(package).Single(offset => Name(package, offset) == "Extra");
        WriteName(package, slot, "A" + separator + "B");
        AssertInvalid(package);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MalformedSummaryTimesUseTheMalformedPackageException(bool oleDate) {
        OfficeOleProperty property = oleDate ? OfficeOleProperty.Double(12, double.NaN) : OfficeOleProperty.Int64(12, -1);
        property.ValueBytes[0] = oleDate ? (byte)7 : (byte)0x40;
        byte[] summary = OfficeOlePropertySetWriter.CreatePropertySet((OfficeOlePropertySetWriter.SummaryInformationFormatId,
            OfficeOlePropertySetWriter.CreateSection(new[] { OfficeOleProperty.Int32(15, 2), property })));
        byte[] package = Rewrite(ArchiveFixtures.Msi(), streams => streams.Add(
            OfficeOlePropertySetWriter.SummaryInformationStreamName, summary));
        AssertInvalid(package);
    }

    private static void AssertInvalid(byte[] package) {
        using var fixture = new ArchiveFixtures();
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(package), fixture.Destination));
        Assert.False(Directory.Exists(fixture.Destination));
        Assert.Empty(Directory.GetDirectories(fixture.Root, ".officeimo-archive-*"));
        Assert.Empty(Directory.GetDirectories(fixture.Root, ".officeimo-msi-alias-*"));
    }

    private static byte[] Rewrite(byte[] package, Action<Dictionary<string, byte[]>> change) {
        Assert.True(OfficeCompoundFileReader.TryRead(package, out Dictionary<string, byte[]> streams, out string? error), error);
        change(streams);
        return OfficeCompoundFileWriter.Write(streams.Select(pair => new OfficeCompoundStream(pair.Key, pair.Value)).ToArray(),
            new Guid("000c1084-0000-0000-c000-000000000046"));
    }

    private static int[] DirectorySlots(byte[] package) {
        // These small writer fixtures are CFBv3 with one FAT sector. Follow its
        // directory chain rather than assuming the directory sectors are adjacent.
        Assert.Equal(1, BitConverter.ToInt32(package, 44));
        int fat = (BitConverter.ToInt32(package, 76) + 1) * 512;
        uint sector = BitConverter.ToUInt32(package, 48);
        var slots = new List<int>();
        while (sector != 0xfffffffe) {
            int offset = checked(((int)sector + 1) * 512);
            for (int entry = 0; entry < 4; entry++) slots.Add(offset + entry * 128);
            sector = BitConverter.ToUInt32(package, fat + checked((int)sector * 4));
            Assert.True(slots.Count < 64);
        }
        return slots.ToArray();
    }

    private static string Name(byte[] package, int offset) {
        int bytes = BitConverter.ToUInt16(package, offset + 64);
        return bytes == 0 ? string.Empty : Encoding.Unicode.GetString(package, offset, bytes - 2);
    }

    private static void WriteEntry(byte[] package, int offset, string name, byte type, Guid classId) {
        Array.Clear(package, offset, 128);
        WriteName(package, offset, name);
        package[offset + 66] = type; package[offset + 67] = 1;
        foreach (int link in new[] { 68, 72, 76 }) Buffer.BlockCopy(BitConverter.GetBytes(uint.MaxValue), 0, package, offset + link, 4);
        Buffer.BlockCopy(classId.ToByteArray(), 0, package, offset + 80, 16);
        Buffer.BlockCopy(BitConverter.GetBytes(0xfffffffeU), 0, package, offset + 116, 4);
    }

    private static void WriteName(byte[] package, int offset, string name) {
        byte[] value = Encoding.Unicode.GetBytes(name + "\0");
        Assert.True(value.Length <= 64);
        Array.Clear(package, offset, 64);
        Buffer.BlockCopy(value, 0, package, offset, value.Length);
        Buffer.BlockCopy(BitConverter.GetBytes((ushort)value.Length), 0, package, offset + 64, 2);
    }
}
