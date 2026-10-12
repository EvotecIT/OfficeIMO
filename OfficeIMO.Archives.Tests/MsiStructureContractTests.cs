using OfficeIMO.Core.Internal;

namespace OfficeIMO.Archives.Tests;

public sealed class MsiStructureContractTests {
    [Fact]
    public void TableRowsDoNotLimitAuxiliaryCompoundStreamsOrDirectorySlots() {
        byte[] package = Rewrite(streams => {
            for (int index = 0; index < 40; index++) streams.Add("Aux" + index, Array.Empty<byte>());
        });
        using var fixture = new ArchiveFixtures();
        var result = ArchiveExtractor.Extract(fixture.Package(package), fixture.Destination,
            new ArchiveExtractionOptions { MaxTableRows = 23 });
        Assert.Single(result.Files);
    }

    [Theory]
    [InlineData("Directory", 's')]
    [InlineData("Directory", 'i')]
    [InlineData("Component", 's')]
    [InlineData("Component", 'I')]
    [InlineData("File", 's')]
    [InlineData("File", 'i')]
    [InlineData("Media", 's')]
    [InlineData("Media", 'I')]
    [InlineData("File", 'b')]
    public void AdditionalPersistentColumnsPreserveRequiredTableProjection(string table, char kind) {
        using var fixture = new ArchiveFixtures();
        var result = ArchiveExtractor.Extract(fixture.Package(ArchiveFixtures.Msi(extraColumnTable: table, extraColumnKind: kind)),
            fixture.Destination);
        Assert.Single(result.Files);
        Assert.Equal("hello", File.ReadAllText(Path.Combine(fixture.Destination, "PolicyDefinitions/en-US/Template.adml")));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AuxiliaryStreamsCannotShareAnotherStreamsAllocation(bool mini) {
        byte[] package = Rewrite(streams => {
            streams.Add("AuxA", new byte[mini ? 64 : 5000]);
            streams.Add("AuxB", new byte[mini ? 64 : 5000]);
        });
        int[] slots = Slots(package);
        int a = slots.Single(offset => Name(package, offset) == "AuxA");
        int b = slots.Single(offset => Name(package, offset) == "AuxB");
        Buffer.BlockCopy(package, a + 116, package, b + 116, 12);
        AssertInvalid(package);
    }

    [Theory]
    [InlineData(false, "free")]
    [InlineData(false, "cycle")]
    [InlineData(false, "outside")]
    [InlineData(true, "free")]
    [InlineData(true, "cycle")]
    [InlineData(true, "outside")]
    public void EveryDeclaredAuxiliaryChainMustTerminateExactly(bool mini, string fault) {
        byte[] package = Rewrite(streams => streams.Add("AuxA", new byte[mini ? 64 : 5000]));
        int slot = Slots(package).Single(offset => Name(package, offset) == "AuxA");
        uint first = BitConverter.ToUInt32(package, slot + 116);
        int allocation = mini ? checked(((int)BitConverter.ToUInt32(package, 60) + 1) * 512)
            : checked(((int)BitConverter.ToUInt32(package, 76) + 1) * 512);
        uint last = first;
        if (!mini) for (int index = 1; index < 10; index++) last = BitConverter.ToUInt32(package, allocation + (int)last * 4);
        uint link = fault == "free" ? uint.MaxValue : fault == "cycle" ? first : 100000U;
        Buffer.BlockCopy(BitConverter.GetBytes(link), 0, package, allocation + (int)last * 4, 4);
        AssertInvalid(package);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(0xfffffffeU)]
    public void RegularAuxiliaryStreamsCannotUseInternalStructureSectors(uint rootOrFat) {
        byte[] package = Rewrite(streams => streams.Add("AuxA", new byte[5000]));
        int[] slots = Slots(package);
        int aux = slots.Single(offset => Name(package, offset) == "AuxA");
        uint first = rootOrFat == 0 ? BitConverter.ToUInt32(package, slots[0] + 116)
            : BitConverter.ToUInt32(package, 76);
        Buffer.BlockCopy(BitConverter.GetBytes(first), 0, package, aux + 116, 4);
        AssertInvalid(package);
    }

    [Fact]
    public void AnUnusedAuxiliaryStreamCannotAliasMetadataAllocation() {
        byte[] package = Rewrite(streams => streams.Add("AuxA", new byte[1]));
        int[] slots = Slots(package);
        int aux = slots.Single(offset => Name(package, offset) == "AuxA");
        int metadata = slots.Single(offset => Name(package, offset) == MsiPackageReader.EncodeStreamName("_StringData", true));
        Buffer.BlockCopy(package, metadata + 116, package, aux + 116, 12);
        AssertInvalid(package);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DisjointAuxiliaryStreamsRemainInertAndValid(bool mini) {
        byte[] package = Rewrite(streams => {
            streams.Add("AuxA", new byte[mini ? 64 : 5000]);
            streams.Add("AuxB", new byte[mini ? 64 : 5000]);
        });
        using var fixture = new ArchiveFixtures();
        Assert.Single(ArchiveExtractor.Extract(fixture.Package(package), fixture.Destination).Files);
    }

    [Fact]
    public void DirectoryChainMustEndWithEndOfChainRatherThanFreeSector() {
        byte[] package = ArchiveFixtures.Msi();
        int fat = checked(((int)BitConverter.ToUInt32(package, 76) + 1) * 512);
        uint last = BitConverter.ToUInt32(package, 48);
        while (BitConverter.ToUInt32(package, fat + (int)last * 4) != 0xfffffffeU)
            last = BitConverter.ToUInt32(package, fat + (int)last * 4);
        Buffer.BlockCopy(BitConverter.GetBytes(uint.MaxValue), 0, package, fat + (int)last * 4, 4);
        AssertInvalid(package);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DifatAllocationMarkersAreValidatedWithoutOpeningAuxiliaryPayloads(bool corrupt) {
        byte[] package = Rewrite(streams => streams.Add("AuxA", new byte[8 * 1024 * 1024]));
        uint difat = BitConverter.ToUInt32(package, 68);
        Assert.True(BitConverter.ToUInt32(package, 72) > 0);
        if (corrupt) {
            int fatIndex = (int)difat / 128;
            uint fatSector = fatIndex < 109 ? BitConverter.ToUInt32(package, 76 + fatIndex * 4)
                : BitConverter.ToUInt32(package, checked(((int)difat + 1) * 512 + (fatIndex - 109) * 4));
            int entry = checked(((int)fatSector + 1) * 512 + (int)(difat % 128) * 4);
            Buffer.BlockCopy(BitConverter.GetBytes(0xfffffffeU), 0, package, entry, 4);
            AssertInvalid(package);
        } else {
            using var fixture = new ArchiveFixtures();
            Assert.Single(ArchiveExtractor.Extract(fixture.Package(package), fixture.Destination).Files);
        }
    }

    private static void AssertInvalid(byte[] package) {
        using var fixture = new ArchiveFixtures();
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(package), fixture.Destination));
        Assert.False(Directory.Exists(fixture.Destination));
    }
    private static byte[] Rewrite(Action<Dictionary<string, byte[]>> edit) {
        Assert.True(OfficeCompoundFileReader.TryRead(ArchiveFixtures.Msi(), out Dictionary<string, byte[]> streams, out string? error), error);
        edit(streams);
        return OfficeCompoundFileWriter.Write(streams.Select(pair => new OfficeCompoundStream(pair.Key, pair.Value)).ToArray(),
            new Guid("000c1084-0000-0000-c000-000000000046"));
    }
    private static int[] Slots(byte[] package) {
        var slots = new List<int>();
        int fat = checked(((int)BitConverter.ToUInt32(package, 76) + 1) * 512);
        for (uint sector = BitConverter.ToUInt32(package, 48); sector != 0xfffffffe; sector = BitConverter.ToUInt32(package, fat + (int)sector * 4)) {
            for (int index = 0; index < 4; index++) slots.Add(checked(((int)sector + 1) * 512 + index * 128));
            Assert.True(slots.Count < 128);
        }
        return slots.ToArray();
    }
    private static string Name(byte[] package, int offset) {
        int length = BitConverter.ToUInt16(package, offset + 64);
        return length < 2 ? "" : Encoding.Unicode.GetString(package, offset, length - 2);
    }
}
