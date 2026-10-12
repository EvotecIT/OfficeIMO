using OfficeIMO.Core.Internal;

namespace OfficeIMO.Archives.Tests;

public sealed class MsiFormatContractTests {
    [Theory]
    [InlineData("File", "FileSize", 'i')]
    [InlineData("File", "Attributes", 'I')]
    [InlineData("Component", "Attributes", 'I')]
    [InlineData("Media", "DiskId", 'I')]
    public void RequiredNumericColumnsRetainTheirStandardWidth(string table, string column, char kind) {
        AssertInvalid(ArchiveFixtures.Msi(editDatabase: (schemas, _) => {
            int index = Array.FindIndex(schemas[table], field => field.Name == column);
            schemas[table][index] = (column, kind);
        }));
    }

    [Fact]
    public void LargePackageSequenceColumnsCanUseFourByteIntegers() {
        byte[] package = ArchiveFixtures.Msi(wideMedia: true, editDatabase: (schemas, _) => {
            schemas["File"][7] = ("Sequence", 'I');
        });
        AssertValid(package);
    }

    [Theory]
    [InlineData("PayloadKey", "../payload")]
    [InlineData("PayloadKey", "1payload")]
    [InlineData("PayloadKey", ".payload")]
    [InlineData("COMP", "Cömponent")]
    [InlineData("COMP", "component-name")]
    [InlineData("POLICY", "../directory")]
    [InlineData("LOCALE", "9locale")]
    public void MatchingIdentifierReferencesCannotMakeInvalidKeysValid(string oldId, string newId) {
        AssertInvalid(RenameIdentifier(oldId, newId));
    }

    [Fact]
    public void UnusedComponentsStillRequireValidIdentifiers() {
        AssertInvalid(ArchiveFixtures.Msi(editDatabase: (_, rows) => {
            rows["Component"] = rows["Component"].Concat(new[] {
                new object?[] { "../unused", null, "LOCALE", 0, null, null }
            }).ToArray();
        }));
    }

    [Fact]
    public void OptionalComponentKeyPathRetainsIdentifierGrammar() {
        AssertInvalid(ArchiveFixtures.Msi(editDatabase: (_, rows) => rows["Component"][0][5] = "../registry"));
    }

    [Theory]
    [InlineData(72, true)]
    [InlineData(73, false)]
    public void StandardIdentifierColumnsRetainTheirMaximumLength(int length, bool accepted) {
        byte[] package = RenameIdentifier("PayloadKey", new string('A', length));
        if (accepted) AssertValid(package); else AssertInvalid(package);
    }

    [Theory]
    [InlineData("PayloadKey")]
    [InlineData("COMP")]
    [InlineData("LOCALE")]
    public void IdentifiersMayContainAnInitialUnderscoreAndLaterDigitsOrPeriods(string oldId) =>
        AssertValid(RenameIdentifier(oldId, "_Valid.123"));

    [Fact]
    public void ZeroMiniStreamCutoffCannotTurnSmallMetadataIntoRegularStreams() {
        byte[] package = RegularOnlyMsi();
        // All streams are physically valid regular chains at the authored zero cutoff.
        // Generic compound lookup remains compatible; strict MSI must reject its header.
        Assert.True(OfficeCompoundFileReader.TryRead(package, out Dictionary<string, byte[]> streams, out string? error), error);
        Assert.Contains(MsiPackageReader.EncodeStreamName("_StringPool", true), streams.Keys);
        AssertInvalid(package);
    }

    private static byte[] RenameIdentifier(string oldId, string newId) => ArchiveFixtures.Msi(
        fileKey: oldId == "PayloadKey" ? newId : "PayloadKey", editDatabase: (_, rows) => {
            foreach (object?[][] table in rows.Values) foreach (object?[] row in table)
                for (int index = 0; index < row.Length; index++)
                    if (row[index] is string value && value == oldId) row[index] = newId;
        });

    private static byte[] RegularOnlyMsi() {
        byte[] source = ArchiveFixtures.Msi();
        Assert.True(OfficeCompoundFileReader.TryRead(source, out Dictionary<string, byte[]> streams, out string? error), error);
        int fatOffset = checked(((int)BitConverter.ToUInt32(source, 76) + 1) * 512);
        using var directory = new MemoryStream();
        for (uint sector = BitConverter.ToUInt32(source, 48); sector != 0xfffffffeU;
            sector = BitConverter.ToUInt32(source, fatOffset + (int)sector * 4))
            directory.Write(source, checked(((int)sector + 1) * 512), 512);
        byte[] entries = directory.ToArray();
        int directorySectors = entries.Length / 512;
        int sectorCount = directorySectors + streams.Values.Sum(data => (data.Length + 511) / 512) + 1;
        Assert.True(sectorCount < 128);
        byte[] output = new byte[(sectorCount + 1) * 512];
        Buffer.BlockCopy(source, 0, output, 0, 512);
        void UInt32(int offset, uint value) => Buffer.BlockCopy(BitConverter.GetBytes(value), 0, output, offset, 4);
        UInt32(44, 1); UInt32(48, 0); UInt32(56, 0); UInt32(60, 0xfffffffeU); UInt32(64, 0);
        UInt32(68, 0xfffffffeU); UInt32(72, 0); UInt32(76, (uint)(sectorCount - 1));
        int fat = sectorCount * 512;
        for (int offset = fat; offset < output.Length; offset += 4) UInt32(offset, uint.MaxValue);
        for (int sector = 0; sector < directorySectors; sector++)
            UInt32(fat + sector * 4, sector + 1 == directorySectors ? 0xfffffffeU : (uint)(sector + 1));
        Buffer.BlockCopy(entries, 0, output, 512, entries.Length);
        UInt32(512 + 116, 0xfffffffeU); UInt32(512 + 120, 0); UInt32(512 + 124, 0);
        int next = directorySectors;
        for (int slot = 128; slot < entries.Length; slot += 128) {
            if (entries[slot + 66] != 2) continue;
            string name = Encoding.Unicode.GetString(entries, slot, BitConverter.ToUInt16(entries, slot + 64) - 2);
            byte[] data = streams[name];
            int count = (data.Length + 511) / 512;
            UInt32(512 + slot + 116, count == 0 ? 0xfffffffeU : (uint)next);
            Buffer.BlockCopy(data, 0, output, (next + 1) * 512, data.Length);
            for (int index = 0; index < count; index++)
                UInt32(fat + (next + index) * 4, index + 1 == count ? 0xfffffffeU : (uint)(next + index + 1));
            next += count;
        }
        UInt32(fat + (sectorCount - 1) * 4, 0xfffffffdU);
        return output;
    }

    private static void AssertValid(byte[] package) {
        using var fixture = new ArchiveFixtures();
        Assert.Single(ArchiveExtractor.Extract(fixture.Package(package), fixture.Destination).Files);
        Assert.Equal("hello", File.ReadAllText(Path.Combine(fixture.Destination, "PolicyDefinitions/en-US/Template.adml")));
    }

    private static void AssertInvalid(byte[] package) {
        using var fixture = new ArchiveFixtures();
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(package), fixture.Destination));
        Assert.False(Directory.Exists(fixture.Destination));
    }
}
