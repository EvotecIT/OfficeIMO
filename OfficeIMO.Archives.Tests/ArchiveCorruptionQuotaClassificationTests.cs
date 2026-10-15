namespace OfficeIMO.Archives.Tests;

public sealed class ArchiveCorruptionQuotaClassificationTests {
    [Theory]
    [InlineData("block-size", 1, "CAB_BLOCK_SIZE")]
    [InlineData("block-size", 65535, "CAB_BLOCK_SIZE")]
    [InlineData("stored-size", 1, "CAB_UNCOMPRESSED")]
    [InlineData("stored-size", 2, "CAB_UNCOMPRESSED")]
    [InlineData("mszip-header", 1, "CAB_MSZIP_HEADER")]
    [InlineData("mszip-header", 2, "CAB_MSZIP_HEADER")]
    [InlineData("mszip-magic", 1, "CAB_MSZIP_HEADER")]
    [InlineData("mszip-magic", 2, "CAB_MSZIP_HEADER")]
    [InlineData("mszip-size", 1, "CAB_MSZIP_HEADER")]
    [InlineData("mszip-size", 2, "CAB_MSZIP_HEADER")]
    public void CheapMalformedFrameFactsPrecedeExpansionAdmission(string defect, long maximum, string code) {
        using var fixture = new ArchiveFixtures();
        byte[] cabinet = MalformedFrame(defect);
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(
            fixture.Package(cabinet), fixture.Destination, new ArchiveExtractionOptions { MaxExpandedBytes = maximum }));
        Assert.StartsWith(code + ":", error.Message);
        Assert.False(Directory.Exists(fixture.Destination));
        Assert.Empty(Directory.GetDirectories(fixture.Root, ".officeimo-archive-*"));
    }

    [Theory]
    [InlineData("block-size", 1, "CAB_BLOCK_SIZE")]
    [InlineData("block-size", 65535, "CAB_BLOCK_SIZE")]
    [InlineData("stored-size", 1, "CAB_UNCOMPRESSED")]
    [InlineData("stored-size", 2, "CAB_UNCOMPRESSED")]
    [InlineData("mszip-header", 1, "CAB_MSZIP_HEADER")]
    [InlineData("mszip-header", 2, "CAB_MSZIP_HEADER")]
    [InlineData("mszip-magic", 1, "CAB_MSZIP_HEADER")]
    [InlineData("mszip-magic", 2, "CAB_MSZIP_HEADER")]
    [InlineData("mszip-size", 1, "CAB_MSZIP_HEADER")]
    [InlineData("mszip-size", 2, "CAB_MSZIP_HEADER")]
    public void PassiveKnownFrameStructureIsValidatedBeforeExpansionAdmission(string defect, long maximum, string code) {
        byte[] cabinet = MalformedFrame(defect);
        using var source = new MemoryStream(cabinet); source.Position = 7;
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => ArchiveInspector.Open(source,
            new ArchiveInspectionOptions { MaxExpandedBytes = maximum }));
        Assert.StartsWith(code + ":", error.Message);
        Assert.Equal(7, source.Position); Assert.True(source.CanRead);
    }

    [Theory]
    [InlineData(4, 1)]
    [InlineData(4, 3)]
    [InlineData(12, 1)]
    [InlineData(12, 3)]
    [InlineData(20, 1)]
    [InlineData(20, 3)]
    [InlineData(8, 1)]
    [InlineData(8, 3)]
    [InlineData(24, 1)]
    [InlineData(24, 3)]
    public void IntrinsicHeaderValidityPrecedesPassiveDirectoryAdmission(int field, int maximum) {
        byte[] cabinet = ArchiveFixtures.StoredCab("file", new byte[] { 1 });
        if (field == 8) Array.Clear(cabinet, 8, 4);
        else cabinet[field] = field == 24 ? (byte)2 : (byte)1;
        using var source = new MemoryStream(cabinet); source.Position = 7;
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => ArchiveInspector.Open(source,
            new ArchiveInspectionOptions { MaxDirectoryRecords = maximum }));
        Assert.StartsWith("CAB_HEADER:", error.Message);
        Assert.Equal(7, source.Position); Assert.True(source.CanRead);
    }

    [Theory]
    [InlineData("expanded", "MaxExpandedBytes", 1)]
    [InlineData("directory", "MaxDirectoryRecords", 1)]
    [InlineData("work", "MaxDecompressionWorkBytes", 32767)]
    public void ValidCabinetStillRefusesRealQuotaExhaustionBeforeReturningPayload(string bound, string limit, long maximum) {
        byte[] cabinet = ArchiveFixtures.StoredCab("file", bound == "expanded" ? new byte[] { 1, 2 } : new byte[] { 1 });
        using var source = new MemoryStream(cabinet);
        OfficeInspectionResourceLimitException error;
        if (bound == "work") {
            using ArchiveInspection inspection = ArchiveInspector.Open(source,
                new ArchiveInspectionOptions { MaxDecompressionWorkBytes = maximum });
            error = Assert.Throws<OfficeInspectionResourceLimitException>(() => inspection.OpenRead(0));
        } else error = Assert.Throws<OfficeInspectionResourceLimitException>(() => ArchiveInspector.Open(source,
            new ArchiveInspectionOptions { MaxExpandedBytes = bound == "expanded" ? maximum : 10,
                MaxDirectoryRecords = bound == "directory" ? (int)maximum : 3 }));
        Assert.Equal("archives", error.Owner); Assert.Equal(limit, error.LimitName); Assert.Equal(maximum, error.Maximum);
    }

    private static byte[] MalformedFrame(string defect) {
        int payloadLength = defect == "mszip-size" ? 32781 : defect == "mszip-magic" ? 2 : 1;
        byte[] cabinet = ArchiveFixtures.StoredCab("file", new byte[payloadLength]);
        // Keep the member declaration valid and small so the public entry quota
        // cannot obscure the independently malformed CFDATA record.
        int file = BitConverter.ToInt32(cabinet, 16);
        cabinet[file] = 1; Array.Clear(cabinet, file + 1, 3);
        int frame = BitConverter.ToInt32(cabinet, 36);
        ushort expanded = defect == "block-size" ? (ushort)65535 : (ushort)2;
        cabinet[frame + 6] = (byte)expanded; cabinet[frame + 7] = (byte)(expanded >> 8);
        if (defect.StartsWith("mszip", StringComparison.Ordinal)) cabinet[42] = 1;
        if (defect == "mszip-size") { cabinet[frame + 8] = 0x43; cabinet[frame + 9] = 0x4b; }
        return cabinet;
    }
}
