namespace OfficeIMO.Archives.Tests;

public sealed partial class Rar5ArchiveTests {
    [Fact]
    public void PublishesExplicitEmptyDirectoriesAndSplitFilesThroughTheStagingExtractor() {
        using var standalone = new ArchiveFixtures();
        ArchiveExtractor.Extract(standalone.Package(Fixture("directories.rar")), standalone.Destination);
        Assert.Empty(Directory.EnumerateFileSystemEntries(Path.Combine(standalone.Destination, "empty-directory")));
        Assert.Equal("The lighthouse is open.\n", File.ReadAllText(Path.Combine(standalone.Destination, "notes", "señal.txt")));

        using var split = new ArchiveFixtures();
        string[] parts = Enumerable.Range(1, 3).Select(index =>
            Path.Combine(AppContext.BaseDirectory, "Fixtures", "IndependentRar5", $"split-blake.part{index}.rar")).ToArray();
        var extraction = ArchiveExtractor.ExtractRar5Volumes(parts, split.Destination);
        Assert.Equal(ArchiveFormat.Rar5, extraction.Format);
        Assert.Equal(Piece(), File.ReadAllBytes(Path.Combine(split.Destination, Assert.Single(extraction.Files).RelativePath)));
    }

    [Theory]
    [InlineData("split-crc")]
    [InlineData("split-blake")]
    public void JoinsAndChecksAllSplitPayloadPartsAcrossOrderedVolumes(string prefix) {
        byte[][] original = Enumerable.Range(1, 3).Select(i => Fixture($"{prefix}.part{i}.rar")).ToArray();
        using var first = new MemoryStream(original[0]);
        using var second = new MemoryStream(original[1]);
        using var third = new MemoryStream(original[2]);
        Stream[] set = { first, second, third };
        var info = ArchiveReader.InspectRar5Volumes(set);
        Assert.Equal(3, info.VolumeCount);
        var entry = Assert.Single(info.Entries);
        Assert.Equal("piece.bin", entry.Name);
        Assert.Equal(Piece().Length, entry.PackedLength);
        foreach (Stream volume in set) volume.Position = 0;
        Assert.Equal(Piece(), Assert.Single(ArchiveReader.ReadRar5Volumes(set)).Data);

        foreach (Stream volume in set) volume.Position = 0;
        Assert.Throws<InvalidDataException>(() => ArchiveReader.ReadRar5Volumes(new[] { first, third, second }));
        first.Position = 0;
        Assert.Throws<InvalidDataException>(() => ArchiveReader.ReadRar5(first));
        foreach (Stream volume in set) volume.Position = 0;
        var aggregate = new ArchiveReadOptions { MaxInputBytes = original.Sum(volume => volume.Length) - 1 };
        Assert.Throws<OfficeInspectionResourceLimitException>(() => ArchiveReader.ReadRar5Volumes(set, aggregate));

        byte[] piece = Piece();
        int firstData = Find(original[0], piece.Take(32).ToArray());
        Assert.True(firstData >= 0);
        original[0][firstData] ^= 1;
        foreach (Stream volume in set) volume.Position = 0;
        var error = Assert.Throws<InvalidDataException>(() => ArchiveReader.ReadRar5Volumes(set));
        Assert.Contains(prefix == "split-crc" ? "CRC-32" : "BLAKE2sp", error.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ChecksTheFinalCrcWhenTheMiddleVolumeHasNoCrc(bool corruptPayload) {
        const int header = 25, flag = 35, checksum = 41;
        byte[] original = Fixture("split-crc.part2.rar");
        byte[] modified = original.Take(checksum).Concat(original.Skip(checksum + 4)).ToArray();
        modified[header + 4] -= 4;
        modified[flag] &= 0xfb;
        int payload = header + 5 + modified[header + 4];
        RefreshCrc(modified, header, payload);
        if (corruptPayload) modified[payload] ^= 1;

        using var first = new MemoryStream(Fixture("split-crc.part1.rar"));
        using var second = new MemoryStream(modified);
        using var third = new MemoryStream(Fixture("split-crc.part3.rar"));
        Stream[] volumes = { first, second, third };
        if (corruptPayload) {
            var error = Assert.Throws<InvalidDataException>(() => ArchiveReader.ReadRar5Volumes(volumes));
            Assert.Contains("CRC-32", error.Message, StringComparison.Ordinal);
        } else {
            Assert.Equal(Piece(), Assert.Single(ArchiveReader.ReadRar5Volumes(volumes)).Data);
        }
    }

    [Fact]
    public void RejectsAContinuationThatDeclaresADirectory() {
        const int header = 25, flag = 35;
        byte[] modified = Fixture("split-crc.part2.rar");
        modified[flag] |= 1;
        RefreshCrc(modified, header, header + 5 + modified[header + 4]);
        using var first = new MemoryStream(Fixture("split-crc.part1.rar"));
        using var second = new MemoryStream(modified);
        using var third = new MemoryStream(Fixture("split-crc.part3.rar"));
        var error = Assert.Throws<InvalidDataException>(() => ArchiveReader.ReadRar5Volumes(new[] { first, second, third }));
        Assert.Contains("volume 2, byte offset 25", error.Message, StringComparison.Ordinal);
    }
}
