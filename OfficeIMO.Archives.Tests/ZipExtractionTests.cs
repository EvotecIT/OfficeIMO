using System.IO.Compression;
using System.Text;

namespace OfficeIMO.Archives.Tests;

public sealed class ZipExtractionTests {
    [Fact]
    public void StoredAndDeflateMembersAreVerifiedAndPublished() {
        using var fixture = new ArchiveFixtures();
        byte[] zip = Zip(("PolicyDefinitions/", null, CompressionLevel.Optimal),
            ("PolicyDefinitions/Example.admx", "<policyDefinitions />", CompressionLevel.Optimal),
            ("PolicyDefinitions/en-US/Example.adml", "<policyDefinitionResources />", CompressionLevel.NoCompression));
        ArchiveExtractionResult result = ArchiveExtractor.Extract(fixture.Package(zip), fixture.Destination);
        Assert.Equal(ArchiveFormat.Zip, result.Format);
        Assert.Equal(new[] { "PolicyDefinitions/Example.admx", "PolicyDefinitions/en-US/Example.adml" },
            result.Files.Select(file => file.RelativePath));
        Assert.Equal("<policyDefinitions />", File.ReadAllText(Path.Combine(fixture.Destination, "PolicyDefinitions", "Example.admx")));
        Assert.Empty(Directory.GetDirectories(fixture.Root, ".officeimo-archive-*"));
    }

    [Fact]
    public void EmptyArchivesPublishAnEmptyDirectory() {
        using var fixture = new ArchiveFixtures();
        ArchiveExtractionResult result = ArchiveExtractor.Extract(fixture.Package(Zip()), fixture.Destination);
        Assert.Equal(ArchiveFormat.Zip, result.Format);
        Assert.Empty(result.Files);
        Assert.True(Directory.Exists(fixture.Destination));
    }

    [Fact]
    public void EmptyZip64ArchivesPublishAnEmptyDirectory() {
        using var fixture = new ArchiveFixtures();
        ArchiveExtractionResult result = ArchiveExtractor.Extract(fixture.Package(WithZip64Locator(Zip(), 0)), fixture.Destination);
        Assert.Equal(ArchiveFormat.Zip, result.Format);
        Assert.Empty(result.Files);
        Assert.True(Directory.Exists(fixture.Destination));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MembersStartingOnAnotherDiskFailBeforeSelection(bool directory) {
        using var fixture = new ArchiveFixtures();
        byte[] zip = Zip((directory ? "Templates/" : "Template.admx", directory ? null : "A", CompressionLevel.NoCompression));
        int central = (int)BitConverter.ToUInt32(zip, zip.Length - 6);
        WriteUInt16(zip, central + 34, 1);
        int selectionCalls = 0;
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(zip), fixture.Destination,
            new ArchiveExtractionOptions { IncludeFile = _ => { selectionCalls++; return false; } }));
        Assert.Equal(0, selectionCalls);
        Assert.False(Directory.Exists(fixture.Destination));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ClassicArchivesAcceptExactly65535EntriesUnlessAnEarlierLocatorMakesThemAmbiguous(bool earlierLocator) {
        // A small hand-authored directory exercises the actual classic count, without a ZIP64 locator.
        using var fixture = new ArchiveFixtures();
        using var bytes = new MemoryStream();
        using (var writer = new BinaryWriter(bytes, Encoding.ASCII, leaveOpen: true)) {
            var offsets = new uint[ushort.MaxValue];
            for (int index = 0; index < ushort.MaxValue; index++) {
                byte[] name = Encoding.ASCII.GetBytes("f" + index);
                offsets[index] = (uint)bytes.Length;
                writer.Write(0x04034b50U); writer.Write((ushort)20); writer.Write(new byte[20]);
                writer.Write((ushort)name.Length); writer.Write((ushort)(earlierLocator && index == 0 ? 24 : 0)); writer.Write(name);
                if (earlierLocator && index == 0) {
                    writer.Write((ushort)0xabcd); writer.Write((ushort)20);
                    writer.Write(0x07064b50U); writer.Write(0U); writer.Write(0UL); writer.Write(1U);
                }
            }
            uint directory = (uint)bytes.Length;
            for (int index = 0; index < ushort.MaxValue; index++) {
                byte[] name = Encoding.ASCII.GetBytes("f" + index);
                writer.Write(0x02014b50U); writer.Write((ushort)20); writer.Write((ushort)20);
                writer.Write(new byte[20]);
                writer.Write((ushort)name.Length); writer.Write(new byte[12]); writer.Write(offsets[index]); writer.Write(name);
            }
            uint size = (uint)bytes.Length - directory;
            writer.Write(0x06054b50U); writer.Write(0U);
            writer.Write(ushort.MaxValue); writer.Write(ushort.MaxValue);
            writer.Write(size); writer.Write(directory); writer.Write((ushort)0);
        }
        string package = fixture.Package(bytes.ToArray());
        var options = new ArchiveExtractionOptions { IncludeFile = _ => false };
        if (earlierLocator) {
            var error = Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(package, fixture.Destination, options));
            Assert.Contains("earlier ZIP64 locator", error.Message, StringComparison.Ordinal);
            Assert.False(Directory.Exists(fixture.Destination));
            return;
        }
        var result = ArchiveExtractor.Extract(package, fixture.Destination, options);
        Assert.Empty(result.Files);
        Assert.True(Directory.Exists(fixture.Destination));
    }

    [Fact]
    public void AllowedFormatsRejectDetectedContentBeforeSelectionOrStaging() {
        using var fixture = new ArchiveFixtures();
        int calls = 0;
        Assert.Throws<NotSupportedException>(() => ArchiveExtractor.Extract(
            fixture.Package(Zip(("Template.admx", "A", CompressionLevel.Optimal))), fixture.Destination,
            new ArchiveExtractionOptions { AllowedFormats = new[] { ArchiveFormat.Cab }, IncludeFile = _ => { calls++; return true; } }));
        Assert.Equal(0, calls);
        Assert.False(Directory.Exists(fixture.Destination));
        Assert.Empty(Directory.GetDirectories(fixture.Root, ".officeimo-archive-*"));
    }

    [Theory]
    [InlineData("../escape.admx")]
    [InlineData("/rooted.admx")]
    [InlineData("C:/drive.admx")]
    [InlineData("a/../b.admx")]
    [InlineData("a//b.admx")]
    [InlineData("CON.admx")]
    [InlineData("a:stream.admx")]
    [InlineData("Templates /Example.admx")]
    [InlineData("Example.admx.")]
    [InlineData("..\\escape.admx")]
    public void UnsafeOrNonportableMemberNamesFailBeforeAnyWrite(string name) {
        using var fixture = new ArchiveFixtures();
        int selectionCalls = 0;
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(
            fixture.Package(Zip(("Safe.admx", "A", CompressionLevel.Optimal), (name, "B", CompressionLevel.Optimal))), fixture.Destination,
            new ArchiveExtractionOptions { IncludeFile = path => { selectionCalls++; return false; } }));
        Assert.True(selectionCalls <= 1);
        Assert.False(Directory.Exists(fixture.Destination));
        Assert.Empty(Directory.GetDirectories(fixture.Root, ".officeimo-archive-*"));
    }

    [Theory]
    [InlineData("Example.admx", "Example.admx")]
    [InlineData("Example.admx", "EXAMPLE.ADMX")]
    [InlineData("Templates", "Templates/Example.admx")]
    [InlineData("Templates/Example.admx", "templates/Other.admx")]
    public void CollidingMembersFailInsteadOfOverwriting(string first, string second) {
        using var fixture = new ArchiveFixtures();
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(
            fixture.Package(Zip((first, "A", CompressionLevel.Optimal), (second, "B", CompressionLevel.Optimal))), fixture.Destination));
        Assert.False(Directory.Exists(fixture.Destination));
    }

    [Fact]
    public void DirectoryRecordsFollowTheSameCollisionRules() {
        using var fixture = new ArchiveFixtures();
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(
            fixture.Package(Zip(("Templates", "A", CompressionLevel.Optimal), ("Templates/", null, CompressionLevel.Optimal))), fixture.Destination));
        Assert.False(Directory.Exists(fixture.Destination));
    }

    [Fact]
    public void SelectionPublishesOnlyIncludedFiles() {
        using var fixture = new ArchiveFixtures();
        ArchiveExtractionResult result = ArchiveExtractor.Extract(fixture.Package(Zip(
                ("Example.admx", "A", CompressionLevel.Optimal), ("setup.exe", "MZ", CompressionLevel.Optimal))), fixture.Destination,
            new ArchiveExtractionOptions { IncludeFile = path => path.EndsWith(".admx", StringComparison.OrdinalIgnoreCase) });
        Assert.Equal("Example.admx", Assert.Single(result.Files).RelativePath);
        Assert.False(File.Exists(Path.Combine(fixture.Destination, "setup.exe")));
    }

    [Fact]
    public void MemberCountIsLimitedFromTheCentralDirectory() {
        using var fixture = new ArchiveFixtures();
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(Zip(
                ("a.admx", "A", CompressionLevel.Optimal), ("b.admx", "B", CompressionLevel.Optimal), ("c.admx", "C", CompressionLevel.Optimal))),
            fixture.Destination, new ArchiveExtractionOptions { MaxEntryCount = 2 }));
        Assert.False(Directory.Exists(fixture.Destination));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExpansionLimitsApplyBeforeWriting(bool aggregate) {
        using var fixture = new ArchiveFixtures();
        string payload = new('x', 4096);
        var options = aggregate
            ? new ArchiveExtractionOptions { MaxExpandedBytes = 6000 }
            : new ArchiveExtractionOptions { MaxEntryBytes = 4000 };
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(Zip(
            ("a.admx", payload, CompressionLevel.Optimal), ("b.admx", payload, CompressionLevel.Optimal))), fixture.Destination, options));
        Assert.False(Directory.Exists(fixture.Destination));
    }

    [Fact]
    public void EncryptedMembersAreNotSupported() {
        using var fixture = new ArchiveFixtures();
        byte[] zip = PatchHeaders(Zip(("Example.admx", "A", CompressionLevel.NoCompression)), flagsToSet: 0x0001);
        Assert.Throws<NotSupportedException>(() => ArchiveExtractor.Extract(fixture.Package(zip), fixture.Destination));
        Assert.False(Directory.Exists(fixture.Destination));
    }

    [Fact]
    public void CompressionMethodsOtherThanStoredAndDeflateAreNotSupported() {
        using var fixture = new ArchiveFixtures();
        byte[] zip = PatchHeaders(Zip(("Example.admx", "A", CompressionLevel.NoCompression)), method: 12);
        Assert.Throws<NotSupportedException>(() => ArchiveExtractor.Extract(fixture.Package(zip), fixture.Destination));
        Assert.False(Directory.Exists(fixture.Destination));
    }

    [Fact]
    public void ChecksumMismatchesFailTheWholeExtraction() {
        using var fixture = new ArchiveFixtures();
        byte[] zip = PatchHeaders(Zip(("Example.admx", "payload", CompressionLevel.NoCompression)), crcXor: 0x00010000);
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(zip), fixture.Destination));
        Assert.False(Directory.Exists(fixture.Destination));
        Assert.Empty(Directory.GetDirectories(fixture.Root, ".officeimo-archive-*"));
    }

    [Fact]
    public void NonAsciiNamesRequireTheUtf8Flag() {
        using (var fixture = new ArchiveFixtures()) {
            ArchiveExtractionResult result = ArchiveExtractor.Extract(
                fixture.Package(Zip(("Szablony/zażółć.admx", "A", CompressionLevel.Optimal))), fixture.Destination);
            Assert.Equal("Szablony/zażółć.admx", Assert.Single(result.Files).RelativePath);
        }
        using (var fixture = new ArchiveFixtures()) {
            byte[] unflagged = PatchHeaders(Zip(("Szablony/zażółć.admx", "A", CompressionLevel.Optimal)), flagsToClear: 0x0800);
            Assert.Throws<NotSupportedException>(() => ArchiveExtractor.Extract(fixture.Package(unflagged), fixture.Destination));
            Assert.False(Directory.Exists(fixture.Destination));
        }
    }

    [Fact]
    public void ZipContentIsDetectedIndependentlyOfTheFileName() {
        using var fixture = new ArchiveFixtures();
        string package = Path.Combine(fixture.Root, "MicrosoftEdgePolicyTemplates.cab");
        File.WriteAllBytes(package, Zip(("windows/admx/msedge.admx", "A", CompressionLevel.Optimal)));
        Assert.Equal(ArchiveFormat.Zip, ArchiveExtractor.Extract(package, fixture.Destination).Format);
    }

    [Fact]
    public void ALaterEndRecordInsideTheCommentIsRejected() {
        // ZipArchive takes the last end-record signature; the scanner takes the one whose comment reaches the end.
        using var fixture = new ArchiveFixtures();
        byte[] zip = Zip(("a.admx", "A", CompressionLevel.NoCompression));
        byte[] inner = zip.Skip(zip.Length - 22).ToArray();
        byte[] comment = inner.Concat(new byte[] { 0 }).ToArray();
        byte[] crafted = zip.Take(zip.Length - 2).Concat(BitConverter.GetBytes((ushort)comment.Length)).Concat(comment).ToArray();
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(crafted), fixture.Destination));
        Assert.False(Directory.Exists(fixture.Destination));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void Zip64LocatorsMustAgreeWithClassicFields(int offsetSkew) {
        using var fixture = new ArchiveFixtures();
        byte[] crafted = WithZip64Locator(Zip(("a.admx", "A", CompressionLevel.NoCompression)), offsetSkew);
        if (offsetSkew == 0) {
            Assert.Equal("a.admx", Assert.Single(ArchiveExtractor.Extract(fixture.Package(crafted), fixture.Destination).Files).RelativePath);
        } else {
            Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(crafted), fixture.Destination));
            Assert.False(Directory.Exists(fixture.Destination));
        }
    }

    [Fact]
    public void SaturatedClassicCountMustMatchTheActualDirectory() {
        // 65,535 can be a classic count, but must describe exactly that many file headers.
        using var fixture = new ArchiveFixtures();
        byte[] zip = Zip(("a.admx", "A", CompressionLevel.NoCompression));
        int end = zip.Length - 22;
        zip[end + 8] = 0xff; zip[end + 9] = 0xff; zip[end + 10] = 0xff; zip[end + 11] = 0xff;
        var error = Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(zip), fixture.Destination));
        Assert.Contains("declares 65535 entries but contains 1", error.Message, StringComparison.Ordinal);
        Assert.False(Directory.Exists(fixture.Destination));
    }

    [Fact]
    public void SignatureBytesNearTheEndAreRejectedOnEveryTarget() {
        // .NET 10 selects even a truncated later signature while older runtimes skip it; reject it consistently.
        using var fixture = new ArchiveFixtures();
        byte[] zip = Zip(("a.admx", "A", CompressionLevel.NoCompression));
        byte[] comment = { 0x50, 0x4b, 0x05, 0x06, 0 };
        byte[] crafted = zip.Take(zip.Length - 2).Concat(BitConverter.GetBytes((ushort)comment.Length)).Concat(comment).ToArray();
        var error = Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(crafted), fixture.Destination));
        Assert.Contains("another end-of-central-directory signature", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void FileHeadersBeyondTheDeclaredDirectoryAreRejected() {
        // .NET Framework reads file headers until another signature appears, regardless of the declared count.
        using var fixture = new ArchiveFixtures();
        byte[] zip = Zip(("a.admx", "A", CompressionLevel.NoCompression));
        int end = zip.Length - 22;
        int directory = (int)BitConverter.ToUInt32(zip, end + 16);
        byte[] record = zip.Skip(directory).Take(end - directory).ToArray();
        byte[] crafted = zip.Take(end).Concat(record).Concat(zip.Skip(end)).ToArray();
        var error = Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(crafted), fixture.Destination));
        Assert.Contains("beyond the declared directory size", error.Message, StringComparison.Ordinal);
        Assert.False(Directory.Exists(fixture.Destination));
    }

    [Fact]
    public void NamesFlaggedAsUtf8MustBeValidUtf8() {
        using var fixture = new ArchiveFixtures();
        byte[] zip = PatchHeaders(Zip(("zx.admx", "A", CompressionLevel.Optimal)), flagsToSet: 0x0800);
        byte[] marker = Encoding.ASCII.GetBytes("zx.admx");
        for (int offset = 0; offset <= zip.Length - marker.Length; offset++)
            if (zip.Skip(offset).Take(marker.Length).SequenceEqual(marker)) zip[offset + 1] = 0xff;
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(zip), fixture.Destination));
        Assert.False(Directory.Exists(fixture.Destination));
    }

    [Theory]
    [InlineData("/")]
    [InlineData("\\")]
    [InlineData("Templates //")]
    public void InvalidDirectoryRecordsAreRejected(string name) {
        using var fixture = new ArchiveFixtures();
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(
            fixture.Package(Zip((name, null, CompressionLevel.Optimal))), fixture.Destination));
        Assert.False(Directory.Exists(fixture.Destination));
    }

    [Fact]
    public void DirectoryRecordsWithContentAreRejected() {
        using var fixture = new ArchiveFixtures();
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(
            fixture.Package(Zip(("Templates/", "content", CompressionLevel.NoCompression))), fixture.Destination));
        Assert.False(Directory.Exists(fixture.Destination));
    }

    [Fact]
    public void ADirectoryRecordFollowedByAFileOfTheSameNameIsRejected() {
        using var fixture = new ArchiveFixtures();
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(
            fixture.Package(Zip(("Templates/", null, CompressionLevel.Optimal), ("Templates", "A", CompressionLevel.Optimal))), fixture.Destination));
        Assert.False(Directory.Exists(fixture.Destination));
    }

    [Fact]
    public void StoredMembersShorterThanTheirDeclaredLengthAreRejected() {
        using var fixture = new ArchiveFixtures();
        byte[] zip = PatchHeaders(Zip(("Example.admx", "payload", CompressionLevel.NoCompression)), uncompressedDelta: 4);
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(zip), fixture.Destination));
        Assert.False(Directory.Exists(fixture.Destination));
    }

    private static byte[] WithZip64Locator(byte[] zip, int offsetSkew) {
        int end = zip.Length - 22;
        ushort entries = BitConverter.ToUInt16(zip, end + 10);
        uint size = BitConverter.ToUInt32(zip, end + 12);
        uint offset = BitConverter.ToUInt32(zip, end + 16);
        using var record = new MemoryStream();
        using (var writer = new BinaryWriter(record, Encoding.ASCII, leaveOpen: true)) {
            writer.Write(0x06064b50U); writer.Write(44UL); writer.Write((ushort)45); writer.Write((ushort)45);
            writer.Write(0U); writer.Write(0U); writer.Write((ulong)entries); writer.Write((ulong)entries);
            writer.Write((ulong)size); writer.Write((ulong)(offset + offsetSkew));
            writer.Write(0x07064b50U); writer.Write(0U); writer.Write((ulong)end); writer.Write(1U);
        }
        return zip.Take(end).Concat(record.ToArray()).Concat(zip.Skip(end)).ToArray();
    }

    private static byte[] Zip(params (string Name, string? Text, CompressionLevel Level)[] members) {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach (var member in members) {
                ZipArchiveEntry entry = archive.CreateEntry(member.Name, member.Level);
                if (member.Text == null) continue;
                using Stream output = entry.Open();
                byte[] bytes = Encoding.UTF8.GetBytes(member.Text);
                output.Write(bytes, 0, bytes.Length);
            }
        }
        return stream.ToArray();
    }

    // Rewrites the matching fields of every local and central file header.
    private static byte[] PatchHeaders(byte[] zip, ushort flagsToSet = 0, ushort flagsToClear = 0, ushort? method = null, uint crcXor = 0,
        uint uncompressedDelta = 0) {
        byte[] bytes = (byte[])zip.Clone();
        for (int offset = 0; offset <= bytes.Length - 4; offset++) {
            uint signature = BitConverter.ToUInt32(bytes, offset);
            int fields = signature == 0x04034b50U ? offset + 6 : signature == 0x02014b50U ? offset + 8 : -1;
            if (fields < 0) continue;
            ushort flags = (ushort)((BitConverter.ToUInt16(bytes, fields) | flagsToSet) & ~flagsToClear);
            WriteUInt16(bytes, fields, flags);
            if (method.HasValue) WriteUInt16(bytes, fields + 2, method.Value);
            uint crc = BitConverter.ToUInt32(bytes, fields + 8) ^ crcXor;
            WriteUInt16(bytes, fields + 8, (ushort)crc);
            WriteUInt16(bytes, fields + 10, (ushort)(crc >> 16));
            uint uncompressed = BitConverter.ToUInt32(bytes, fields + 16) + uncompressedDelta;
            WriteUInt16(bytes, fields + 16, (ushort)uncompressed);
            WriteUInt16(bytes, fields + 18, (ushort)(uncompressed >> 16));
        }
        return bytes;
    }

    private static void WriteUInt16(byte[] bytes, int offset, ushort value) {
        bytes[offset] = (byte)value;
        bytes[offset + 1] = (byte)(value >> 8);
    }
}
