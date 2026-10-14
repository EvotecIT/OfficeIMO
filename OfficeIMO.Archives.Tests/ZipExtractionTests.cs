using System.IO.Compression;
using System.Text;

namespace OfficeIMO.Archives.Tests;

public sealed partial class ZipExtractionTests {
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

    [Fact]
    public void DeflateMembersExpandingBeyondTheirDeclaredPrefixAreRejected() {
        using var fixture = new ArchiveFixtures();
        byte[] zip = Zip(("Example.admx", "AB", CompressionLevel.Optimal));
        int central = (int)BitConverter.ToUInt32(zip, zip.Length - 6);
        WriteUInt32(zip, 22, 1); WriteUInt32(zip, central + 24, 1);
        WriteUInt32(zip, 14, 0xd3d99e8b); WriteUInt32(zip, central + 16, 0xd3d99e8b); // CRC of A only
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(zip), fixture.Destination));
        Assert.False(Directory.Exists(fixture.Destination));
    }

    [Theory]
    [InlineData(1, 0)]
    [InlineData(0, 8)]
    public void ConflictingLocalHeaderMetadataFailsEvenForExcludedMembers(ushort flags, ushort method) {
        using var fixture = new ArchiveFixtures();
        byte[] zip = Zip(("Example.admx", "A", CompressionLevel.NoCompression));
        int central = (int)BitConverter.ToUInt32(zip, zip.Length - 6);
        WriteUInt16(zip, 6, flags);
        if (method != 0) WriteUInt16(zip, 8, (ushort)(BitConverter.ToUInt16(zip, central + 10) == 0 ? 8 : 0));
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(zip), fixture.Destination,
            new ArchiveExtractionOptions { IncludeFile = _ => false }));
        Assert.False(Directory.Exists(fixture.Destination));
    }

    [Fact]
    public void PatchedDataSetsAreRejectedBeforeSelection() {
        using var fixture = new ArchiveFixtures();
        byte[] zip = PatchHeaders(Zip(("Example.admx", "A", CompressionLevel.NoCompression)), flagsToSet: 0x0020);
        int calls = 0;
        Assert.Throws<NotSupportedException>(() => ArchiveExtractor.Extract(fixture.Package(zip), fixture.Destination,
            new ArchiveExtractionOptions { IncludeFile = _ => { calls++; return false; } }));
        Assert.Equal(0, calls);
        Assert.False(Directory.Exists(fixture.Destination));
    }

    [Theory]
    [InlineData(0, false)]
    [InlineData(1, false)]
    [InlineData(0, true)]
    public void Zip64DiskStartIsResolvedFromTheExtraField(uint disk, bool allCoordinates) {
        using var fixture = new ArchiveFixtures();
        byte[] original = Zip(("Example.admx", "A", CompressionLevel.NoCompression));
        int end = original.Length - 22;
        int central = (int)BitConverter.ToUInt32(original, end + 16);
        int extra = central + 46 + BitConverter.ToUInt16(original, central + 28);
        ushort extraLength = (ushort)(allCoordinates ? 32 : 8);
        byte[] zip = original.Take(extra).Concat(new byte[extraLength]).Concat(original.Skip(extra)).ToArray();
        WriteUInt16(zip, central + 34, ushort.MaxValue); WriteUInt16(zip, central + 30, extraLength);
        WriteUInt16(zip, extra, 1); WriteUInt16(zip, extra + 2, (ushort)(extraLength - 4));
        if (allCoordinates) {
            WriteUInt32(zip, extra + 4, BitConverter.ToUInt32(original, central + 24));
            WriteUInt32(zip, extra + 12, BitConverter.ToUInt32(original, central + 20));
            WriteUInt32(zip, extra + 20, BitConverter.ToUInt32(original, central + 42));
            WriteUInt32(zip, central + 20, uint.MaxValue); WriteUInt32(zip, central + 24, uint.MaxValue); WriteUInt32(zip, central + 42, uint.MaxValue);
        }
        WriteUInt32(zip, extra + extraLength - 4, disk);
        WriteUInt32(zip, end + extraLength + 12, BitConverter.ToUInt32(original, end + 12) + extraLength);
        var bytes = OfficeIMO.Core.Internal.OfficeArchiveSafety.ScanZipCentralDirectory(zip, 10);
        using var stream = new MemoryStream(zip, writable: false);
        var streaming = OfficeIMO.Core.Internal.OfficeArchiveSafety.ScanZipCentralDirectory(stream, stream.Length, 10);
        Assert.Equal(disk == 0, bytes.IsValid);
        Assert.Equal(disk == 0, streaming.IsValid);
        if (disk != 0) Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(zip), fixture.Destination));
        else Assert.Equal("Example.admx", Assert.Single(ArchiveExtractor.Extract(fixture.Package(zip), fixture.Destination).Files).RelativePath);
    }

    [Fact]
    public void SaturatedClassicCountRejectsALowerLimitBeforeReadingTheBody() {
        byte[] zip = new byte[1024 * 1024];
        int end = zip.Length - 22;
        WriteUInt32(zip, end, 0x06054b50); WriteUInt16(zip, end + 8, ushort.MaxValue); WriteUInt16(zip, end + 10, ushort.MaxValue);
        using var source = new CountingStream(zip);
        var result = OfficeIMO.Core.Internal.OfficeArchiveSafety.ScanZipCentralDirectory(source, source.Length, 5);
        Assert.True(result.LimitExceeded);
        Assert.True(source.BytesRead < 70000, "The low entry limit should be decided from the tail.");
        Assert.Equal(0, source.Position);
    }

    private sealed class CountingStream(byte[] bytes) : MemoryStream(bytes, writable: false) {
        internal long BytesRead { get; private set; }
        public override int Read(byte[] buffer, int offset, int count) {
            int read = base.Read(buffer, offset, count);
            BytesRead += read;
            return read;
        }
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, true)]
    [InlineData(true, false)]
    public void DeflateRequiresAFinalBlockAndExactCompressedFraming(bool final, bool trailing) {
        using var fixture = new ArchiveFixtures();
        byte[] compressed = new byte[] { (byte)(final ? 1 : 0), 1, 0, 0xfe, 0xff, 0x41 };
        if (trailing) compressed = compressed.Concat(new byte[] { 0 }).ToArray();
        byte[] zip = ZipWithRawDeflate(compressed, 1, 0xd3d99e8b);
        if (final && !trailing) {
            Assert.Single(ArchiveExtractor.Extract(fixture.Package(zip), fixture.Destination).Files);
            Assert.Equal("A", File.ReadAllText(Path.Combine(fixture.Destination, "Example.admx")));
        } else {
            Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(zip), fixture.Destination));
            Assert.False(Directory.Exists(fixture.Destination));
        }
    }

    private static byte[] ZipWithRawDeflate(byte[] compressed, uint expanded, uint crc) {
        byte[] original = Zip(("Example.admx", "A", CompressionLevel.Optimal));
        int central = (int)BitConverter.ToUInt32(original, original.Length - 6);
        int payload = 30 + BitConverter.ToUInt16(original, 26) + BitConverter.ToUInt16(original, 28);
        byte[] zip = original.Take(payload).Concat(compressed).Concat(original.Skip(central)).ToArray();
        central = payload + compressed.Length;
        WriteUInt16(zip, 6, 0); WriteUInt16(zip, central + 8, 0);
        WriteUInt16(zip, 8, 8); WriteUInt16(zip, central + 10, 8);
        WriteUInt32(zip, 14, crc); WriteUInt32(zip, central + 16, crc);
        WriteUInt32(zip, 18, (uint)compressed.Length); WriteUInt32(zip, central + 20, (uint)compressed.Length);
        WriteUInt32(zip, 22, expanded); WriteUInt32(zip, central + 24, expanded);
        WriteUInt32(zip, zip.Length - 6, (uint)central);
        return zip;
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void StreamDeflateValidationHandlesBufferBoundariesAndRestoresPositionOnCancellation(bool cancel) {
        using var bytes = new MemoryStream();
        bytes.WriteByte(0x7f); // The validated slice need not start at source position zero.
        using (var writer = new BinaryWriter(bytes, Encoding.ASCII, leaveOpen: true)) {
            for (int index = 0; index < 4; index++) {
                writer.Write((byte)(index == 3 ? 1 : 0));
                writer.Write(ushort.MaxValue); writer.Write((ushort)0);
                writer.Write(new byte[ushort.MaxValue]);
            }
        }
        using var cancellation = new CancellationTokenSource();
        using var source = new CancelOnReadStream(bytes.ToArray(), cancel ? cancellation : null) { Position = 1 };
        if (cancel) Assert.Throws<OperationCanceledException>(() => OfficeIMO.Core.Internal.OfficeDeflateStreamValidator.TryValidateExact(
            source, 4L * ushort.MaxValue, cancellation.Token));
        else Assert.True(OfficeIMO.Core.Internal.OfficeDeflateStreamValidator.TryValidateExact(source, 4L * ushort.MaxValue));
        Assert.Equal(1, source.Position);
    }

    private sealed class CancelOnReadStream(byte[] bytes, CancellationTokenSource? cancellation) : MemoryStream(bytes, writable: false) {
        public override int Read(byte[] buffer, int offset, int count) {
            int read = base.Read(buffer, offset, count);
            cancellation?.Cancel();
            return read;
        }
    }

    private static void WriteUInt32(byte[] bytes, int offset, uint value) {
        WriteUInt16(bytes, offset, (ushort)value); WriteUInt16(bytes, offset + 2, (ushort)(value >> 16));
    }

    [Theory]
    [InlineData(14)]
    [InlineData(18)]
    [InlineData(22)]
    public void NonDescriptorLocalCrcAndSizesMustMatchTheCentralRecord(int field) {
        using var fixture = new ArchiveFixtures();
        byte[] zip = Zip(("Example.admx", "A", CompressionLevel.Optimal));
        WriteUInt32(zip, field, BitConverter.ToUInt32(zip, field) ^ 1);
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(zip), fixture.Destination,
            new ArchiveExtractionOptions { IncludeFile = _ => false }));
        Assert.False(Directory.Exists(fixture.Destination));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ExplicitDirectoriesAreUniqueButMayFollowAnImplicitParent(bool duplicate) {
        using var fixture = new ArchiveFixtures();
        byte[] zip = duplicate ? Zip(("Templates/", null, CompressionLevel.Optimal), ("Templates/", null, CompressionLevel.Optimal))
            : Zip(("Templates/Example.admx", "A", CompressionLevel.Optimal), ("Templates/", null, CompressionLevel.Optimal));
        if (duplicate) Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(zip), fixture.Destination));
        else Assert.Single(ArchiveExtractor.Extract(fixture.Package(zip), fixture.Destination).Files);
    }

    [Fact]
    public void AClassicDirectoryInsideTheEndRecordCommentIsRejected() {
        using var fixture = new ArchiveFixtures();
        byte[] original = Zip(("Example.admx", "A", CompressionLevel.Optimal));
        int end = original.Length - 22;
        int central = (int)BitConverter.ToUInt32(original, end + 16);
        byte[] directory = original.Skip(central).Take(end - central).ToArray();
        byte[] zip = original.Take(central).Concat(original.Skip(end)).Concat(directory).ToArray();
        WriteUInt32(zip, central + 16, (uint)(central + 22));
        WriteUInt16(zip, central + 20, (ushort)directory.Length);
        Assert.False(OfficeIMO.Core.Internal.OfficeArchiveSafety.ScanZipCentralDirectory(zip, 10).IsValid);
        using var stream = new MemoryStream(zip, writable: false);
        Assert.False(OfficeIMO.Core.Internal.OfficeArchiveSafety.ScanZipCentralDirectory(stream, stream.Length, 10).IsValid);
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(zip), fixture.Destination));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void DescriptorMetadataMaySupplySizesButMustMatchTheCentralRecord(bool zip64, bool signature) {
        byte[] original = Zip(("Example.admx", "A", CompressionLevel.Optimal));
        int central = (int)BitConverter.ToUInt32(original, original.Length - 6);
        using var descriptor = new MemoryStream();
        using (var writer = new BinaryWriter(descriptor, Encoding.ASCII, leaveOpen: true)) {
            if (signature) writer.Write(0x08074b50U);
            writer.Write(BitConverter.ToUInt32(original, central + 16));
            if (zip64) { writer.Write((ulong)BitConverter.ToUInt32(original, central + 20)); writer.Write((ulong)BitConverter.ToUInt32(original, central + 24)); }
            else { writer.Write(BitConverter.ToUInt32(original, central + 20)); writer.Write(BitConverter.ToUInt32(original, central + 24)); }
        }
        int count = (int)descriptor.Length;
        byte[] zip = original.Take(central).Concat(descriptor.ToArray()).Concat(original.Skip(central)).ToArray();
        WriteUInt16(zip, 6, 8); WriteUInt16(zip, central + count + 8, 8);
        WriteUInt32(zip, 14, 0); WriteUInt32(zip, 18, 0); WriteUInt32(zip, 22, 0);
        WriteUInt32(zip, zip.Length - 6, (uint)(central + count));
        if (zip64) {
            int extraOffset = 30 + BitConverter.ToUInt16(zip, 26);
            zip = zip.Take(extraOffset).Concat(new byte[20]).Concat(zip.Skip(extraOffset)).ToArray();
            WriteUInt16(zip, 28, 20); WriteUInt16(zip, extraOffset, 1); WriteUInt16(zip, extraOffset + 2, 16);
            WriteUInt32(zip, extraOffset + 4, BitConverter.ToUInt32(original, central + 24));
            WriteUInt32(zip, extraOffset + 12, BitConverter.ToUInt32(original, central + 20));
            central += 20;
            WriteUInt32(zip, zip.Length - 6, (uint)(central + count));
        }
        using (var valid = new ArchiveFixtures()) Assert.Single(ArchiveExtractor.Extract(valid.Package(zip), valid.Destination).Files);
        zip[central + (signature ? 4 : 0)] ^= 1;
        using var invalid = new ArchiveFixtures();
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(invalid.Package(zip), invalid.Destination,
            new ArchiveExtractionOptions { IncludeFile = _ => false }));
    }

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void LocalZip64SizesAreResolvedInConditionalOrder(bool compressed, bool expanded) {
        byte[] original = Zip(("Example.admx", "A", CompressionLevel.Optimal));
        int extra = 30 + BitConverter.ToUInt16(original, 26);
        int extraLength = 4 + (compressed ? 8 : 0) + (expanded ? 8 : 0);
        byte[] zip = original.Take(extra).Concat(new byte[extraLength]).Concat(original.Skip(extra)).ToArray();
        WriteUInt16(zip, 28, (ushort)extraLength); WriteUInt16(zip, extra, 1); WriteUInt16(zip, extra + 2, (ushort)(extraLength - 4));
        int cursor = extra + 4;
        if (expanded) { WriteUInt32(zip, cursor, BitConverter.ToUInt32(original, 22)); cursor += 8; WriteUInt32(zip, 22, uint.MaxValue); }
        if (compressed) { WriteUInt32(zip, cursor, BitConverter.ToUInt32(original, 18)); WriteUInt32(zip, 18, uint.MaxValue); }
        WriteUInt32(zip, zip.Length - 6, BitConverter.ToUInt32(original, original.Length - 6) + (uint)extraLength);
        using (var valid = new ArchiveFixtures()) Assert.Single(ArchiveExtractor.Extract(valid.Package(zip), valid.Destination).Files);
        zip[extra + 4] ^= 1;
        using var invalid = new ArchiveFixtures();
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(invalid.Package(zip), invalid.Destination));
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
