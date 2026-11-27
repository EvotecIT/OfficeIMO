using OfficeIMO.Archives.Rar5;
using OfficeIMO.Drawing;

namespace OfficeIMO.Archives.Tests;

public sealed partial class Rar5ArchiveTests {
    [Theory]
    [InlineData("abc", "508C5E8C327C14E2E1A72BA34EEB452F37458B209ED63A294D999B4C86675982")]
    [InlineData("", "69217A3079908094E11121D042354A7C1F55B6482CA1A51E1B250DFD1ED0EEF9")]
    public void Blake2sMatchesPublishedAndEmptyMessageVectors(string message, string digest) =>
        Assert.Equal(digest, BitConverter.ToString(Blake2s.Digest(Encoding.ASCII.GetBytes(message))).Replace("-", ""));

    [Theory]
    [InlineData("prefixed.sfx")]
    [InlineData("zero-prefixed.sfx")]
    public void ExtractsRar5BehindTheProducerSfxPrefix(string name) {
        using var files = new ArchiveFixtures();
        var result = ArchiveExtractor.Extract(files.Package(Fixture(name)), files.Destination);
        Assert.Equal(ArchiveFormat.Rar5, result.Format);
        Assert.Equal("notes/señal.txt", Assert.Single(result.Files).RelativePath);
    }

    [Fact]
    public void BoundsTheSfxSignatureSearch() {
        byte[] source = Fixture("prefixed.sfx");
        const int bound = 1024 * 1024;
        const int stubSize = 31;
        byte[] padded = new byte[bound + 1 - stubSize + source.Length];
        Buffer.BlockCopy(source, 0, padded, bound - stubSize, source.Length);
        Assert.Equal("notes/señal.txt", Assert.Single(ArchiveReader.ReadRar5(padded)).RelativePath);
        Buffer.BlockCopy(source, 0, padded, bound + 1 - stubSize, source.Length);
        Assert.Throws<NotSupportedException>(() => ArchiveReader.ReadRar5(padded));
    }

    [Theory]
    [InlineData("CRC", "header checksum")]
    [InlineData("header size", "header has an invalid size")]
    [InlineData("vint", "variable integer exceeds")]
    [InlineData("data size", "data or extra area")]
    public void CorruptPrimaryHeaderNamesItsByteOffset(string fault, string diagnostic) {
        byte[] source = Fixture("plain.rar");
        const int member = 16;
        int payload = member + 5 + source[member + 4];
        switch (fault) {
            case "CRC":
                source[member] ^= 1;
                break;
            case "header size":
                for (int i = member + 4; i < member + 7; i++) source[i] = 0x80;
                break;
            case "vint":
                for (int i = member + 7; i < member + 17; i++) source[i] = 0x80;
                RefreshCrc(source, member, payload);
                break;
            case "data size":
                source[member + 8] = 3;
                RefreshCrc(source, member, payload);
                break;
        }
        using var stream = new MemoryStream(source, writable: false);
        var error = Assert.Throws<InvalidDataException>(() => ArchiveReader.ListRar5(stream));
        Assert.Contains("byte offset 16", error.Message, StringComparison.Ordinal);
        Assert.Contains(diagnostic, error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ReadsAuthoritativeNamesEvenWhenQuickOpenCopiesDisagreeAndRejectsTraversal() {
        // Edit the primary header and leave its producer-authored Quick Open cache intact.
        byte[] stored = Fixture("stored.rar");
        byte[] old = Encoding.UTF8.GetBytes("notes/señal.txt");
        int start = Find(stored, old);
        int another = Find(stored, old, start + old.Length);
        Assert.True(another > start);
        ReplaceNameAndRefreshHeader(stored, start, old, "notes/sonal.txtq");
        using var stream = new MemoryStream(stored, writable: false);
        var names = ArchiveReader.ListRar5(stream).Select(entry => entry.Name).ToArray();
        Assert.Contains("notes/sonal.txtq", names);
        Assert.DoesNotContain("notes/señal.txt", names);
        var options = new ArchiveReadOptions { IncludeFile = path => path == "notes/sonal.txtq" };
        Assert.Equal("notes/sonal.txtq", Assert.Single(ArchiveReader.ReadRar5(stored, options)).RelativePath);

        byte[] unsafePath = Fixture("plain.rar");
        ReplaceNameAndRefreshHeader(unsafePath, Find(unsafePath, old), old, "../bad/file.txtq");
        using var files = new ArchiveFixtures();
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(files.Package(unsafePath), files.Destination));
        Assert.False(Directory.Exists(files.Destination));
    }

    [Fact]
    public void ReadsRarCommentOwnerTimesAndRedirectionsWithoutFollowingFilesystemLinks() {
        byte[] source = Fixture("metadata.rar");
        using var input = new MemoryStream(source, writable: false);
        var archive = ArchiveReader.InspectRar5(input);
        Assert.Equal("RAR comment café\n", archive.Comment);
        Assert.Equal(new[] { "piece.bin", "linked.bin", "copy.bin", "directory-link", "empty-directory" }, archive.Entries.Select(entry => entry.Name));
        var original = archive.Entries[0];
        Assert.NotEmpty(Assert.IsType<Rar5UnixOwner>(original.UnixOwner).UserName!);
        Assert.NotEmpty(original.UnixOwner!.GroupName!);
        Assert.Equal(DateTimeOffset.FromUnixTimeSeconds(2200000002).AddTicks(1234567), original.LastWriteTime);
        Assert.Equal(123456700, original.Times.Modified!.Value.Nanoseconds);
        Assert.NotNull(original.Times.Created);
        Assert.NotNull(original.Times.Accessed);
        Assert.Equal(Rar5RedirectionKind.HardLink, archive.Entries[1].Redirection!.Kind);
        Assert.Equal("piece.bin", archive.Entries[1].Redirection!.Target);
        Assert.Equal(Rar5RedirectionKind.FileCopy, archive.Entries[2].Redirection!.Kind);
        Assert.Equal(Rar5RedirectionKind.UnixSymbolicLink, archive.Entries[3].Redirection!.Kind);
        Assert.True(archive.Entries[3].Redirection!.TargetIsDirectory);
        Assert.Equal("empty-directory", archive.Entries[3].Redirection!.Target);
        Assert.Equal(ArchiveEntryKind.Directory, archive.Entries[4].Kind);

        var readCopies = new ArchiveReadOptions { IncludeFile = name => name == "copy.bin" };
        Assert.Equal(Piece(), Assert.Single(ArchiveReader.ReadRar5(source, readCopies)).Data);
        byte[] corruptCopy = (byte[])source.Clone();
        int copyName = Find(corruptCopy, Encoding.ASCII.GetBytes("copy.bin"));
        byte[] emptyDigest = Blake2s.TreeDigest(Array.Empty<byte>(), 0, 0, CancellationToken.None);
        int digestOffset = Find(corruptCopy, emptyDigest, copyName + "copy.bin".Length);
        ReplaceAndRefreshHeader(corruptCopy, digestOffset, new[] { (byte)(emptyDigest[0] ^ 1) });
        var copyError = Assert.Throws<InvalidDataException>(() => ArchiveReader.ReadRar5(corruptCopy, readCopies));
        Assert.Contains("copy.bin", copyError.Message, StringComparison.Ordinal);
        Assert.Contains("BLAKE2sp", copyError.Message, StringComparison.Ordinal);

        byte[] data = Piece();
        source[Find(source, data)] ^= 1;
        var error = Assert.Throws<InvalidDataException>(() => ArchiveReader.ReadRar5(source, readCopies));
        Assert.Contains("BLAKE2sp", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ListsWindowsSymlinksAndJunctionsAndNeverPublishesThem() {
        byte[] data = Fixture("windows-links.rar");
        using var stream = new MemoryStream(data, writable: false);
        var entries = ArchiveReader.ListRar5(stream);
        var link = Assert.Single(entries, entry => entry.Name == "file-link");
        Assert.Equal(ArchiveEntryKind.SymbolicLink, link.Kind);
        Assert.Equal(Rar5RedirectionKind.WindowsSymbolicLink, link.Redirection!.Kind);
        Assert.Equal("directory/page.txt", link.Redirection!.Target);
        Assert.False(link.CanRead);
        var junction = Assert.Single(entries, entry => entry.Name == "junction");
        Assert.Equal(Rar5RedirectionKind.WindowsJunction, junction.Redirection!.Kind);
        Assert.True(junction.Redirection!.TargetIsDirectory);
        Assert.EndsWith("/directory", junction.Redirection!.Target, StringComparison.Ordinal);
        Assert.False(junction.CanRead);

        using var files = new ArchiveFixtures();
        Assert.Throws<NotSupportedException>(() => ArchiveExtractor.Extract(files.Package(data), files.Destination));
        Assert.False(Directory.Exists(files.Destination));
    }

    [Fact]
    public void ReadsNumericUnixOwnersWhenNamesAreAbsent() {
        using var source = new MemoryStream(Fixture("numeric-owner.rar"), writable: false);
        var owner = Assert.IsType<Rar5UnixOwner>(Assert.Single(ArchiveReader.ListRar5(source)).UnixOwner);
        Assert.Equal(42421UL, owner.UserId);
        Assert.Equal(42422UL, owner.GroupId);
        Assert.Null(owner.UserName);
        Assert.Null(owner.GroupName);
    }

    private static byte[] Piece() => Enumerable.Range(0, 251 * 7).Select(index => (byte)(index % 251)).ToArray();

    private static int Find(byte[] source, byte[] bytes, int start = 0) {
        for (int i = start; i <= source.Length - bytes.Length; i++) {
            if (source.Skip(i).Take(bytes.Length).SequenceEqual(bytes)) return i;
        }
        return -1;
    }

    private static void ReplaceNameAndRefreshHeader(byte[] source, int start, byte[] old, string replacement) {
        byte[] name = Encoding.UTF8.GetBytes(replacement);
        Assert.Equal(old.Length, name.Length);
        ReplaceAndRefreshHeader(source, start, name);
    }

    private static void ReplaceAndRefreshHeader(byte[] source, int start, byte[] replacement) {
        Assert.True(start >= 0);
        // Find the enclosing original, CRC-valid primary header; no location from QO is used.
        for (int offset = 8; offset < start; offset++) {
            int end = offset + 5 + source[offset + 4];
            if (end < start + replacement.Length || end > source.Length || (source[offset + 4] & 128) != 0) continue;
            uint expected = OfficePngCrc32.Complete(OfficePngCrc32.Append(OfficePngCrc32.Begin(), source, offset + 4, end - offset - 4));
            if (Rar5HeaderReader.UInt32At(source, offset) != expected) continue;
            Buffer.BlockCopy(replacement, 0, source, start, replacement.Length);
            RefreshCrc(source, offset, end);
            return;
        }
        throw new InvalidOperationException("Cannot locate the producer's primary header.");
    }

    private static void RefreshCrc(byte[] source, int offset, int end) {
        uint crc = OfficePngCrc32.Complete(OfficePngCrc32.Append(OfficePngCrc32.Begin(), source, offset + 4, end - offset - 4));
        for (int i = 0; i < 4; i++) source[offset + i] = (byte)(crc >> (8 * i));
    }
}
