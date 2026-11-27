namespace OfficeIMO.Archives.Tests;

public sealed class ZipEncryptionTests {
    private const string Password = "correct horse";
    private static readonly string[] Members = { "text.txt", "mixed.bin", "tiny.txt", "empty.bin" };

    [Theory]
    [InlineData("zipcrypto-infozip.zip")]
    [InlineData("zipcrypto-infozip-stored.zip")]
    [InlineData("zipcrypto-7zip.zip")]
    [InlineData("aes128-7zip.zip")]
    [InlineData("aes192-7zip.zip")]
    [InlineData("aes256-7zip.zip")]
    [InlineData("aes256-7zip-stored.zip")]
    public void ReadsEveryMemberWrittenByTheReferenceTools(string fixture) {
        IReadOnlyList<ArchiveEntryData> entries = ArchiveReader.ReadZip(ReadFixture(fixture), new ArchiveReadOptions { ZipPassword = Password });
        Assert.Equal(Members.OrderBy(name => name, StringComparer.Ordinal), entries.Select(entry => entry.RelativePath).OrderBy(name => name, StringComparer.Ordinal));
        foreach (ArchiveEntryData entry in entries) Assert.Equal(ReadInput(entry.RelativePath), entry.Data);
    }

    [Fact]
    public void ZipCryptoWithADataDescriptorChecksTheTimeByte() {
        // Info-ZIP reading standard input writes a descriptor and puts the time's high byte in the encryption header.
        ArchiveEntryData entry = Assert.Single(ArchiveReader.ReadZip(ReadFixture("zipcrypto-infozip-descriptor.zip"),
            new ArchiveReadOptions { ZipPassword = Password }));
        Assert.Equal(ReadInput("text.txt"), entry.Data);
    }

    [Theory]
    [InlineData("zipcrypto-infozip.zip")]
    [InlineData("aes256-7zip.zip")]
    public void MissingAndWrongPasswordsRaiseThePasswordError(string fixture) {
        ArchivePasswordException missing = Assert.Throws<ArchivePasswordException>(() => ArchiveReader.ReadZip(ReadFixture(fixture)));
        Assert.False(missing.PasswordSupplied);

        ArchivePasswordException wrong = Assert.Throws<ArchivePasswordException>(() =>
            ArchiveReader.ReadZip(ReadFixture(fixture), new ArchiveReadOptions { ZipPassword = "battery staple" }));
        Assert.True(wrong.PasswordSupplied);
    }

    [Fact]
    public void InspectionListsEncryptedMembersAndOpensThemOnlyWithThePassword() {
        using (ArchiveInspection locked = ArchiveInspector.Open(new MemoryStream(ReadFixture("aes128-7zip.zip")))) {
            ArchiveInspectionEntry text = locked.Entries.Single(entry => entry.RawName == "text.txt");
            Assert.Equal("ZIP_PASSWORD", text.Diagnostic?.Code);
            Assert.Equal(ArchiveCompression.Deflate, text.Compression);
            Assert.Throws<ArchivePasswordException>(() => locked.OpenRead(text.Ordinal));
            Assert.Throws<ArchivePasswordException>(() => locked.ReadMembers(new[] { text.Ordinal }, long.MaxValue));
        }

        using ArchiveInspection open = ArchiveInspector.Open(new MemoryStream(ReadFixture("aes128-7zip.zip")),
            new ArchiveInspectionOptions { ZipPassword = Password });
        Assert.All(open.Entries, entry => Assert.Null(entry.Diagnostic));
        Assert.All(open.Verify().Members, member => Assert.Equal(ArchiveVerificationStatus.Verified, member.Status));
        ArchiveInspectionEntry mixed = open.Entries.Single(entry => entry.RawName == "mixed.bin");
        using Stream content = open.OpenRead(mixed.Ordinal);
        var copy = new MemoryStream();
        content.CopyTo(copy);
        Assert.Equal(ReadInput("mixed.bin"), copy.ToArray());
    }

    [Fact]
    public void ExtractionDecryptsToTheDestination() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-zip-encryption-" + Guid.NewGuid().ToString("N"));
        try {
            string package = Path.Combine(root, "package.zip");
            Directory.CreateDirectory(root);
            File.WriteAllBytes(package, ReadFixture("zipcrypto-7zip.zip"));
            ArchiveExtractor.Extract(package, Path.Combine(root, "out"), new ArchiveExtractionOptions { ZipPassword = Password });
            foreach (string member in Members) Assert.Equal(ReadInput(member), File.ReadAllBytes(Path.Combine(root, "out", member)));
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void ChangedAesCiphertextFailsAuthenticationAtTheCode() {
        byte[] archive = ReadFixture("aes256-7zip.zip");
        ZipMember member = FindMember(archive, "mixed.bin");
        archive[member.Payload + 16 + 2 + 100] ^= 0x01;
        ArchiveCorruptionException error = Assert.Throws<ArchiveCorruptionException>(() =>
            ArchiveReader.ReadZip(archive, new ArchiveReadOptions { ZipPassword = Password }));
        Assert.Equal("ZIP", error.Format);
        Assert.Equal(member.Payload + member.CompressedSize - 10, error.Offset);
    }

    [Fact]
    public void AeVersionOneMembersAlsoCheckTheCrc() {
        // 7-Zip writes AE-2 only. AE-1 differs in the extra field's version and in storing the real CRC.
        byte[] archive = ReadFixture("aes256-7zip.zip");
        ZipMember member = FindMember(archive, "text.txt");
        uint crc = Crc32(ReadInput("text.txt"));
        SetAesVersion(archive, member.Local + 30 + member.NameLength, ReadUInt16(archive, member.Local + 28), 1);
        SetAesVersion(archive, member.Central + 46 + member.NameLength, ReadUInt16(archive, member.Central + 30), 1);
        WriteUInt32(archive, member.Local + 14, crc);
        WriteUInt32(archive, member.Central + 16, crc);
        var options = new ArchiveReadOptions { ZipPassword = Password };
        Assert.Equal(ReadInput("text.txt"), ArchiveReader.ReadZip(archive, options).Single(entry => entry.RelativePath == "text.txt").Data);

        WriteUInt32(archive, member.Local + 14, crc ^ 1);
        WriteUInt32(archive, member.Central + 16, crc ^ 1);
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => ArchiveReader.ReadZip(archive, options));
        Assert.Contains("checksum", error.Message);
    }

    [Fact]
    public void StrongEncryptionIsReportedAsUnsupported() {
        byte[] archive = ReadFixture("zipcrypto-7zip.zip");
        ZipMember member = FindMember(archive, "text.txt");
        archive[member.Central + 8] |= 0x40;
        archive[member.Local + 6] |= 0x40;
        using ArchiveInspection inspection = ArchiveInspector.Open(new MemoryStream(archive), new ArchiveInspectionOptions { ZipPassword = Password });
        Assert.Equal("ZIP_ENCRYPTION", inspection.Entries.Single(entry => entry.RawName == "text.txt").Diagnostic?.Code);
        Assert.Throws<NotSupportedException>(() => ArchiveReader.ReadZip(archive, new ArchiveReadOptions { ZipPassword = Password }));
    }

    [Theory]
    [InlineData("zipcrypto-7zip.zip")]
    [InlineData("zipcrypto-infozip.zip")]
    [InlineData("aes128-7zip.zip")]
    public void FlippedBitsFailWithTypedErrorsOnly(string fixture) {
        byte[] original = ReadFixture(fixture);
        var random = new Random(5);
        var options = new ArchiveReadOptions { ZipPassword = Password };
        for (int trial = 0; trial < 300; trial++) {
            byte[] damaged = (byte[])original.Clone();
            int position = random.Next(damaged.Length);
            damaged[position] ^= (byte)(1 << random.Next(8));
            IReadOnlyList<ArchiveEntryData>? entries = null;
            Exception? error = Record.Exception(() => entries = ArchiveReader.ReadZip(damaged, options));
            if (error == null) {
                // Bits no reader checks (timestamps, attributes, comments) leave the decoded members unchanged.
                foreach (ArchiveEntryData entry in entries!) Assert.Equal(ReadInput(entry.RelativePath), entry.Data);
                continue;
            }
            Assert.True(error is ArchivePasswordException or ArchiveCorruptionException or InvalidDataException
                    or NotSupportedException or OfficeInspectionResourceLimitException,
                $"Flipping a bit at byte {position} raised {error.GetType().Name}: {error.Message}");
        }
    }

    private readonly struct ZipMember {
        internal ZipMember(int central, int local, int nameLength, long payload, long compressedSize) {
            Central = central; Local = local; NameLength = nameLength; Payload = payload; CompressedSize = compressedSize;
        }
        internal int Central { get; }
        internal int Local { get; }
        internal int NameLength { get; }
        internal long Payload { get; }
        internal long CompressedSize { get; }
    }

    private static ZipMember FindMember(byte[] archive, string name) {
        int end = archive.Length - 22;
        while (ReadUInt32(archive, end) != 0x06054b50) end--;
        int cursor = (int)ReadUInt32(archive, end + 16);
        for (int index = 0; index < ReadUInt16(archive, end + 10); index++) {
            int nameLength = ReadUInt16(archive, cursor + 28);
            if (Encoding.UTF8.GetString(archive, cursor + 46, nameLength) == name) {
                int local = (int)ReadUInt32(archive, cursor + 42);
                long payload = local + 30L + ReadUInt16(archive, local + 26) + ReadUInt16(archive, local + 28);
                return new ZipMember(cursor, local, nameLength, payload, ReadUInt32(archive, cursor + 20));
            }
            cursor += 46 + nameLength + ReadUInt16(archive, cursor + 30) + ReadUInt16(archive, cursor + 32);
        }
        throw new InvalidOperationException("Member not found: " + name);
    }

    private static void SetAesVersion(byte[] archive, int extra, int length, ushort version) {
        for (int offset = extra; offset < extra + length;) {
            if (ReadUInt16(archive, offset) == 0x9901) {
                archive[offset + 4] = (byte)version;
                archive[offset + 5] = (byte)(version >> 8);
                return;
            }
            offset += 4 + ReadUInt16(archive, offset + 2);
        }
        throw new InvalidOperationException("No WinZip AES extra field.");
    }

    private static uint Crc32(byte[] data) {
        uint crc = 0xFFFFFFFF;
        foreach (byte value in data) {
            crc ^= value;
            for (int bit = 0; bit < 8; bit++) crc = (crc & 1) != 0 ? 0xEDB88320 ^ (crc >> 1) : crc >> 1;
        }
        return ~crc;
    }

    private static ushort ReadUInt16(byte[] data, int offset) => (ushort)(data[offset] | data[offset + 1] << 8);
    private static uint ReadUInt32(byte[] data, int offset) => (uint)(data[offset] | data[offset + 1] << 8 | data[offset + 2] << 16 | data[offset + 3] << 24);
    private static void WriteUInt32(byte[] data, int offset, uint value) {
        for (int index = 0; index < 4; index++) data[offset + index] = (byte)(value >> (8 * index));
    }

    private static byte[] ReadFixture(string name) =>
        File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "IndependentZipEncryption", name));

    private static byte[] ReadInput(string name) => name == "-"
        ? ReadInput("text.txt")
        : File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "IndependentZipEncryption", "inputs", name));
}
