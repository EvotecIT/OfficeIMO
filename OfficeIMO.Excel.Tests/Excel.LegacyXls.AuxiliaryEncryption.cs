using OfficeIMO.Core.Internal;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls.Biff;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelLegacyAuxiliaryEncryptionTests {
    private static string Fixture => Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", "ExcelPreservationCorpus", "excel-populated-macro.xls");

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void All_auxiliary_BIFF_streams_are_encrypted_and_survive_password_rotation(bool bookAlias, bool cryptoApi) {
        byte[] pivot = BuildAuxiliaryStream(false), revisions = BuildAuxiliaryStream(true);
        byte[] source = AddAuxiliaryStreams(pivot, revisions, bookAlias);
        if (cryptoApi) source = ExcelLegacyCryptoApiStorageTests.EncryptWithSummary(source, out _, out _);
        using var document = cryptoApi ? ExcelDocument.LoadEncrypted(new MemoryStream(source), "test-password") : ExcelDocument.Load(new MemoryStream(source));
        document.Sheets[0].CellValue(1, 1, "Encrypted auxiliary preservation");
        using var output = new MemoryStream(); document.SaveEncrypted(output, "test-password", ExcelFileFormat.Xls);
        Assert.True(OfficeCompoundFileReader.TryRead(output.ToArray(), out OfficeCompoundFile? encrypted, out string? error), error);
        Assert.True(OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? original, out error), error);
        string main = bookAlias ? "Book" : "Workbook";
        var filePass = Assert.Single(BiffRecordReader.ReadRecords(encrypted!.Streams[main], new()), record => record.Type == 0x002f);
        foreach (var pair in new[] { ("_SX_DB_CUR/0001", pivot), ("Revision Log", revisions) }) {
            byte[] ciphertext = encrypted.Streams[pair.Item1];
            Assert.NotEqual(pair.Item2, ciphertext);
            AssertClearRecordFields(pair.Item2, ciphertext);
            byte[] plaintext;
            Assert.True(cryptoApi
                ? BiffRc4CryptoApiEncryption.TryDecrypt(ciphertext, filePass, "test-password", new(), out plaintext)
                : BiffRc4Encryption.TryDecrypt(ciphertext, filePass, "test-password", new(), out plaintext));
            Assert.Equal(pair.Item2, plaintext);
        }
        foreach (var stream in original!.Streams.Where(item => item.Key != main && item.Key != "_SX_DB_CUR/0001" && item.Key != "Revision Log" && item.Key != "encryption"))
            Assert.Equal(stream.Value, encrypted.Streams[stream.Key]);
        using var loaded = ExcelDocument.LoadEncrypted(output, "test-password");
        Assert.Equal("Encrypted auxiliary preservation", loaded.Sheets[0].CellAt(1, 1).GetValue<string>());
        byte[] unencryptedBytes = loaded.ToBytes(ExcelFileFormat.Xls);
        Assert.True(OfficeCompoundFileReader.TryRead(unencryptedBytes, out OfficeCompoundFile? plaintextCompound, out error), error);
        Assert.Equal(pivot, plaintextCompound!.Streams["_SX_DB_CUR/0001"]);
        Assert.Equal(revisions, plaintextCompound.Streams["Revision Log"]);
        Assert.DoesNotContain(BiffRecordReader.ReadRecords(plaintextCompound.Streams[main], new()), record => record.Type == 0x002f);
        loaded.Sheets[0].CellValue(1, 1, "Rotated auxiliary preservation");
        using var rotated = new MemoryStream(); loaded.SaveEncrypted(rotated, "new-password", ExcelFileFormat.Xls);
        if (cryptoApi) {
            Assert.True(OfficeCompoundFileReader.TryRead(rotated.ToArray(), out OfficeCompoundFile? protectedOutput, out error), error);
            Assert.True(protectedOutput!.Streams.ContainsKey("encryption"));
            Assert.DoesNotContain(protectedOutput.Streams.Keys, name => name.StartsWith("ProofStorage/", StringComparison.OrdinalIgnoreCase));
        }
        using var reloaded = ExcelDocument.LoadEncrypted(rotated, "new-password");
        Assert.Equal("Rotated auxiliary preservation", reloaded.Sheets[0].CellAt(1, 1).GetValue<string>());
        Assert.Equal(4, reloaded.LegacyXlsChartSheets.Count);
        rotated.Position = 0; Assert.ThrowsAny<Exception>(() => ExcelDocument.LoadEncrypted(rotated, "test-password"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Malformed_auxiliary_records_are_rejected_before_destination_mutation(bool revision) {
        byte[] broken = { 0xcd, 0, 32, 0, 1, 2, 3 };
        byte[] valid = BuildAuxiliaryStream(!revision);
        using var document = ExcelDocument.Load(new MemoryStream(AddAuxiliaryStreams(revision ? valid : broken, revision ? broken : valid, false)));
        using var output = new MemoryStream(); output.WriteByte(42);
        Assert.Throws<InvalidDataException>(() => document.SaveEncrypted(output, "test-password", ExcelFileFormat.Xls,
            new ExcelSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.Equal(new byte[] { 42 }, output.ToArray());
    }

    [Theory]
    [InlineData(false, "bfea405495d6a27bef29d7d7bfb30040755c2e507b16763249efefa664631591")]
    [InlineData(true, "74d0be66def0ec41591dd9fa7519302d24c020e76c522f9e9169c92eaaef778a")]
    public void Auxiliary_encryption_matches_independent_ARC4_block_vectors(bool revision, string expectedHash) {
        // cryptography 50.0.2 ARC4, with independent RFC 6229 validation and fixed salt 00..0f.
        byte[] salt = Enumerable.Range(0, 16).Select(value => (byte)value).ToArray();
        byte[] encrypted = BiffRc4Encryption.EncryptWorkbookStream(BuildAuxiliaryStream(revision), "test-password", salt);
        using var hash = System.Security.Cryptography.SHA256.Create();
        Assert.Equal(expectedHash, BitConverter.ToString(hash.ComputeHash(encrypted)).Replace("-", string.Empty).ToLowerInvariant());
    }

    private static byte[] AddAuxiliaryStreams(byte[] pivot, byte[] revision, bool bookAlias) {
        Assert.True(OfficeCompoundFileReader.TryRead(File.ReadAllBytes(Fixture), out OfficeCompoundFile? source, out string? error), error);
        var replacements = new Dictionary<string, byte[]> { ["_SX_DB_CUR/0001"] = pivot, ["Revision Log"] = revision };
        if (bookAlias) replacements.Add("Book", source!.Streams["Workbook"]);
        return OfficeCompoundFileWriter.Rewrite(source!, replacements, removedPaths: bookAlias ? new[] { "Workbook" } : null);
    }

    private static byte[] BuildAuxiliaryStream(bool revision) {
        // Cipher-contract control: exempt records and payloads straddle several 1024-byte key blocks.
        using var stream = new MemoryStream();
        WriteRecord(stream, revision ? (ushort)BiffRecordType.RrdInfo : (ushort)0x00c6, new byte[20]);
        WriteRecord(stream, 0x00cd, Encoding.Unicode.GetBytes(string.Concat(Enumerable.Repeat("Łódź private cache value ", 100))));
        if (revision) WriteRecord(stream, (ushort)BiffRecordType.RrdHead, Enumerable.Range(0, 32).Select(value => (byte)value).ToArray());
        WriteRecord(stream, 0x00cd, Encoding.Unicode.GetBytes("Second auxiliary record"));
        WriteRecord(stream, 0x000a, Array.Empty<byte>());
        return stream.ToArray();
    }

    private static void WriteRecord(Stream stream, ushort type, byte[] payload) {
        stream.WriteByte((byte)type); stream.WriteByte((byte)(type >> 8));
        stream.WriteByte((byte)payload.Length); stream.WriteByte((byte)(payload.Length >> 8)); stream.Write(payload, 0, payload.Length);
    }

    private static void AssertClearRecordFields(byte[] plain, byte[] encrypted) {
        var before = BiffRecordReader.ReadRecords(plain, new()); var after = BiffRecordReader.ReadRecords(encrypted, new());
        Assert.Equal(before.Select(record => (record.Type, record.Offset, record.Payload.Length)), after.Select(record => (record.Type, record.Offset, record.Payload.Length)));
        for (int index = 0; index < before.Count; index++) {
            if (before[index].Type == (ushort)BiffRecordType.RrdInfo || before[index].Type == (ushort)BiffRecordType.RrdHead)
                Assert.Equal(before[index].Payload, after[index].Payload);
        }
    }
}
