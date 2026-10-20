using OfficeIMO.Core.Internal;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelLegacyNativeEncryptionTests {
    private static string Corpus => Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", "ExcelNativeEncryptionCorpus");

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Native_encrypted_embedding_retains_properties_and_survives_edit_and_password_rotation(bool removeClearMirror) {
        byte[] bytes = File.ReadAllBytes(Path.Combine(Corpus, "excel-native-embedding.encrypted.xls"));
        OfficeCompoundFile raw = ReadCompound(bytes);
        string storageName = EmbeddingName(raw);
        OfficeCompoundFile nativePlain = ReadCompound(File.ReadAllBytes(Path.Combine(Corpus, "excel-native-embedding.plain.xls")));
        byte[] expectedPackage = nativePlain.Streams[EmbeddingName(nativePlain) + "/Package"];
        Assert.Equal(expectedPackage, raw.Streams[storageName + "/Package"]);
        using (var embedded = ExcelDocument.Load(new MemoryStream(expectedPackage))) {
            Assert.Equal("Owned embedded native control", embedded.Sheets[0].CellAt(1, 1).GetValue<string>());
            Assert.Equal(42, embedded.Sheets[0].CellAt(1, 2).GetValue<double>());
        }
        if (removeClearMirror) bytes = OfficeCompoundFileWriter.Rewrite(raw, new Dictionary<string, byte[]>(), removedPaths: new[] { storageName });

        using var document = ExcelDocument.LoadEncrypted(new MemoryStream(bytes), "test-password");
        Assert.Equal("Native encrypted auxiliary control", document.BuiltinDocumentProperties.Title);
        Assert.Equal("Native encrypted nested storage", document.BuiltinDocumentProperties.Subject);
        Assert.Equal(50, document.Sheets[0].CellAt(3, 2).GetValue<double>());
        document.Sheets[0].CellValue(2, 1, "Edited native embedding");
        document.Sheets[0].CellFormula(1, 2, "40+2");
        OfficeCompoundFile plain = ReadCompound(document.ToBytes(ExcelFileFormat.Xls));
        Assert.Equal(expectedPackage, plain.Streams[storageName + "/Package"]);
        Assert.Equal(new Guid("00020830-0000-0000-c000-000000000046"),
            Assert.Single(plain.Entries, entry => entry.Path == storageName).ClassId);
        Assert.False(plain.Streams.ContainsKey("encryption"));
        var nativeBefore = LegacyXlsWorkbook.Load(File.ReadAllBytes(Path.Combine(Corpus, "excel-native-embedding.plain.xls")));
        var nativeAfter = LegacyXlsWorkbook.Load(document.ToBytes(ExcelFileFormat.Xls));
        Assert.Equal(nativeBefore.Worksheets[0].Cells.Single(cell => cell.Row == 1 && cell.Column == 2).StyleIndex,
            nativeAfter.Worksheets[0].Cells.Single(cell => cell.Row == 1 && cell.Column == 2).StyleIndex);
        Assert.Equal("40+2", nativeAfter.Worksheets[0].Cells.Single(cell => cell.Row == 1 && cell.Column == 2).FormulaText);

        using var rotated = new MemoryStream();
        document.SaveEncrypted(rotated, "new-password", ExcelFileFormat.Xls);
        using var reloaded = ExcelDocument.LoadEncrypted(rotated, "new-password");
        Assert.Equal("Edited native embedding", reloaded.Sheets[0].CellAt(2, 1).GetValue<string>());
        Assert.Equal("Native encrypted auxiliary control", reloaded.BuiltinDocumentProperties.Title);
        Assert.Equal(expectedPackage, ReadCompound(reloaded.ToBytes(ExcelFileFormat.Xls)).Streams[storageName + "/Package"]);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Conflicting_clear_embedding_copies_are_rejected(bool changeClassIdentity) {
        OfficeCompoundFile raw = ReadCompound(File.ReadAllBytes(Path.Combine(Corpus, "excel-native-embedding.encrypted.xls")));
        string name = EmbeddingName(raw);
        var replacements = new Dictionary<string, byte[]>();
        if (changeClassIdentity) {
            var entries = raw.Entries.Select(entry => entry.Path == name
                ? new OfficeCompoundFileEntry(entry.Name, entry.Path, entry.ObjectType, entry.Size,
                    classId: new Guid("00020906-0000-0000-c000-000000000046"), stateBits: entry.StateBits,
                    creationTime: entry.CreationTime, modifiedTime: entry.ModifiedTime)
                : entry).ToArray();
            raw = new OfficeCompoundFile(raw.Streams, entries, raw.RootEntry);
        } else {
            byte[] changed = (byte[])raw.Streams[name + "/Package"].Clone();
            changed[changed.Length / 2] ^= 1;
            replacements.Add(name + "/Package", changed);
        }
        byte[] bytes = OfficeCompoundFileWriter.Rewrite(raw, replacements);
        Assert.Throws<InvalidDataException>(() => ExcelDocument.LoadEncrypted(new MemoryStream(bytes), "test-password"));
    }

    [Fact]
    public void Property_encryption_flag_requires_encrypted_property_descriptors() {
        OfficeCompoundFile raw = ReadCompound(File.ReadAllBytes(Path.Combine(Corpus, "excel-native-embedding.encrypted.xls")));
        byte[] workbook = (byte[])raw.Streams["Workbook"].Clone();
        BiffRecord filePass = Assert.Single(BiffRecordReader.ReadRecords(workbook, new()), record => record.Type == 0x002f);
        // FilePass's prefix precedes the outer and inner CryptoAPI header flags.
        Assert.Equal(0x0c, filePass.Payload[6]);
        Assert.Equal(0x0c, filePass.Payload[14]);
        workbook[filePass.Offset + 4 + 6] &= 0xf7;
        workbook[filePass.Offset + 4 + 14] &= 0xf7;
        byte[] bytes = OfficeCompoundFileWriter.Rewrite(raw, new Dictionary<string, byte[]> { ["Workbook"] = workbook });
        Assert.Throws<InvalidDataException>(() => ExcelDocument.LoadEncrypted(new MemoryStream(bytes), "test-password"));
    }

    [Theory]
    [InlineData("Workbook")]
    [InlineData("Book")]
    [InlineData("encryption")]
    public void Encrypted_storages_cannot_shadow_reserved_workbook_entries(string name) {
        OfficeCompoundFile raw = ReadCompound(File.ReadAllBytes(Path.Combine(Corpus, "excel-native-embedding.encrypted.xls")));
        BiffRecord filePass = Assert.Single(BiffRecordReader.ReadRecords(raw.Streams["Workbook"], new()), record => record.Type == 0x002f);
        var session = BiffRc4CryptoApiEncryption.OpenSession(filePass, "test-password");
        var storage = Assert.Single(OfficeBinaryEncryptedSummary.Decrypt(raw.Streams["encryption"], session, 1024 * 1024));
        byte[] summary = ExcelLegacyCryptoApiStorageTests.EncodeSummary(new[] { (name, false, storage.Bytes) }, session);
        byte[] bytes = OfficeCompoundFileWriter.Rewrite(raw, new Dictionary<string, byte[]> { ["encryption"] = summary });
        Assert.Throws<InvalidDataException>(() => ExcelDocument.LoadEncrypted(new MemoryStream(bytes), "test-password"));
    }

    private static string EmbeddingName(OfficeCompoundFile compound) =>
        Assert.Single(compound.Entries, entry => !entry.IsFallback && entry.ObjectType == 1
            && entry.Path.StartsWith("MBD", StringComparison.Ordinal) && entry.Path.IndexOf('/') < 0).Path;

    private static OfficeCompoundFile ReadCompound(byte[] bytes) {
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out string? error), error);
        return compound!;
    }
}
