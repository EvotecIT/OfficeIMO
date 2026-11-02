using OfficeIMO.Core.Internal;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Model;
using OfficeIMO.Excel.LegacyXls.Write;
using System.Text;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelLegacyCryptoApiStorageTests {
    [Fact]
    public void Encrypted_summary_restores_nested_storage_metadata_and_supports_plaintext_save_and_password_rotation() {
        string path = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", "ExcelPreservationCorpus", "excel-populated-macro.xls");
        byte[] source = File.ReadAllBytes(path);
        byte[] encrypted = EncryptWithSummary(source, out OfficeCompoundFile original, out Guid storageClassId);
        Assert.True(OfficeCompoundFileReader.TryRead(encrypted, out OfficeCompoundFile? encryptedCompound, out string? encryptedError), encryptedError);
        foreach (var stream in original.Streams.Where(item => item.Key.StartsWith("_VBA_PROJECT_CUR/", StringComparison.OrdinalIgnoreCase)))
            Assert.Equal(stream.Value, encryptedCompound!.Streams[stream.Key]);
        using var document = ExcelDocument.LoadEncrypted(new MemoryStream(encrypted), "test-password");
        document.Sheets[0].CellValue(1, 1, "CryptoAPI restored storage");
        byte[] plain = document.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xls });
        Assert.True(OfficeCompoundFileReader.TryRead(plain, out OfficeCompoundFile? result, out string? error), error);
        Assert.False(result!.Streams.ContainsKey("encryption"));
        Assert.DoesNotContain(BiffRecordReader.ReadRecords(result.Streams["Workbook"], new()), item => item.Type == 0x002f);
        foreach (var stream in original.Streams.Where(item => item.Key.StartsWith("_VBA_PROJECT_CUR/", StringComparison.OrdinalIgnoreCase)))
            Assert.Equal(stream.Value, result.Streams[stream.Key]);
        var storage = Assert.Single(result.Entries, entry => entry.Path == "ProofStorage");
        Assert.Equal(storageClassId, storage.ClassId);
        Assert.Equal(0x12345678U, storage.StateBits);
        Assert.Equal(Encoding.UTF8.GetBytes("Inert encrypted nested payload"), result.Streams["ProofStorage/Nested/Contents"]);
        using var output = new MemoryStream(); document.SaveEncrypted(output, "new-password", ExcelFileFormat.Xls);
        Assert.True(OfficeCompoundFileReader.TryRead(output.ToArray(), out OfficeCompoundFile? protectedOutput, out error), error);
        Assert.DoesNotContain(protectedOutput!.Streams.Keys, name => name.StartsWith("ProofStorage/", StringComparison.OrdinalIgnoreCase));
        Assert.False(protectedOutput.Streams.ContainsKey("\u0005SummaryInformation"));
        var descriptor = Assert.Single(BiffRecordReader.ReadRecords(protectedOutput.Streams["Workbook"], new()), record => record.Type == 0x002f);
        var session = BiffRc4CryptoApiEncryption.OpenSession(descriptor, "new-password");
        Assert.True(session.EncryptsDocumentProperties);
        Assert.Equal(3, OfficeBinaryEncryptedSummary.Decrypt(protectedOutput.Streams["encryption"], session, 1024 * 1024).Count);
        foreach (var stream in original.Streams.Where(item => item.Key.StartsWith("_VBA_PROJECT_CUR/", StringComparison.OrdinalIgnoreCase)))
            Assert.Equal(stream.Value, protectedOutput.Streams[stream.Key]);
        using var reloaded = ExcelDocument.LoadEncrypted(output, "new-password");
        Assert.Equal("CryptoAPI restored storage", reloaded.Sheets[0].CellAt(1, 1).GetValue<string>());
        Assert.Equal(4, reloaded.LegacyXlsChartSheets.Count);
    }

    [Fact]
    public void Shared_summary_writer_matches_independent_descriptor_encoding_and_honors_bounds_and_cancellation() {
        var session = OfficeBinaryRc4CryptoApiSession.Create("test-password", 128, true, out _);
        var entries = new[] { ("\u0005SummaryInformation", true, Encoding.UTF8.GetBytes("Property vector")),
            ("Załącznik", false, Encoding.UTF8.GetBytes("Opaque storage vector")) };
        byte[] expected = EncodeSummary(entries, session);
        var input = entries.Select(entry => new OfficeBinaryEncryptedSummary.Entry(entry.Item1, entry.Item2, entry.Item3)).ToArray();
        Assert.Equal(expected, OfficeBinaryEncryptedSummary.Encrypt(input, session, expected.Length));
        Assert.Throws<InvalidDataException>(() => OfficeBinaryEncryptedSummary.Encrypt(input, session, expected.Length - 1));
        Assert.Throws<OperationCanceledException>(() => OfficeBinaryEncryptedSummary.Encrypt(input, session,
            cancellationToken: new CancellationToken(canceled: true)));
        Assert.Equal("Property vector", Encoding.UTF8.GetString(entries[0].Item3));
        Assert.Equal("Opaque storage vector", Encoding.UTF8.GetString(entries[1].Item3));
    }

    [Fact]
    public void Truncated_encrypted_summary_is_rejected_instead_of_losing_hidden_content() {
        string path = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", "ExcelPreservationCorpus", "excel-populated-macro.xls");
        byte[] encrypted = EncryptWithSummary(File.ReadAllBytes(path), out _, out _);
        Assert.True(OfficeCompoundFileReader.TryRead(encrypted, out OfficeCompoundFile? compound, out string? error), error);
        byte[] damaged = OfficeCompoundFileWriter.Rewrite(compound!, new Dictionary<string, byte[]> { ["encryption"] = new byte[7] });
        Assert.Throws<InvalidDataException>(() => ExcelDocument.LoadEncrypted(new MemoryStream(damaged), "test-password"));
    }

    [Theory]
    [InlineData("officeimo-cryptoapi-summary.xls")]
    [InlineData("excel-cryptoapi-summary-readback.xls")]
    public void Native_consumer_readback_retains_encrypted_property_values_and_inert_VBA(string filename) {
        string path = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", "ExcelPreservationCorpus", filename);
        byte[] bytes = File.ReadAllBytes(path);
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? raw, out string? error), error);
        Assert.All(raw!.Streams.Keys.Where(name => name.StartsWith("_VBA_PROJECT_CUR/", StringComparison.OrdinalIgnoreCase)),
            name => Assert.NotEmpty(raw.Streams[name]));
        if (filename.StartsWith("officeimo", StringComparison.Ordinal)) {
            Assert.True(raw.Streams.ContainsKey("encryption"));
            Assert.False(raw.Streams.ContainsKey("\u0005SummaryInformation"));
        }
        using var document = ExcelDocument.LoadEncrypted(new MemoryStream(bytes), "test-password");
        Assert.Equal("Encrypted summary native control", document.BuiltinDocumentProperties.Title);
        Assert.Equal("Nested inert VBA storage acceptance", document.BuiltinDocumentProperties.Subject);
        Assert.Equal(10, document.Sheets.Count); Assert.Equal(4, document.LegacyXlsChartSheets.Count);
        var native = LegacyXlsWorkbook.Load(bytes, new() { Password = "test-password" });
        var projectStorage = OfficeCompoundStorage.Extract(native.SourceCompoundFile!, "_VBA_PROJECT_CUR");
        var project = OfficeIMO.OfficeVbaProject.Load(OfficeCompoundFileWriter.Rewrite(projectStorage, new Dictionary<string, byte[]>()));
        Assert.Contains(project.Modules, module => module.Name == "PreservationNotes");
        Assert.Contains(project.Modules, module => module.Name == "DataCode1");
    }

    internal static byte[] EncryptWithSummary(byte[] bytes, out OfficeCompoundFile original, out Guid storageClassId) {
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out string? error), error);
        original = compound!;
        var session = OfficeBinaryRc4CryptoApiSession.Create("test-password", 128, true, out byte[] header);
        var filePass = new byte[header.Length + 2]; filePass[0] = 1;
        Buffer.BlockCopy(header, 0, filePass, 2, header.Length);
        using var input = new MemoryStream(bytes);
        var model = LegacyXlsWorkbook.Load(input);
        byte[] prepared = new LegacyXlsPreservingWorkbook(model, CancellationToken.None)
            .Write(Array.Empty<LegacyXlsLiteralEdit>(), filePass, int.MaxValue, CancellationToken.None);
        Assert.True(OfficeCompoundFileReader.TryRead(prepared, out OfficeCompoundFile? withDescriptor, out error), error);
        // VBA stays in the clear under MS-XLS 2.2.10. Exercise the generic
        // encrypted-summary storage container with inert opaque test data.
        const string storageName = "ProofStorage";
        storageClassId = new Guid("6842D9E3-032D-448D-B646-BDA84BDE218A");
        var storageEntries = new List<OfficeCompoundFileEntry> {
            new("Nested", "Nested", 1, 0), new("Contents", "Nested/Contents", 2, 30)
        };
        var root = new OfficeCompoundFileEntry("Root Entry", "", 5, 0, classId: storageClassId, stateBits: 0x12345678);
        var storageStreams = new Dictionary<string, byte[]> { ["Nested/Contents"] = Encoding.UTF8.GetBytes("Inert encrypted nested payload") };
        byte[] nested = OfficeCompoundFileWriter.Rewrite(new OfficeCompoundFile(storageStreams, storageEntries, root), new Dictionary<string, byte[]>());
        var entries = new[] { (storageName, false, nested), ("\u0005SummaryInformation", true, original.Streams["\u0005SummaryInformation"]),
            ("\u0005DocumentSummaryInformation", true, original.Streams["\u0005DocumentSummaryInformation"]) };
        byte[] summary = EncodeSummary(entries, session);
        string workbookName = withDescriptor!.Streams.ContainsKey("Workbook") ? "Workbook" : "Book";
        var replacements = new Dictionary<string, byte[]> {
            [workbookName] = BiffRc4RecordStreamCipher.Transform(withDescriptor.Streams[workbookName], session),
            ["encryption"] = summary,
            ["\u0005DocumentSummaryInformation"] = OfficeOlePropertySetWriter.CreatePropertySet((OfficeOlePropertySetWriter.DocumentSummaryInformationFormatId,
                OfficeOlePropertySetWriter.CreateSection(Array.Empty<OfficeOleProperty>())))
        };
        foreach (var item in withDescriptor.Streams)
            if (OfficeIMO.Excel.LegacyXls.Compound.LegacyXlsCompoundEncryption.IsAuxiliaryBiffStream(item.Key))
                replacements[item.Key] = BiffRc4RecordStreamCipher.Transform(item.Value, session);
        return OfficeCompoundFileWriter.Rewrite(withDescriptor, replacements, removedPaths: new[] { storageName, "\u0005SummaryInformation" });
    }

    internal static byte[] EncodeSummary((string Name, bool IsStream, byte[] Bytes)[] entries, OfficeBinaryRc4CryptoApiSession session) {
        using var output = new MemoryStream(); using var data = new BinaryWriter(output, Encoding.Unicode, true);
        data.Write(new byte[8]);
        using var descriptions = new MemoryStream(); using var descriptors = new BinaryWriter(descriptions, Encoding.Unicode, true);
        descriptors.Write(entries.Length);
        for (ushort index = 0; index < entries.Length; index++) {
            var entry = entries[index]; var encrypted = (byte[])entry.Bytes.Clone();
            session.TransformInPlace(encrypted, 0, encrypted.Length, index);
            descriptors.Write((uint)output.Position); descriptors.Write(encrypted.Length); descriptors.Write(index);
            descriptors.Write((byte)entry.Name.Length); descriptors.Write(entry.IsStream ? (byte)1 : (byte)0); descriptors.Write(0U);
            descriptors.Write(Encoding.Unicode.GetBytes(entry.Name)); descriptors.Write((ushort)0);
            data.Write(encrypted);
        }
        uint offset = (uint)output.Position; byte[] descriptorBytes = descriptions.ToArray();
        session.TransformInPlace(descriptorBytes, 0, descriptorBytes.Length, 0); data.Write(descriptorBytes);
        var location = new byte[8]; Buffer.BlockCopy(BitConverter.GetBytes(offset), 0, location, 0, 4);
        Buffer.BlockCopy(BitConverter.GetBytes(descriptorBytes.Length), 0, location, 4, 4);
        session.TransformInPlace(location, 0, location.Length, 0); output.Position = 0; data.Write(location);
        return output.ToArray();
    }
}
