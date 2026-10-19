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
    public void Encrypted_summary_restores_VBA_storage_metadata_and_supports_plaintext_save_and_password_rotation() {
        string path = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", "ExcelPreservationCorpus", "excel-populated-macro.xls");
        byte[] source = File.ReadAllBytes(path);
        byte[] encrypted = EncryptWithSummary(source, out OfficeCompoundFile original, out Guid storageClassId);
        using var document = ExcelDocument.LoadEncrypted(new MemoryStream(encrypted), "test-password");
        document.Sheets[0].CellValue(1, 1, "CryptoAPI restored storage");
        byte[] plain = document.ToBytes(ExcelFileFormat.Xls);
        Assert.True(OfficeCompoundFileReader.TryRead(plain, out OfficeCompoundFile? result, out string? error), error);
        Assert.False(result!.Streams.ContainsKey("encryption"));
        Assert.DoesNotContain(BiffRecordReader.ReadRecords(result.Streams["Workbook"], new()), item => item.Type == 0x002f);
        foreach (var stream in original.Streams.Where(item => item.Key.StartsWith("_VBA_PROJECT_CUR/", StringComparison.OrdinalIgnoreCase)))
            Assert.Equal(stream.Value, result.Streams[stream.Key]);
        var storage = Assert.Single(result.Entries, entry => entry.Path == "_VBA_PROJECT_CUR");
        Assert.Equal(storageClassId, storage.ClassId);
        Assert.Equal(0x12345678U, storage.StateBits);
        using var output = new MemoryStream(); document.SaveEncrypted(output, "new-password", ExcelFileFormat.Xls);
        using var reloaded = ExcelDocument.LoadEncrypted(output, "new-password");
        Assert.Equal("CryptoAPI restored storage", reloaded.Sheets[0].CellAt(1, 1).GetValue<string>());
        Assert.Equal(4, reloaded.LegacyXlsChartSheets.Count);
    }

    [Fact]
    public void Truncated_encrypted_summary_is_rejected_instead_of_losing_hidden_content() {
        string path = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", "ExcelPreservationCorpus", "excel-populated-macro.xls");
        byte[] encrypted = EncryptWithSummary(File.ReadAllBytes(path), out _, out _);
        Assert.True(OfficeCompoundFileReader.TryRead(encrypted, out OfficeCompoundFile? compound, out string? error), error);
        byte[] damaged = OfficeCompoundFileWriter.Rewrite(compound!, new Dictionary<string, byte[]> { ["encryption"] = new byte[7] });
        Assert.Throws<InvalidDataException>(() => ExcelDocument.LoadEncrypted(new MemoryStream(damaged), "test-password"));
    }

    private static byte[] EncryptWithSummary(byte[] bytes, out OfficeCompoundFile original, out Guid storageClassId) {
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
        const string storageName = "_VBA_PROJECT_CUR";
        storageClassId = new Guid("6842D9E3-032D-448D-B646-BDA84BDE218A");
        var storageEntries = original.Entries.Where(entry => entry.Path.StartsWith(storageName + "/", StringComparison.OrdinalIgnoreCase))
            .Select(entry => new OfficeCompoundFileEntry(entry.Name, entry.Path.Substring(storageName.Length + 1), entry.ObjectType,
                entry.Size, classId: entry.ClassId, stateBits: entry.StateBits, creationTime: entry.CreationTime, modifiedTime: entry.ModifiedTime)).ToList();
        var root = new OfficeCompoundFileEntry("Root Entry", "", 5, 0, classId: storageClassId, stateBits: 0x12345678);
        var storageStreams = original.Streams.Where(item => item.Key.StartsWith(storageName + "/", StringComparison.OrdinalIgnoreCase))
            .ToDictionary(item => item.Key.Substring(storageName.Length + 1), item => item.Value);
        byte[] nested = OfficeCompoundFileWriter.Rewrite(new OfficeCompoundFile(storageStreams, storageEntries, root), new Dictionary<string, byte[]>());
        var entries = new[] { (storageName, false, nested), ("\u0005SummaryInformation", true, original.Streams["\u0005SummaryInformation"]),
            ("\u0005DocumentSummaryInformation", true, original.Streams["\u0005DocumentSummaryInformation"]) };
        byte[] summary = EncodeSummary(entries, session);
        return OfficeCompoundFileWriter.Rewrite(withDescriptor!, new Dictionary<string, byte[]> {
            ["Workbook"] = BiffRc4RecordStreamCipher.Transform(withDescriptor!.Streams["Workbook"], session),
            ["encryption"] = summary,
            ["\u0005DocumentSummaryInformation"] = OfficeOlePropertySetWriter.CreatePropertySet((OfficeOlePropertySetWriter.DocumentSummaryInformationFormatId,
                OfficeOlePropertySetWriter.CreateSection(Array.Empty<OfficeOleProperty>())))
        }, removedPaths: new[] { storageName, "\u0005SummaryInformation" });
    }

    private static byte[] EncodeSummary((string Name, bool IsStream, byte[] Bytes)[] entries, OfficeBinaryRc4CryptoApiSession session) {
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
