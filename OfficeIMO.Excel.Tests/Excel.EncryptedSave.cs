using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelEncryptedSaveTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Encrypted_output_is_not_replaced_by_plaintext_on_dispose(bool stream) {
        const string password = "test-password";
        if (stream) {
            using var destination = new MemoryStream();
            using (var document = ExcelDocument.Create(destination, new() { PersistenceMode = DocumentPersistenceMode.SaveOnDispose })) {
                document.AddSheet("Data").CellValue(1, 1, "Encrypted value"); document.SaveEncrypted(destination, password);
            }
            destination.Position = 0;
            using var saved = ExcelDocument.LoadEncrypted(destination, password);
            Assert.Equal("Encrypted value", saved.Sheets[0].CellAt(1, 1).GetValue<string>());
        } else {
            string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-Encrypted-" + Guid.NewGuid().ToString("N") + ".xlsx");
            try {
                using (var document = ExcelDocument.Create(path, new ExcelCreateOptions { PersistenceMode = DocumentPersistenceMode.SaveOnDispose })) {
                    document.AddSheet("Data").CellValue(1, 1, "Encrypted value"); document.SaveEncrypted(path, password);
                }
                using var saved = ExcelDocument.LoadEncrypted(path, password);
                Assert.Equal("Encrypted value", saved.Sheets[0].CellAt(1, 1).GetValue<string>());
            } finally { if (File.Exists(path)) File.Delete(path); }
        }
    }
    [Theory]
    [InlineData(ExcelFileFormat.Xls)]
    [InlineData(ExcelFileFormat.Xlsb)]
    public void Encrypted_native_formats_round_trip_cells_formulas_and_password_changes(ExcelFileFormat format) {
        using var document = ExcelDocument.Create(); var sheet = document.AddSheet("Data");
        sheet.CellValue(1, 1, 12.5); sheet.CellValue(2, 1, "00127"); sheet.CellValue(3, 1, true);
        sheet.CellValue(4, 1, 25d); sheet.CellFormula(4, 1, "A1*2");
        using var output = new MemoryStream(); document.SaveEncrypted(output, "test-password", format);
        Assert.Equal(0, output.Position);
        Assert.Equal(new byte[] { 0xd0, 0xcf, 0x11, 0xe0 }, output.ToArray().Take(4));
        output.Position = 0; using var saved = ExcelDocument.LoadEncrypted(output, "test-password");
        Assert.Equal(format, saved.SourceFormat); Assert.Equal(12.5, saved.Sheets[0].CellAt(1, 1).GetValue<double>());
        Assert.Equal("00127", saved.Sheets[0].CellAt(2, 1).GetValue<string>()); Assert.True(saved.Sheets[0].CellAt(3, 1).GetValue<bool>());
        Assert.Equal("A1*2", saved.Sheets[0].GetFormulaText(4, 1));
        saved.Sheets[0].CellValue(1, 1, 50d);
        using var second = new MemoryStream(); saved.SaveEncrypted(second, "new-password", format);
        second.Position = 0; using var updated = ExcelDocument.LoadEncrypted(second, "new-password");
        Assert.Equal(50d, updated.Sheets[0].CellAt(1, 1).GetValue<double>());
        second.Position = 0; Assert.ThrowsAny<Exception>(() => ExcelDocument.LoadEncrypted(second, "test-password"));
        Assert.NotEqual(output.ToArray(), second.ToArray());
    }

    [Theory]
    [InlineData(".xls", ExcelFileFormat.Xls)]
    [InlineData(".xlsb", ExcelFileFormat.Xlsb)]
    public void File_extension_selects_the_encrypted_native_encoder(string extension, ExcelFileFormat format) {
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-Encrypted-" + Guid.NewGuid().ToString("N") + extension);
        try {
            using (var document = ExcelDocument.Create()) {
                document.AddSheet("Data").CellValue(1, 1, "Native encrypted output"); document.SaveEncrypted(path, "test-password");
            }
            using var saved = ExcelDocument.LoadEncrypted(path, "test-password");
            Assert.Equal(format, saved.SourceFormat); Assert.Equal("Native encrypted output", saved.Sheets[0].CellAt(1, 1).GetValue<string>());
        } finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Unsupported_encryption_policy_does_not_touch_destinations() {
        using var document = ExcelDocument.Create(); document.AddSheet("Data").CellValue(1, 1, "Keep");
        using var output = new MemoryStream(new byte[] { 1, 2, 3, 4 }, writable: true);
        Assert.Throws<ArgumentException>(() => document.SaveEncrypted(output, new string('x', 16), ExcelFileFormat.Xls));
        Assert.Equal(new byte[] { 1, 2, 3, 4 }, output.ToArray());
        Assert.Throws<NotSupportedException>(() => document.SaveEncrypted(output, "test-password", ExcelFileFormat.Xlsb, new() { SafePreflight = true }));
        Assert.Equal(new byte[] { 1, 2, 3, 4 }, output.ToArray());
    }

    [Fact]
    public void Exporting_encrypted_to_another_stream_preserves_original_automatic_persistence() {
        using var original = new MemoryStream(); using var encrypted = new MemoryStream();
        using (var document = ExcelDocument.Create(original, new() { PersistenceMode = DocumentPersistenceMode.SaveOnDispose })) {
            var sheet = document.AddSheet("Data"); sheet.CellValue(1, 1, "Encrypted snapshot"); document.SaveEncrypted(encrypted, "test-password");
            sheet.CellValue(1, 1, "Original after export");
        }
        original.Position = 0; encrypted.Position = 0;
        using var saved = ExcelDocument.Load(original); using var copy = ExcelDocument.LoadEncrypted(encrypted, "test-password");
        Assert.Equal("Original after export", saved.Sheets[0].CellAt(1, 1).GetValue<string>());
        Assert.Equal("Encrypted snapshot", copy.Sheets[0].CellAt(1, 1).GetValue<string>());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async System.Threading.Tasks.Task Later_edits_and_async_disposal_do_not_implicitly_overwrite_encryption(bool fileBacked) {
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-Encrypted-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try {
            using (var seed = ExcelDocument.Create()) { seed.AddSheet("Data").CellValue(1, 1, "Snapshot"); seed.Save(path); }
            var options = new ExcelLoadOptions { PersistenceMode = DocumentPersistenceMode.SaveOnDispose };
            await using (var document = fileBacked ? ExcelDocument.OpenFileBacked(path, options) : ExcelDocument.Load(path, options)) {
                document.SaveEncrypted(path, "test-password"); document.Sheets[0].CellValue(1, 1, "Unsaved edit");
            }
            using var saved = ExcelDocument.LoadEncrypted(path, "test-password"); Assert.Equal("Snapshot", saved.Sheets[0].CellAt(1, 1).GetValue<string>());
        } finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Theory]
    [InlineData("encrypted-native-control.xls", ExcelFileFormat.Xls)]
    [InlineData("encrypted-native-control.xlsb", ExcelFileFormat.Xlsb)]
    public void Exact_native_accepted_encrypted_controls_keep_their_payload(string fixture, ExcelFileFormat format) {
        string path = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", "ExcelEncryptionCorpus", fixture);
        using var document = ExcelDocument.LoadEncrypted(path, "test-password");
        Assert.Equal(format, document.SourceFormat);
        Assert.Equal("Native encrypted control", document.Sheets[0].CellAt(1, 1).GetValue<string>());
        Assert.Equal(12.5d, document.Sheets[0].CellAt(2, 1).GetValue<double>());
        Assert.True(document.Sheets[0].CellAt(3, 1).GetValue<bool>());
        Assert.Equal("Łódź 00127", document.Sheets[0].CellAt(4, 1).GetValue<string>());
        Assert.Equal("A2*2", document.Sheets[0].GetFormulaText(5, 1));
        Assert.Equal(25d, document.Sheets[0].CellAt(5, 1).GetValue<double>());
    }

    [Theory]
    [InlineData(ExcelFileFormat.Xls)]
    [InlineData(ExcelFileFormat.Xlsb)]
    public void Native_encrypted_materialization_is_bounded_before_destination_writes(ExcelFileFormat format) {
        using var document = ExcelDocument.Create(); var sheet = document.AddSheet("Data");
        for (int i = 1; i <= 20; i++) sheet.CellValue(i, 1, new string((char)('A' + i), 500));
        using var output = new MemoryStream(new byte[] { 1, 2, 3, 4 }, writable: true);
        var exception = Assert.Throws<InvalidDataException>(() => document.SaveEncrypted(output, "test-password", format,
            new ExcelSaveOptions { MaxInMemoryPackageBytes = 1024 }));
        Assert.True(exception.Data.Contains("OfficeIMO.OutputLimit"));
        Assert.Equal(new byte[] { 1, 2, 3, 4 }, output.ToArray());
    }

    [Theory]
    [InlineData(ExcelFileFormat.Xlsx)]
    [InlineData(ExcelFileFormat.Xls)]
    [InlineData(ExcelFileFormat.Xlsb)]
    public void A_flush_failure_after_encrypted_writing_never_causes_plaintext_copyback(ExcelFileFormat format) {
        using var output = new FailingFlushStream();
        using (var document = ExcelDocument.Create(output, new() { PersistenceMode = DocumentPersistenceMode.SaveOnDispose })) {
            document.AddSheet("Data").CellValue(1, 1, "Protected snapshot");
            output.FailFlush = true;
            Assert.Throws<IOException>(() => document.SaveEncrypted(output, "test-password", format));
            output.FailFlush = false;
        }
        output.Position = 0;
        using var saved = ExcelDocument.LoadEncrypted(output, "test-password");
        Assert.Equal("Protected snapshot", saved.Sheets[0].CellAt(1, 1).GetValue<string>());
    }

    [Theory]
    [InlineData(ExcelFileFormat.Xlsx)]
    [InlineData(ExcelFileFormat.Xls)]
    [InlineData(ExcelFileFormat.Xlsb)]
    public void A_partial_encrypted_write_never_triggers_plaintext_copyback(ExcelFileFormat format) {
        using var output = new FailingFlushStream(); byte[] failedSnapshot;
        using (var document = ExcelDocument.Create(output, new() { PersistenceMode = DocumentPersistenceMode.SaveOnDispose })) {
            document.AddSheet("Data").CellValue(1, 1, "Protected snapshot");
            output.FailWrite = true;
            Assert.Throws<IOException>(() => document.SaveEncrypted(output, "test-password", format));
            output.FailWrite = false; failedSnapshot = output.ToArray();
        }
        Assert.Equal(new byte[] { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 }, failedSnapshot);
        Assert.Equal(failedSnapshot, output.ToArray());
    }

    private sealed class FailingFlushStream : MemoryStream {
        internal bool FailFlush { get; set; }
        internal bool FailWrite { get; set; }
        public override void Write(byte[] buffer, int offset, int count) {
            if (FailWrite) { base.Write(buffer, offset, Math.Min(8, count)); throw new IOException("Synthetic partial write failure."); }
            base.Write(buffer, offset, count);
        }
        public override void Flush() { if (FailFlush) throw new IOException("Synthetic destination flush failure."); base.Flush(); }
    }

}
