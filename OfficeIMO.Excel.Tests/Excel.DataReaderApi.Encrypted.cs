using System.Globalization;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Core.Internal;
using OfficeIMO.Tests;
using OpenMcdf;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsb")]
    public void EncryptedReader_MatchesIndependentPlaintextAndOrdinaryReader(string extension) {
        byte[] encrypted = EncryptedReaderFixture(extension);
        byte[] plain = File.ReadAllBytes(EncryptedReaderFixturePath(extension, plain: true));
        // These plaintext bytes come from an independent decryptor. This proves
        // decryption compatibility without our encryptor generating the input.
        Assert.Equal(plain, OfficeEncryption.DecryptPackage(encrypted, "hunter2"));
        Assert.Throws<InvalidDataException>(() => OfficeEncryption.DecryptPackage(
            encrypted, "hunter2", CancellationToken.None, plain.Length - 1));
        using ExcelWorkbookDataReader expected = ExcelDocument.OpenDataReader(plain,
            new ExcelReadOptions { HasHeaderRow = false });
        using ExcelWorkbookDataReader actual = ExcelDocument.OpenEncryptedDataReader(encrypted, "hunter2",
            new ExcelReadOptions { HasHeaderRow = false });
        Assert.Equal(CaptureEncryptedReader(expected), CaptureEncryptedReader(actual));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void EncryptedReader_StreamOwnershipAndRemainingByteLimits(bool seekable) {
        byte[] encrypted = EncryptedReaderFixture("xlsx");
        long limit = encrypted.Length * 2L;
        int prefixLength = checked((int)limit + 17);
        byte[] prefixed = new byte[encrypted.Length + prefixLength];
        Array.Copy(encrypted, 0, prefixed, prefixLength, encrypted.Length);
        using Stream source = seekable
            ? new MemoryStream(prefixed, writable: false) { Position = prefixLength }
            : new NonSeekableReadStream(encrypted);
        using (ExcelWorkbookDataReader reader = ExcelDocument.OpenEncryptedDataReader(
            source, "hunter2", new ExcelReadOptions { MaxInputBytes = limit })) {
            Assert.True(reader.Read());
            Assert.Equal(1, reader.GetInt32(0));
            if (seekable) Assert.Equal(prefixLength, source.Position);
        }
        Assert.True(source.CanRead);
    }

    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsb")]
    public async Task EncryptedReader_PathOpenersReleaseFilesBeforeReaderDisposal(string extension) {
        string path = Path.Combine(Path.GetTempPath(), $"encrypted-reader-{Guid.NewGuid():N}.package");
        File.WriteAllBytes(path, EncryptedReaderFixture(extension));
        try {
            using (ExcelWorkbookDataReader reader = ExcelDocument.OpenEncryptedDataReader(path, "hunter2")) {
                using (File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { }
                Assert.True(reader.Read());
            }
            using (ExcelWorkbookDataReader reader = await ExcelDocument.OpenEncryptedDataReaderAsync(path, "hunter2")) {
                using (File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { }
                Assert.True(await reader.ReadAsync(CancellationToken.None));
            }
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task EncryptedReader_RejectsWrongPasswordsAndPreservesCallerPosition() {
        byte[] encrypted = EncryptedReaderFixture("xlsx");
        using var source = new MemoryStream(encrypted, writable: false);
        Assert.Throws<CryptographicException>(() => ExcelDocument.OpenEncryptedDataReader(source, "wrong"));
        Assert.Equal(0, source.Position);
        Assert.True(source.CanRead);
        await Assert.ThrowsAsync<CryptographicException>(() => ExcelDocument.OpenEncryptedDataReaderAsync(source, "wrong"));
        Assert.Equal(0, source.Position);
        Assert.True(source.CanRead);
    }

    [Fact]
    public void EncryptedReader_AlwaysVerifiesCiphertextIntegrity() {
        using var source = new MemoryStream();
        byte[] encrypted = EncryptedReaderFixture("xlsx");
        source.Write(encrypted, 0, encrypted.Length);
        source.Position = 0;
        using (RootStorage root = RootStorage.Open(source, StorageModeFlags.LeaveOpen)) {
            using CfbStream payload = root.OpenStream("EncryptedPackage");
            payload.Position = 19;
            int original = payload.ReadByte();
            payload.Position = 19;
            payload.WriteByte((byte)(original ^ 1));
            root.Flush();
        }
        Assert.Throws<CryptographicException>(() => ExcelDocument.OpenEncryptedDataReader(source.ToArray(), "hunter2"));
    }

    [Fact]
    public async Task EncryptedReader_BoundsInputAndRestoresStreamsOnLimitFailure() {
        byte[] encrypted = EncryptedReaderFixture("xlsx");
        var options = new ExcelReadOptions { MaxInputBytes = encrypted.Length - 1 };
        Assert.Throws<InvalidDataException>(() => ExcelDocument.OpenEncryptedDataReader(encrypted, "hunter2", options));
        using var source = new MemoryStream(encrypted, writable: false);
        Assert.Throws<InvalidDataException>(() => ExcelDocument.OpenEncryptedDataReader(source, "hunter2", options));
        Assert.Equal(0, source.Position);
        await Assert.ThrowsAsync<InvalidDataException>(() => ExcelDocument.OpenEncryptedDataReaderAsync(source, "hunter2", options));
        Assert.Equal(0, source.Position);
        using var forward = new NonSeekableReadStream(encrypted);
        Assert.Throws<InvalidDataException>(() => ExcelDocument.OpenEncryptedDataReader(forward, "hunter2", options));
        Assert.True(forward.CanRead);
    }

    [Fact]
    public async Task EncryptedReader_AsyncUsesAsyncInputAndSnapshotsOptions() {
        byte[] encrypted = EncryptedReaderFixture("xlsx");
        var gate = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        using var source = new AsyncOpeningReadStream(encrypted, seekable: true) { ReadGate = gate.Task };
        var options = new ExcelReadOptions { HasHeaderRow = false };
        Task<ExcelWorkbookDataReader> opening = ExcelDocument.OpenEncryptedDataReaderAsync(source, "hunter2", options);
        Assert.False(opening.IsCompleted);
        options.MaxInputBytes = 1;
        options.SheetName = "Missing";
        options.HasHeaderRow = true;
        gate.SetResult(true);
        using ExcelWorkbookDataReader actual = await opening;
        using ExcelWorkbookDataReader expected = ExcelDocument.OpenDataReader(
            File.ReadAllBytes(EncryptedReaderFixturePath("xlsx", plain: true)), new ExcelReadOptions { HasHeaderRow = false });
        Assert.Equal(CaptureEncryptedReader(expected), CaptureEncryptedReader(actual));
        Assert.True(source.AsyncReads > 0);
        Assert.Equal(0, source.Position);
        actual.Close();
        Assert.True(source.CanRead);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task EncryptedReader_AsyncRetainsBothOpeningCancellationTokens(bool optionsToken) {
        using var optionsCancellation = new CancellationTokenSource();
        using var openingCancellation = new CancellationTokenSource();
        using var source = new MemoryStream(EncryptedReaderFixture("xlsx"), writable: false);
        using ExcelWorkbookDataReader reader = await ExcelDocument.OpenEncryptedDataReaderAsync(source, "hunter2",
            new ExcelReadOptions { CancellationToken = optionsCancellation.Token }, openingCancellation.Token);
        (optionsToken ? optionsCancellation : openingCancellation).Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => reader.Read());
        Assert.True(source.CanRead);
    }

    [Fact]
    public async Task EncryptedReader_ObservesOpeningCancellationWithoutClosingTheSource() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        using var beforeRead = new AsyncOpeningReadStream(EncryptedReaderFixture("xlsx"), seekable: true);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => ExcelDocument.OpenEncryptedDataReaderAsync(
            beforeRead, "hunter2", cancellationToken: cancellation.Token));
        Assert.Equal(0, beforeRead.AsyncReads);
        Assert.True(beforeRead.CanRead);

        using var duringReadCancellation = new CancellationTokenSource();
        using var duringRead = new AsyncOpeningReadStream(EncryptedReaderFixture("xlsx"), seekable: true) {
            BeforeRead = () => duringReadCancellation.Cancel()
        };
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => ExcelDocument.OpenEncryptedDataReaderAsync(
            duringRead, "hunter2", cancellationToken: duringReadCancellation.Token));
        Assert.Equal(0, duringRead.Position);
        Assert.True(duringRead.CanRead);
    }

    [Fact]
    public void EncryptedReader_RequiresAnEncryptedOpenXmlPackageAndValidArguments() {
        byte[] encrypted = EncryptedReaderFixture("xlsx");
        Assert.Throws<ArgumentNullException>(() => ExcelDocument.OpenEncryptedDataReader((byte[])null!, "hunter2"));
        Assert.Throws<ArgumentNullException>(() => ExcelDocument.OpenEncryptedDataReader(encrypted, null!));
        Assert.Throws<ArgumentNullException>(() => ExcelDocument.OpenEncryptedDataReader((Stream)null!, "hunter2"));
        Assert.Throws<ArgumentException>(() => ExcelDocument.OpenEncryptedDataReader(" ", "hunter2"));
        Assert.Throws<InvalidDataException>(() => ExcelDocument.OpenEncryptedDataReader(
            File.ReadAllBytes(EncryptedReaderFixturePath("xlsx", plain: true)), "hunter2"));
        string legacyPath = Path.Combine(AppContext.BaseDirectory, "Documents", "LegacyXlsDiagnosticCorpus",
            "excel-com-generated", "encrypted-password.xls");
        Assert.Throws<InvalidDataException>(() => ExcelDocument.OpenEncryptedDataReader(File.ReadAllBytes(legacyPath), "openpass"));
    }

    private static byte[] EncryptedReaderFixture(string extension) => File.ReadAllBytes(EncryptedReaderFixturePath(extension));

    private static string EncryptedReaderFixturePath(string extension, bool plain = false) => Path.Combine(
        AppContext.BaseDirectory, "Documents", "ExcelEncryptionCorpus", $"agile-aes256-sha512{(plain ? ".plain" : "")}.{extension}");

    private static List<string> CaptureEncryptedReader(ExcelWorkbookDataReader reader) {
        var values = new List<string>();
        do {
            values.Add(reader.CurrentSheetName);
            values.Add(reader.FieldCount.ToString(CultureInfo.InvariantCulture));
            while (reader.Read()) {
                values.Add("row");
                for (int ordinal = 0; ordinal < reader.FieldCount; ordinal++) {
                    object value = reader.GetValue(ordinal);
                    values.Add(value.GetType().Name + ":" + Convert.ToString(value, CultureInfo.InvariantCulture));
                }
            }
        } while (reader.NextResult());
        return values;
    }
}
