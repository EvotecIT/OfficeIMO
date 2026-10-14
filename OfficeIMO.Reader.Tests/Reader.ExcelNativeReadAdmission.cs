using System;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Reader.Excel;
using OfficeIMO.Reader.Zip;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderExcelNativeReadAdmissionTests {
    [Theory]
    [InlineData("MaxSharedStringItems", 1)]
    [InlineData("MaxSharedStringItemCharacters", 1)]
    [InlineData("MaxSharedStringCharacters", 6)]
    [InlineData("MaxMetadataPartBytes", 64)]
    [InlineData("MaxWorksheets", 1)]
    [InlineData("MaxInputBytes", 64)]
    public void NativeDataReaderPreservesConfiguredAdmission(string name, long maximum) {
        byte[] bytes = Workbook();
        ExcelReadOptions options = Configure(name, maximum);
        OfficeInspectionResourceLimitException error = Assert.Throws<OfficeInspectionResourceLimitException>(() => {
            using DbDataReader reader = ExcelDocument.OpenDataReader(bytes, options);
            do { while (reader.Read()) _ = reader.GetValue(0); } while (reader.NextResult());
        });
        Assert.Equal("excel", error.Owner); Assert.Equal(name, error.LimitName); Assert.Equal(maximum, error.Maximum);
    }

    [Theory]
    [InlineData("MaxSharedStringItems", 1, "loaded")]
    [InlineData("MaxSharedStringItems", 1, "loaded-editable")]
    [InlineData("MaxSharedStringItemCharacters", 1, "loaded")]
    [InlineData("MaxSharedStringCharacters", 6, "loaded")]
    [InlineData("MaxSharedStringItems", 1, "direct")]
    [InlineData("MaxSharedStringItems", 1, "incremental")]
    [InlineData("MaxSharedStringItems", 1, "batch")]
    [InlineData("MaxSharedStringItems", 1, "folder")]
    [InlineData("MaxSharedStringItems", 1, "zip")]
    public async Task ReaderProjectionKeepsNativeAdmissionTerminal(string name, long maximum, string route) {
        byte[] bytes = Workbook();
        var format = new ReaderExcelOptions { HeadersInFirstRow = false, ReadOptions = Configure(name, maximum) };
        var options = new ReaderOptions { Projection = route == "zip" ? ReaderProjection.Full : ReaderProjection.TextAndMetadata };
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddExcelHandler(format).AddZipHandler().Build();
        string directory = Path.Combine(Path.GetTempPath(), "excel-admission-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(directory);
        try {
            ReaderResourceLimitException error = await Assert.ThrowsAsync<ReaderResourceLimitException>(async () => {
                if (route == "loaded" || route == "loaded-editable") {
                    using ExcelDocument document = ExcelDocument.Load(new MemoryStream(bytes), new ExcelLoadOptions { AccessMode = route == "loaded" ? DocumentAccessMode.ReadOnly : DocumentAccessMode.ReadWrite });
                    document.ReadDocument("source.xlsx", options, format);
                } else if (route == "direct") reader.ReadDocument(bytes, "source.xlsx", options);
                else if (route == "zip") reader.ReadDocument(Wrap(bytes), "source.zip", options);
                else {
                    string path = Path.Combine(directory, "source.xlsx"); File.WriteAllBytes(path, bytes);
                    if (route == "incremental") reader.EnumerateChunks(path, options).ToArray();
                    else if (route == "batch") await reader.ReadDocumentsDetailedAsync(new[] { path }, options);
                    else reader.ReadFolder(directory, options: options).ToArray();
                }
            });
            Assert.Equal("excel." + name, error.LimitName); Assert.Equal(maximum, error.Maximum);
            OfficeInspectionResourceLimitException native = Assert.IsType<OfficeInspectionResourceLimitException>(error.InnerException);
            Assert.Equal(name, native.LimitName); Assert.Equal(maximum, native.Maximum);
        } finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public async Task MalformedWorkbookRemainsAnOrdinaryPerFileFailure() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try {
            File.WriteAllBytes(path, new byte[] { 1, 2, 3 });
            ReaderDocumentReadOutcome outcome = Assert.Single(await new OfficeDocumentReaderBuilder().AddExcelHandler().Build().ReadDocumentsDetailedAsync(new[] { path }));
            Assert.NotNull(outcome.Error); Assert.IsNotType<ReaderResourceLimitException>(outcome.Error); Assert.Null(outcome.Document);
        } finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void OrdinaryLoadedProjectionRetainsCallerAndFullOutput() {
        using ExcelDocument document = ExcelDocument.Load(new MemoryStream(Workbook()));
        OfficeDocumentReadResult result = document.ReadDocument("source.xlsx", formatOptions: new ReaderExcelOptions { HeadersInFirstRow = false });
        Assert.Contains(result.Chunks, chunk => chunk.Text!.Contains("first", StringComparison.Ordinal));
        Assert.NotNull(document.OpenXmlDocument.WorkbookPart); Assert.Equal(2, document.Sheets.Count);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task NativeInputBudgetRestoresCallerStream(bool asynchronous) {
        using var stream = new MemoryStream(Workbook(), writable: false);
        stream.Position = 7;
        var options = new ExcelReadOptions { MaxInputBytes = 64 };
        OfficeInspectionResourceLimitException error = asynchronous
            ? await Assert.ThrowsAsync<OfficeInspectionResourceLimitException>(() => ExcelDocument.OpenDataReaderAsync(stream, options))
            : Assert.Throws<OfficeInspectionResourceLimitException>(() => ExcelDocument.OpenDataReader(stream, options));
        Assert.Equal("excel", error.Owner); Assert.Equal(nameof(ExcelReadOptions.MaxInputBytes), error.LimitName); Assert.Equal(64, error.Maximum);
        Assert.Equal(7, stream.Position); Assert.True(stream.CanRead);
    }

    [Theory]
    [InlineData("MaxMetadataPartBytes", 64)]
    [InlineData("MaxWorksheets", 1)]
    public void NativeSdkMetadataAndNameOnlyRoutesKeepExactLimit(string name, long maximum) {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsm");
        try {
            File.WriteAllBytes(path, Workbook());
            OfficeInspectionResourceLimitException error = Assert.Throws<OfficeInspectionResourceLimitException>(() => ExcelDocument.OpenDataReader(path, Configure(name, maximum)));
            Assert.Equal("excel", error.Owner); Assert.Equal(name, error.LimitName); Assert.Equal(maximum, error.Maximum);
            error = Assert.Throws<OfficeInspectionResourceLimitException>(() => ExcelDocument.GetSheetNames(path, Configure(name, maximum)));
            Assert.Equal(name, error.LimitName); Assert.Equal(maximum, error.Maximum);
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void NativeDataReaderConfiguredAllocationBoundsRemainTyped(bool columns) {
        var options = new ExcelReadOptions { HasHeaderRow = false, A1Range = "A1:B2", MaxDataReaderChunkRows = 1 };
        if (columns) options.MaxDataReaderColumns = 1;
        else options.MaxDataReaderBufferedCells = 1;
        OfficeInspectionResourceLimitException error = Assert.Throws<OfficeInspectionResourceLimitException>(() => ExcelDocument.OpenDataReader(Workbook(), options));
        Assert.Equal("excel", error.Owner);
        Assert.Equal(columns ? nameof(ExcelReadOptions.MaxDataReaderColumns) : nameof(ExcelReadOptions.MaxDataReaderBufferedCells), error.LimitName);
        Assert.Equal(1, error.Maximum);
    }

    [Fact]
    public void LoadedProjectionDoesNotRetroactivelyAdmitPackageInput() {
        using ExcelDocument document = ExcelDocument.Load(new MemoryStream(Workbook()));
        var format = new ReaderExcelOptions { HeadersInFirstRow = false, ReadOptions = new ExcelReadOptions { MaxInputBytes = 1, MaxMetadataPartBytes = 1, MaxWorksheets = 1 } };
        OfficeDocumentReadResult result = document.ReadDocument("source.xlsx", new ReaderOptions { Projection = ReaderProjection.TextAndMetadata }, format);
        Assert.Contains(result.Chunks, chunk => chunk.Text.Contains("first", StringComparison.Ordinal));
        Assert.Equal(2, document.Sheets.Count);
    }

    [Fact]
    public void NativeProjectionFailurePreservesLoadedOwnerAndReadScope() {
        using ExcelDocument document = ExcelDocument.Load(new MemoryStream(Workbook()));
        Assert.Throws<ReaderResourceLimitException>(() => document.ReadDocument("source.xlsx",
            new ReaderOptions { Projection = ReaderProjection.TextAndMetadata },
            new ReaderExcelOptions { HeadersInFirstRow = false, ReadOptions = Configure(nameof(ExcelReadOptions.MaxSharedStringItems), 1) }));
        Assert.Equal(2, document.Sheets.Count); Assert.NotNull(document.OpenXmlDocument.WorkbookPart);
        OfficeDocumentReadResult recovered = document.ReadDocument("source.xlsx", formatOptions: new ReaderExcelOptions { HeadersInFirstRow = false });
        Assert.Contains(recovered.Chunks, chunk => chunk.Text.Contains("second", StringComparison.Ordinal));
    }

    [Fact]
    public void CanonicalDecryptedPackageAdmissionRetainsItsOwner() {
        byte[] encrypted = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Documents", "ExcelEncryptionCorpus", "agile-aes256-sha512.xlsx"));
        byte[] plain = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Documents", "ExcelEncryptionCorpus", "agile-aes256-sha512.plain.xlsx"));
        OfficeInspectionResourceLimitException error = Assert.Throws<OfficeInspectionResourceLimitException>(() =>
            OfficeIMO.Core.Internal.OfficeEncryption.DecryptPackage(encrypted, "hunter2", default, plain.Length - 1));
        Assert.Equal("encryption", error.Owner); Assert.Equal("MaximumDecryptedPackageBytes", error.LimitName); Assert.Equal(plain.Length - 1, error.Maximum);
        Assert.Equal(plain, OfficeIMO.Core.Internal.OfficeEncryption.DecryptPackage(encrypted, "hunter2"));
        error = Assert.Throws<OfficeInspectionResourceLimitException>(() => ExcelDocument.OpenEncryptedDataReader(encrypted, "hunter2", new ExcelReadOptions { MaxInputBytes = encrypted.Length - 1 }));
        Assert.Equal("excel", error.Owner); Assert.Equal(nameof(ExcelReadOptions.MaxInputBytes), error.LimitName); Assert.Equal(encrypted.Length - 1, error.Maximum);
    }

    private static ExcelReadOptions Configure(string name, long maximum) {
        var options = new ExcelReadOptions { HasHeaderRow = false };
        switch (name) {
            case "MaxSharedStringItems": options.MaxSharedStringItems = (int)maximum; break;
            case "MaxSharedStringItemCharacters": options.MaxSharedStringItemCharacters = (int)maximum; break;
            case "MaxSharedStringCharacters": options.MaxSharedStringCharacters = maximum; break;
            case "MaxMetadataPartBytes": options.MaxMetadataPartBytes = (int)maximum; break;
            case "MaxWorksheets": options.MaxWorksheets = (int)maximum; break;
            case "MaxInputBytes": options.MaxInputBytes = maximum; break;
        }
        return options;
    }

    private static byte[] Workbook() {
        using ExcelDocument document = ExcelDocument.Create();
        var sheet = document.AddWorksheet("Evidence"); sheet.CellValue(1, 1, "first"); sheet.CellValue(2, 1, "second"); document.AddWorksheet("Second");
        WorkbookPart workbook = document.OpenXmlDocument.WorkbookPart!;
        SharedStringTablePart strings = workbook.SharedStringTablePart ?? workbook.AddNewPart<SharedStringTablePart>();
        strings.SharedStringTable = new SharedStringTable(new SharedStringItem(new Text("first")), new SharedStringItem(new Text("second"))) { Count = 2, UniqueCount = 2 };
        var cells = workbook.WorksheetParts.First().Worksheet!.Descendants<Cell>().ToArray();
        for (int index = 0; index < cells.Length; index++) { cells[index].DataType = CellValues.SharedString; cells[index].CellValue = new CellValue(index.ToString()); cells[index].InlineString = null; }
        return document.ToBytes();
    }

    private static byte[] Wrap(byte[] bytes) {
        using var output = new MemoryStream();
        using (var archive = new System.IO.Compression.ZipArchive(output, System.IO.Compression.ZipArchiveMode.Create, true)) {
            using Stream member = archive.CreateEntry("source.xlsx").Open(); member.Write(bytes, 0, bytes.Length);
        }
        return output.ToArray();
    }
}
