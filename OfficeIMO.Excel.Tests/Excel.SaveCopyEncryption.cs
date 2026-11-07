using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelSaveCopyEncryptionTests {
    private const string Password = "copy-password";

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task EarlyBiffCopiesHonorTheOutputPassword(bool asynchronous) {
        string directory = CreateDirectory();
        string source = Path.Combine(directory, "source.xls");
        string destination = Path.Combine(directory, "copy.xlsx");
        byte[] input = {
            9, 4, 6, 0, 0, 0, 16, 0, 0, 0,
            3, 2, 14, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 240, 63,
            10, 0, 0, 0
        };
        try {
            File.WriteAllBytes(source, input);
            File.WriteAllBytes(destination, new byte[] { 17, 29, 41 });
            using var document = ExcelDocument.Load(source);
            string? originalPath = document.FilePath;
            string? originalSource = document.SourcePath;
            var options = new ExcelSaveOptions { Password = Password, Format = ExcelFileFormat.Xlsb };

            await Copy(document, destination, options, asynchronous);

            AssertEncrypted(destination, Password, 1d);
            Assert.Equal(input, File.ReadAllBytes(source));
            Assert.Equal(originalPath, document.FilePath);
            Assert.Equal(originalSource, document.SourcePath);
            Assert.Equal(ExcelFileFormat.Xls, document.SourceFormat);
            Assert.Equal(Password, options.Password);
            Assert.Equal(ExcelFileFormat.Xlsb, options.Format);
        } finally { Directory.Delete(directory, recursive: true); }
    }

    [Theory]
    [InlineData(false, "none")]
    [InlineData(true, "none")]
    [InlineData(false, "unset")]
    [InlineData(true, "unset")]
    [InlineData(false, "clear")]
    [InlineData(true, "clear")]
    [InlineData(false, "empty")]
    [InlineData(true, "empty")]
    [InlineData(false, "rotate")]
    [InlineData(true, "rotate")]
    public async Task CopiesPreserveLoadedPasswordIntentAndSourceAssociation(bool asynchronous, string policy) {
        string directory = CreateDirectory();
        string source = Path.Combine(directory, "source.xlsx");
        string destination = Path.Combine(directory, "copy.xlsx");
        try {
            using (var created = ExcelDocument.Create()) {
                created.AddSheet("Data").Cell(1, 1).SetValue(42d);
                File.WriteAllBytes(source, created.ToBytes(new ExcelSaveOptions { Password = Password }));
            }
            byte[] original = File.ReadAllBytes(source);
            using var document = ExcelDocument.Load(source, new ExcelLoadOptions { Password = Password });
            ExcelSaveOptions? options = policy == "none" ? null : new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb };
            if (policy == "clear") options!.Password = null;
            if (policy == "empty") options!.Password = string.Empty;
            if (policy == "rotate") options!.Password = "rotated-password";

            await Copy(document, destination, options, asynchronous);

            if (policy == "clear") {
                Assert.Equal(new byte[] { 0x50, 0x4b }, File.ReadAllBytes(destination).Take(2));
                using var copy = ExcelDocument.Load(destination);
                Assert.Equal(42d, copy.Sheets[0].Cell(1, 1).GetValue<double>());
            } else {
                AssertEncrypted(destination, policy == "empty" ? string.Empty : policy == "rotate" ? "rotated-password" : Password, 42d);
            }
            Assert.Equal(source, document.FilePath);
            Assert.Equal(source, document.SourcePath);
            Assert.Equal(original, File.ReadAllBytes(source));
            if (options != null) Assert.Equal(ExcelFileFormat.Xlsb, options.Format);
            using var next = new MemoryStream(document.ToBytes());
            using var stillProtected = ExcelDocument.Load(next, new ExcelLoadOptions { Password = Password });
            Assert.Equal(42d, stillProtected.Sheets[0].Cell(1, 1).GetValue<double>());
        } finally { Directory.Delete(directory, recursive: true); }
    }

    [Theory]
    [InlineData(false, ".xlsx", SpreadsheetDocumentType.Workbook)]
    [InlineData(true, ".xlsx", SpreadsheetDocumentType.Workbook)]
    [InlineData(false, ".xlsm", SpreadsheetDocumentType.MacroEnabledWorkbook)]
    [InlineData(true, ".xlsm", SpreadsheetDocumentType.MacroEnabledWorkbook)]
    [InlineData(false, ".xltx", SpreadsheetDocumentType.Template)]
    [InlineData(true, ".xltx", SpreadsheetDocumentType.Template)]
    [InlineData(false, ".xltm", SpreadsheetDocumentType.MacroEnabledTemplate)]
    [InlineData(true, ".xltm", SpreadsheetDocumentType.MacroEnabledTemplate)]
    [InlineData(false, ".xlam", SpreadsheetDocumentType.AddIn)]
    [InlineData(true, ".xlam", SpreadsheetDocumentType.AddIn)]
    public async Task EncryptedCopiesFinalizeTheirOwnPackageTypeAndRetainLiveSourceSheets(
        bool asynchronous, string extension, SpreadsheetDocumentType expectedType) {
        string directory = CreateDirectory();
        try {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue(42d);
            string destination = Path.Combine(directory, "copy" + extension);

            await Copy(document, destination, new ExcelSaveOptions { Password = Password }, asynchronous);

            AssertEncrypted(destination, Password, 42d);
            using (var copy = ExcelDocument.Load(destination, new ExcelLoadOptions { Password = Password })) {
                using var plain = new MemoryStream(copy.ToBytes(new ExcelSaveOptions { Password = null }));
                using var package = SpreadsheetDocument.Open(plain, false);
                Assert.Equal(expectedType, package.DocumentType);
            }
            Assert.Null(document.FilePath);
            sheet.Cell(1, 2).SetValue("After copy");
            using var sourceBytes = new MemoryStream(document.ToBytes());
            using (var package = SpreadsheetDocument.Open(sourceBytes, false)) {
                Assert.Equal(SpreadsheetDocumentType.Workbook, package.DocumentType);
            }
            sourceBytes.Position = 0;
            using var source = ExcelDocument.Load(sourceBytes);
            Assert.Equal("After copy", source.Sheets[0].Cell(1, 2).GetValue<string>());
        } finally { Directory.Delete(directory, recursive: true); }
    }

    [Theory]
    [InlineData(false, ".xls", ExcelFileFormat.Xls)]
    [InlineData(true, ".xls", ExcelFileFormat.Xls)]
    [InlineData(false, ".xlsb", ExcelFileFormat.Xlsb)]
    [InlineData(true, ".xlsb", ExcelFileFormat.Xlsb)]
    public async Task NativeCopiesUseTheSelectedPasswordAndFormat(bool asynchronous, string extension, ExcelFileFormat format) {
        string directory = CreateDirectory();
        try {
            using var document = ExcelDocument.Create();
            document.AddSheet("Data").Cell(1, 1).SetValue(42d);
            string destination = Path.Combine(directory, "copy" + extension);
            var options = new ExcelSaveOptions { Password = Password };

            await Copy(document, destination, options, asynchronous);

            using var copy = ExcelDocument.Load(destination, new ExcelLoadOptions { Password = Password });
            Assert.Equal(format, copy.SourceFormat);
            Assert.Equal(42d, copy.Sheets[0].Cell(1, 1).GetValue<double>());
            Assert.Null(document.FilePath);
            Assert.Equal(ExcelFileFormat.Xlsx, options.Format);
        } finally { Directory.Delete(directory, recursive: true); }
    }

    [Theory]
    [InlineData(false, "xml")]
    [InlineData(true, "xml")]
    [InlineData(false, "xls-password")]
    [InlineData(true, "xls-password")]
    [InlineData(false, "xlsb-policy")]
    [InlineData(true, "xlsb-policy")]
    [InlineData(false, "macro-type")]
    [InlineData(true, "macro-type")]
    public async Task CopyRefusalsPreserveAnExistingDestination(bool asynchronous, string refusal) {
        string directory = CreateDirectory();
        byte[] sentinel = { 17, 29, 41, 53 };
        try {
            string? source = refusal == "macro-type" ? Path.Combine(directory, "source.xlsm") : null;
            using var document = source == null ? ExcelDocument.Create() : ExcelDocument.Create(source);
            document.AddSheet("Data").Cell(1, 1).SetValue(42d);
            Assert.False(document.HasMacros);
            string extension = refusal == "xml" ? ".xml" : refusal == "xls-password" ? ".xls" : refusal == "xlsb-policy" ? ".xlsb" : ".xlsx";
            string destination = Path.Combine(directory, "copy" + extension);
            File.WriteAllBytes(destination, sentinel);
            var options = new ExcelSaveOptions { Password = refusal == "xls-password" ? new string('x', 16) : Password };
            if (refusal == "xlsb-policy") options.SafePreflight = true;

            Exception? failure = await Record.ExceptionAsync(() => Copy(document, destination, options, asynchronous));

            if (refusal == "xls-password") Assert.IsAssignableFrom<ArgumentException>(failure);
            else if (refusal == "macro-type") Assert.IsType<InvalidOperationException>(failure);
            else Assert.IsType<NotSupportedException>(failure);
            Assert.Equal(sentinel, File.ReadAllBytes(destination));
            Assert.Equal(source, document.FilePath);
            Assert.Equal(42d, document.Sheets[0].Cell(1, 1).GetValue<double>());
        } finally { Directory.Delete(directory, recursive: true); }
    }

    private static async Task Copy(ExcelDocument document, string path, ExcelSaveOptions? options, bool asynchronous) {
        if (asynchronous) await document.SaveCopyAsync(path, options);
        else document.SaveCopy(path, options);
    }

    private static void AssertEncrypted(string path, string password, double value) {
        Assert.Equal(new byte[] { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 }, File.ReadAllBytes(path).Take(8));
        Assert.Throws<InvalidDataException>(() => ExcelDocument.Load(path));
        using var copy = ExcelDocument.Load(path, new ExcelLoadOptions { Password = password });
        Assert.Equal(value, copy.Sheets[0].Cell(1, 1).GetValue<double>());
    }

    private static string CreateDirectory() => Directory.CreateDirectory(
        Path.Combine(Path.GetTempPath(), "OfficeIMO-CopyEncryption-" + Guid.NewGuid().ToString("N"))).FullName;
}
