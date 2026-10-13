using System;
using System.IO;
using System.Threading.Tasks;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EncryptedSourceInspectionContracts {
    private const string Password = "source-inspection-test";

    [Theory]
    [InlineData("word", "path")]
    [InlineData("word", "stream")]
    [InlineData("excel", "path")]
    [InlineData("excel", "stream")]
    [InlineData("excel", "path-async")]
    [InlineData("excel", "stream-async")]
    [InlineData("powerpoint", "path")]
    [InlineData("powerpoint", "stream")]
    [InlineData("powerpoint", "path-async")]
    [InlineData("powerpoint", "stream-async")]
    [InlineData("legacy-powerpoint", "path")]
    [InlineData("legacy-powerpoint", "stream")]
    [InlineData("legacy-powerpoint", "path-async")]
    [InlineData("legacy-powerpoint", "stream-async")]
    public async Task EncryptedLoadCapturesOriginalContainerAndPreservesOptions(string format, string route) {
        byte[] source = CreateEncryptedSource(format);
        var policy = new OfficePackageSecurityOptions();
        var settings = new OfficeOpenXmlLoadSettings { MaxCharactersInPart = 1_000_000 };
        string path = Path.Combine(Path.GetTempPath(), "office-source-" + Guid.NewGuid().ToString("N") + Extension(format));
        File.WriteAllBytes(path, source);
        using var stream = new MemoryStream(source, writable: false);
        try {
            if (format == "word") {
                var options = new WordLoadOptions {
                    AccessMode = DocumentAccessMode.ReadOnly, CaptureSourcePackageSecurity = true,
                    PackageSecurity = policy, OpenSettings = settings, OverrideStyles = false
                };
                using WordDocument loaded = route == "path"
                    ? WordDocument.LoadEncrypted(path, Password, options)
                    : WordDocument.LoadEncrypted(stream, Password, options);
                AssertSource(loaded.SourcePackageSecurityReport, loaded.SourcePackageSecurityScope, source);
                AssertOptions(options, policy);
                Assert.Same(settings, options.OpenSettings);
                Assert.False(options.OverrideStyles);
                Assert.Contains(loaded.Paragraphs, paragraph => paragraph.Text == "Source evidence");
            } else if (format == "excel") {
                var options = new ExcelLoadOptions {
                    AccessMode = DocumentAccessMode.ReadOnly, CaptureSourcePackageSecurity = true,
                    PackageSecurity = policy, OpenSettings = settings
                };
                using ExcelDocument loaded = route switch {
                    "path" => ExcelDocument.LoadEncrypted(path, Password, options),
                    "stream" => ExcelDocument.LoadEncrypted(stream, Password, options),
                    "path-async" => await ExcelDocument.LoadEncryptedAsync(path, Password, options),
                    _ => await ExcelDocument.LoadEncryptedAsync(stream, Password, options)
                };
                AssertSource(loaded.SourcePackageSecurityReport, loaded.SourcePackageSecurityScope, source);
                AssertOptions(options, policy);
                Assert.Same(settings, options.OpenSettings);
                Assert.Equal("Evidence", loaded.Sheets[0].Name);
            } else {
                var legacyOptions = new LegacyPptImportOptions { Password = "caller-retained-password", MaxCommentCount = 123 };
                var options = new PowerPointLoadOptions {
                    AccessMode = DocumentAccessMode.ReadOnly, CaptureSourcePackageSecurity = true,
                    PackageSecurity = policy, OpenSettings = settings, LegacyPptImportOptions = legacyOptions
                };
                using PowerPointPresentation loaded = route switch {
                    "path" => PowerPointPresentation.LoadEncrypted(path, Password, options),
                    "stream" => PowerPointPresentation.LoadEncrypted(stream, Password, options),
                    "path-async" => await PowerPointPresentation.LoadEncryptedAsync(path, Password, options),
                    _ => await PowerPointPresentation.LoadEncryptedAsync(stream, Password, options)
                };
                AssertSource(loaded.SourcePackageSecurityReport, loaded.SourcePackageSecurityScope, source);
                AssertOptions(options, policy);
                Assert.Same(settings, options.OpenSettings);
                Assert.Same(legacyOptions, options.LegacyPptImportOptions);
                Assert.Equal("caller-retained-password", legacyOptions.Password);
                Assert.Equal(123, legacyOptions.MaxCommentCount);
                Assert.Single(loaded.Slides);
            }
            Assert.True(stream.CanRead);
            Assert.Equal(source, File.ReadAllBytes(path));
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("word")]
    [InlineData("excel")]
    [InlineData("powerpoint")]
    [InlineData("legacy-powerpoint")]
    public void SourceCaptureRemainsOptInForEncryptedLoads(string format) {
        using var source = new MemoryStream(CreateEncryptedSource(format), writable: false);
        if (format == "word") {
            using WordDocument loaded = WordDocument.LoadEncrypted(source, Password);
            Assert.Null(loaded.SourcePackageSecurityReport);
        } else if (format == "excel") {
            using ExcelDocument loaded = ExcelDocument.LoadEncrypted(source, Password);
            Assert.Null(loaded.SourcePackageSecurityReport);
        } else {
            using PowerPointPresentation loaded = PowerPointPresentation.LoadEncrypted(source, Password);
            Assert.Null(loaded.SourcePackageSecurityReport);
        }
        Assert.True(source.CanRead);
    }

    [Fact]
    public void DecryptedPackagePolicyStillAppliesWhenItsSourceCaptureIsSuppressed() {
        using var encrypted = new MemoryStream();
        using (WordDocument source = WordDocument.Create()) {
            source.AddParagraph("Linked source").AddHyperLink("inert external target", new Uri("https://example.test/evidence"));
            source.SaveEncrypted(encrypted, Password);
        }
        byte[] sourceBytes = encrypted.ToArray();
        var policy = OfficePackageSecurityOptions.UntrustedDefaults;
        Assert.True(OfficePackageSecurityInspector.Inspect(sourceBytes, policy).IsValid);
        encrypted.Position = 0;
        var options = new WordLoadOptions { CaptureSourcePackageSecurity = true, PackageSecurity = policy };

        OfficePackageSecurityException exception = Assert.Throws<OfficePackageSecurityException>(() =>
            WordDocument.LoadEncrypted(encrypted, Password, options));

        Assert.Equal(OfficePackageSecurityRule.ExternalRelationships, exception.Rule);
        AssertOptions(options, policy, DocumentAccessMode.ReadWrite);
        Assert.True(encrypted.CanRead);
    }

    [Theory]
    [InlineData("word")]
    [InlineData("excel")]
    [InlineData("powerpoint")]
    [InlineData("legacy-powerpoint")]
    public void OrdinaryLoadStillCapturesItsOriginalSource(string format) {
        byte[] bytes;
        if (format == "word") {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("Ordinary source");
            bytes = document.ToBytes();
        } else if (format == "excel") {
            using ExcelDocument document = ExcelDocument.Create();
            document.AddWorksheet("Ordinary");
            bytes = document.ToBytes();
        } else {
            using PowerPointPresentation document = PowerPointPresentation.Create();
            document.AddSlide().AddTextBox("Ordinary source", 100000, 100000, 2200000, 600000);
            bytes = document.ToBytes(format == "legacy-powerpoint" ? PowerPointFileFormat.Ppt : PowerPointFileFormat.Pptx);
        }
        using var input = new MemoryStream(bytes, writable: false);
        OfficePackageSecurityReport? report;
        string? scope;
        if (format == "word") {
            using WordDocument loaded = WordDocument.Load(input, new WordLoadOptions { CaptureSourcePackageSecurity = true });
            report = loaded.SourcePackageSecurityReport;
            scope = loaded.SourcePackageSecurityScope;
        } else if (format == "excel") {
            using ExcelDocument loaded = ExcelDocument.Load(input, new ExcelLoadOptions { CaptureSourcePackageSecurity = true });
            report = loaded.SourcePackageSecurityReport;
            scope = loaded.SourcePackageSecurityScope;
        } else {
            using PowerPointPresentation loaded = PowerPointPresentation.Load(input, new PowerPointLoadOptions { CaptureSourcePackageSecurity = true });
            report = loaded.SourcePackageSecurityReport;
            scope = loaded.SourcePackageSecurityScope;
        }
        Assert.NotNull(report);
        Assert.Equal(bytes.LongLength, report!.PackageBytes);
        Assert.Equal(format == "legacy-powerpoint" ? OfficePackageContainerKind.CompoundBinary : OfficePackageContainerKind.OpenXml, report.ContainerKind);
        Assert.Equal("source-container", scope);
        Assert.True(input.CanRead);
    }

    private static byte[] CreateEncryptedSource(string format) {
        using var output = new MemoryStream();
        if (format == "word") {
            using WordDocument source = WordDocument.Create();
            source.AddParagraph("Source evidence");
            source.SaveEncrypted(output, Password);
        } else if (format == "excel") {
            using ExcelDocument source = ExcelDocument.Create();
            source.AddWorksheet("Evidence").CellValue(1, 1, "Source evidence");
            source.SaveEncrypted(output, Password);
        } else {
            using PowerPointPresentation source = PowerPointPresentation.Create();
            source.AddSlide().AddTextBox("Source evidence", 100000, 100000, 2200000, 600000);
            return source.ToEncryptedBytes(Password, format == "legacy-powerpoint" ? PowerPointFileFormat.Ppt : PowerPointFileFormat.Pptx);
        }
        return output.ToArray();
    }

    private static string Extension(string format) => format switch {
        "word" => ".docx", "excel" => ".xlsx", "legacy-powerpoint" => ".ppt", _ => ".pptx"
    };

    private static void AssertSource(OfficePackageSecurityReport? report, string? scope, byte[] source) {
        Assert.NotNull(report);
        Assert.Equal(OfficePackageContainerKind.CompoundBinary, report!.ContainerKind);
        Assert.Equal(source.LongLength, report.PackageBytes);
        Assert.Equal("encrypted-source-container", scope);
    }

    private static void AssertOptions(DocumentLoadOptions options, OfficePackageSecurityOptions policy,
        DocumentAccessMode accessMode = DocumentAccessMode.ReadOnly) {
        Assert.True(options.CaptureSourcePackageSecurity);
        Assert.Same(policy, options.PackageSecurity);
        Assert.Equal(accessMode, options.AccessMode);
        Assert.Equal(DocumentPersistenceMode.Explicit, options.PersistenceMode);
    }
}
