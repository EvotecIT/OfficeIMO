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
                    Password = Password, AccessMode = DocumentAccessMode.ReadOnly, CaptureSourcePackageSecurity = true,
                    PackageSecurity = policy, OpenSettings = settings, OverrideStyles = false
                };
                using WordDocument loaded = route == "path"
                    ? WordDocument.Load(path, options)
                    : WordDocument.Load(stream, options);
                AssertSource(loaded.SourcePackageSecurityReport, loaded.SourcePackageSecurityScope, source);
                AssertOptions(options, policy);
                Assert.Same(settings, options.OpenSettings);
                Assert.False(options.OverrideStyles);
                Assert.Contains(loaded.Paragraphs, paragraph => paragraph.Text == "Source evidence");
            } else if (format == "excel") {
                var options = new ExcelLoadOptions {
                    Password = Password, AccessMode = DocumentAccessMode.ReadOnly, CaptureSourcePackageSecurity = true,
                    PackageSecurity = policy, OpenSettings = settings
                };
                using ExcelDocument loaded = route switch {
                    "path" => ExcelDocument.Load(path, options),
                    "stream" => ExcelDocument.Load(stream, options),
                    "path-async" => await ExcelDocument.LoadAsync(path, options),
                    _ => await ExcelDocument.LoadAsync(stream, options)
                };
                AssertSource(loaded.SourcePackageSecurityReport, loaded.SourcePackageSecurityScope, source);
                AssertOptions(options, policy);
                Assert.Same(settings, options.OpenSettings);
                Assert.Equal("Evidence", loaded.Sheets[0].Name);
            } else {
                                var options = new PowerPointLoadOptions {
                    Password = Password, AccessMode = DocumentAccessMode.ReadOnly, CaptureSourcePackageSecurity = true,
                    PackageSecurity = policy, OpenSettings = settings, MaxCommentCount = 123
                };
                using PowerPointDocument loaded = route switch {
                    "path" => PowerPointDocument.Load(path, options),
                    "stream" => PowerPointDocument.Load(stream, options),
                    "path-async" => await PowerPointDocument.LoadAsync(path, options),
                    _ => await PowerPointDocument.LoadAsync(stream, options)
                };
                AssertSource(loaded.SourcePackageSecurityReport, loaded.SourcePackageSecurityScope, source);
                AssertOptions(options, policy);
                Assert.Same(settings, options.OpenSettings);
                Assert.Equal(Password, options.Password);
                Assert.Equal(123, options.MaxCommentCount);
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
            using WordDocument loaded = WordDocument.Load(source,new OfficeIMO.Word.WordLoadOptions{Password=Password});
            Assert.Null(loaded.SourcePackageSecurityReport);
        } else if (format == "excel") {
            using ExcelDocument loaded = ExcelDocument.Load(source,new OfficeIMO.Excel.ExcelLoadOptions{Password=Password});
            Assert.Null(loaded.SourcePackageSecurityReport);
        } else {
            using PowerPointDocument loaded = PowerPointDocument.Load(source,new OfficeIMO.PowerPoint.PowerPointLoadOptions{Password=Password});
            Assert.Null(loaded.SourcePackageSecurityReport);
        }
        Assert.True(source.CanRead);
    }

    [Fact]
    public void DecryptedPackagePolicyStillAppliesWhenItsSourceCaptureIsSuppressed() {
        using var encrypted = new MemoryStream();
        using (WordDocument source = WordDocument.Create()) {
            source.AddParagraph("Linked source").AddHyperLink("inert external target", new Uri("https://example.test/evidence"));
            source.Save(encrypted,new OfficeIMO.Word.WordSaveOptions{Password=Password});
        }
        byte[] sourceBytes = encrypted.ToArray();
        var policy = OfficePackageSecurityOptions.UntrustedDefaults;
        Assert.True(OfficePackageSecurityInspector.Inspect(sourceBytes, policy).IsValid);
        encrypted.Position = 0;
        var options = new WordLoadOptions { Password = Password, CaptureSourcePackageSecurity = true, PackageSecurity = policy };

        OfficePackageSecurityException exception = Assert.Throws<OfficePackageSecurityException>(() =>
            WordDocument.Load(encrypted, options));

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
            using PowerPointDocument document = PowerPointDocument.Create();
            document.AddSlide().AddTextBox("Ordinary source", 100000, 100000, 2200000, 600000);
            bytes = document.ToBytes(new OfficeIMO.PowerPoint.PowerPointSaveOptions{Format=format == "legacy-powerpoint" ? PowerPointFileFormat.Ppt : PowerPointFileFormat.Pptx});
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
            using PowerPointDocument loaded = PowerPointDocument.Load(input, new PowerPointLoadOptions { CaptureSourcePackageSecurity = true });
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
            source.Save(output,new OfficeIMO.Word.WordSaveOptions{Password=Password});
        } else if (format == "excel") {
            using ExcelDocument source = ExcelDocument.Create();
            source.AddWorksheet("Evidence").CellValue(1, 1, "Source evidence");
            source.Save(output,new OfficeIMO.Excel.ExcelSaveOptions{Password=Password});
        } else {
            using PowerPointDocument source = PowerPointDocument.Create();
            source.AddSlide().AddTextBox("Source evidence", 100000, 100000, 2200000, 600000);
            return source.ToBytes(new OfficeIMO.PowerPoint.PowerPointSaveOptions{Password=Password,Format=format == "legacy-powerpoint" ? PowerPointFileFormat.Ppt : PowerPointFileFormat.Pptx});
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
