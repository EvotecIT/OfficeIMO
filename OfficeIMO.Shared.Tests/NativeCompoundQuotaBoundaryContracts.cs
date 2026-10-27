using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using OfficeIMO.Core.Internal;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls.Model;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc.Model;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Word;
using OfficeIMO.Reader.Excel;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class NativeCompoundQuotaBoundaryContracts {
    // Word/Excel expose the canonical fixed directory ceiling rather than a smaller public option.
    // One valid, writer-produced empty-stream directory crosses that ceiling before payload parsing.
    private static readonly Lazy<byte[]> DirectoryQuota = new(() => {
        var streams = new List<OfficeCompoundStream>();
        for (int index = 0; index < 65536; index++) streams.Add(new OfficeCompoundStream("Entry" + index.ToString("D5"), Array.Empty<byte>()));
        return OfficeCompoundFileWriter.Write(streams);
    });

    [Theory]
    [InlineData("word-model", "word")]
    [InlineData("word-native", "word")]
    [InlineData("excel-model", "excel")]
    [InlineData("excel-native", "excel")]
    [InlineData("excel-tabular", "excel")]
    [InlineData("powerpoint-native", "powerpoint")]
    public void PublicLoadRoutesRetainCanonicalDirectoryQuota(string route, string owner) {
        byte[] bytes = DirectoryQuota.Value;
        using var source = new MemoryStream(bytes, writable: false);
        OfficeInspectionResourceLimitException error = Assert.Throws<OfficeInspectionResourceLimitException>(() => {
            switch (route) {
                case "word-model": LegacyDocDocument.Load(bytes); break;
                case "word-native": using (WordDocument.Load(source)) { } break;
                case "excel-model": LegacyXlsWorkbook.Load(bytes); break;
                case "excel-native": using (ExcelDocument.Load(source)) { } break;
                case "powerpoint-native": using (PowerPointDocument.Load(source)) { } break;
                default:
                    string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");
                    try {
                        File.WriteAllBytes(path, bytes);
                        using var reader = ExcelDocument.OpenReader(path);
                        _ = reader.SheetNames;
                    } finally { if (File.Exists(path)) File.Delete(path); }
                    break;
            }
        });
        if (route.EndsWith("-native", StringComparison.Ordinal)) Assert.True(source.CanRead);
        Assert.Equal(owner, error.Owner); Assert.Equal("MaxDirectoryEntries", error.LimitName); Assert.Equal(65536, error.Maximum);
        Assert.NotNull(error.InnerException);
        Assert.Equal("OfficeCompoundStreamLimitExceededException", error.InnerException!.GetType().Name);
        Assert.Equal(typeof(OfficeInspectionResourceLimitException).Assembly, error.InnerException.GetType().Assembly);
    }

    [Theory]
    [InlineData("model")]
    [InlineData("native")]
    [InlineData("encrypted")]
    public void PowerPointExistingSmallRecordPolicyPreservesEffectiveCompoundLimit(string route) {
        using PowerPointDocument source = PowerPointDocument.Create();
        source.AddSlide().AddTextBox("Ordinary binary content", 100000, 100000, 2000000, 500000);
        var options = new LegacyPptImportOptions { MaxRecordCount = 3 };
        byte[] bytes = route == "encrypted" ? source.ToBytes(new OfficeIMO.PowerPoint.PowerPointSaveOptions{Password="fixture-password",Format=PowerPointFileFormat.Ppt}): source.ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Ppt });
        OfficeInspectionResourceLimitException error = Assert.Throws<OfficeInspectionResourceLimitException>(() => {
            if (route == "model") { LegacyPptDocument.Load(bytes, options); return; }
            using var input = new MemoryStream(bytes, writable: false);
            var load = new PowerPointLoadOptions { MaxRecordCount = options.MaxRecordCount, Password = route == "encrypted" ? "fixture-password" : null };
            using PowerPointDocument loaded = PowerPointDocument.Load(input, load);
        });
        Assert.Equal("powerpoint", error.Owner); Assert.Equal("MaxDirectoryEntries", error.LimitName); Assert.Equal(3, error.Maximum);
        Assert.Equal("OfficeCompoundStreamLimitExceededException", error.InnerException!.GetType().Name);
        Assert.Equal(3, options.MaxRecordCount);
    }

    [Theory]
    [InlineData("word", "stream")]
    [InlineData("word", "batch")]
    [InlineData("word", "probe")]
    [InlineData("word", "folder")]
    [InlineData("excel", "stream")]
    [InlineData("excel", "batch")]
    [InlineData("excel", "probe")]
    [InlineData("excel", "folder")]
    public async Task LegacyReaderQuotaRemainsTerminalAcrossImportProbeAndAggregate(string format, string route) {
        // A recognized foreign Office root makes PreferContent exercise the legacy extension validator.
        byte[] bytes = OfficeCompoundFileWriter.Write(new[] { new OfficeCompoundStream(format == "word" ? "Workbook" : "WordDocument", new byte[] { 1 }) });
        var limits = new OfficeLegacyImportLimits { MaxCompoundStreams = 3 };
        var builder = new OfficeDocumentReaderBuilder();
        if (format == "word") builder.AddLegacyWordHandler(new global::OfficeIMO.Word.Legacy.LegacyWordImportOptions { Limits = limits });
        else builder.AddLegacySpreadsheetHandler(new global::OfficeIMO.Excel.Legacy.LegacySpreadsheetImportOptions { Limits = limits });
        OfficeDocumentReader reader = builder.Build();
        var options = new ReaderOptions { DetectionMode = route == "probe" ? ReaderDetectionMode.PreferContent : ReaderDetectionMode.ContentWhenUnknown };
        string extension = format == "word" ? ".lwp" : ".qpw";
        ReaderResourceLimitException error = await Assert.ThrowsAsync<ReaderResourceLimitException>(async () => {
            if (route == "stream" || route == "probe") { reader.ReadDocument(bytes, "source" + extension, options); return; }
            string directory = Path.Combine(Path.GetTempPath(), "native-reader-quota-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory); string path = Path.Combine(directory, "source" + extension);
            try {
                File.WriteAllBytes(path, bytes);
                if (route == "batch") await reader.ReadDocumentsDetailedAsync(new[] { path }, options);
                else reader.ReadFolder(directory, options: options).ToArray();
            } finally { File.Delete(path); Directory.Delete(directory); }
        });
        Assert.EndsWith(".MaxDirectoryEntries", error.LimitName, StringComparison.Ordinal); Assert.Equal(3, error.Maximum);
        OfficeInspectionResourceLimitException native = Assert.IsType<OfficeInspectionResourceLimitException>(error.InnerException);
        Assert.Equal("MaxDirectoryEntries", native.LimitName); Assert.Equal(3, native.Maximum); Assert.NotNull(native.InnerException);
        Assert.Equal(3, limits.MaxCompoundStreams);
    }

    [Theory]
    [InlineData("word")]
    [InlineData("excel")]
    public async Task MalformedLegacyReaderInputKeepsPerFileFailureBehavior(string format) {
        var builder = new OfficeDocumentReaderBuilder();
        if (format == "word") builder.AddLegacyWordHandler();
        else builder.AddLegacySpreadsheetHandler();
        OfficeDocumentReader reader = builder.Build();
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + (format == "word" ? ".lwp" : ".qpw"));
        try {
            File.WriteAllBytes(path, new byte[1024]);
            ReaderDocumentReadOutcome outcome = Assert.Single(await reader.ReadDocumentsDetailedAsync(new[] { path }));
            Assert.False(outcome.Succeeded); Assert.Null(outcome.Document);
            Assert.IsType<InvalidDataException>(outcome.Error);
        } finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void CoreLegacyDirectoryHelperPreservesQuotaAndMalformedBehavior() {
        byte[] bytes = OfficeCompoundFileWriter.Write(new[] { new OfficeCompoundStream("Payload", Array.Empty<byte>()) });
        var limits = new OfficeLegacyImportLimits { MaxCompoundStreams = 3 };
        OfficeInspectionResourceLimitException error = Assert.Throws<OfficeInspectionResourceLimitException>(() =>
            OfficeLegacyCompoundInspector.IsValidCompound(bytes, limits));
        Assert.NotEmpty(error.Owner); Assert.Equal("MaxDirectoryEntries", error.LimitName); Assert.Equal(3, error.Maximum);
        Assert.Equal("OfficeCompoundStreamLimitExceededException", error.InnerException!.GetType().Name);
        Assert.Equal(typeof(OfficeInspectionResourceLimitException).Assembly, error.InnerException.GetType().Assembly);
        Assert.False(OfficeLegacyCompoundInspector.IsValidCompound(new byte[1024], limits));
    }

    [Fact]
    public void OrdinaryMalformedLegacyInputsKeepTheirDiagnosticOrMalformedContract() {
        byte[] malformed = new byte[1024];
        Assert.Contains(LegacyDocDocument.Load(malformed).Diagnostics, item => item.Code == "doc-compound-invalid");
        Assert.Contains(LegacyXlsWorkbook.Load(malformed).Diagnostics, item => item.Code == "xls-compound-signature");
        Assert.Throws<InvalidDataException>(() => LegacyPptDocument.Load(malformed));
    }

    [Fact]
    public void OrdinaryFullNativeLegacyLoadsKeepReadableContent() {
        using (WordDocument source = WordDocument.Create()) {
            source.AddParagraph("Ordinary full content");
            using var input = new MemoryStream(source.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc }));
            using WordDocument loaded = WordDocument.Load(input);
            Assert.Contains(loaded.Paragraphs, paragraph => paragraph.Text.Contains("Ordinary full content", StringComparison.Ordinal));
        }
        using (ExcelDocument source = ExcelDocument.Create()) {
            var sheet = source.AddWorksheet("Evidence"); sheet.CellValue(1, 1, "Ordinary full content");
            using var input = new MemoryStream(source.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xls }));
            using ExcelDocument loaded = ExcelDocument.Load(input);
            Assert.True(loaded.Sheets.Single().TryGetCellText(1, 1, out string value));
            Assert.Equal("Ordinary full content", value);
        }
        using (PowerPointDocument source = PowerPointDocument.Create()) {
            source.AddSlide().AddTextBox("Ordinary full content", 100000, 100000, 2000000, 500000);
            using var input = new MemoryStream(source.ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Ppt }));
            using PowerPointDocument loaded = PowerPointDocument.Load(input);
            Assert.Contains(loaded.ExtractMarkdownChunks(), chunk => chunk.Text!.Contains("Ordinary full content", StringComparison.Ordinal));
        }
    }
}
