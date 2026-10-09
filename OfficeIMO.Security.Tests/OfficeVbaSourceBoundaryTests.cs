using System.IO;
using OfficeIMO.Core.Internal;
using OfficeIMO.Excel;
using OfficeIMO.Word;
using OfficeIMO.PowerPoint;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Security.Tests;

public sealed class OfficeVbaSourceBoundaryTests {
    [Fact]
    public void DecoderRejectsLiteralAfterCopyFillsTheChunk() {
        byte[] malformed = { 1, 4, 0xb0, 2, 0x41, 0xfc, 0x0f, 0x42 };
        Assert.False(OfficeVbaCompression.TryDecompress(malformed, 10000, out _, out string detail));
        Assert.Contains("4096", detail);
        byte[] boundary = { 1, 3, 0xb0, 2, 0x41, 0xfc, 0x0f };
        Assert.True(OfficeVbaCompression.TryDecompress(boundary, 10000, out byte[] expanded, out _));
        Assert.Equal(4096, expanded.Length); Assert.All(expanded, value => Assert.Equal((byte)'A', value));
    }

    [Fact]
    public void SourceReadLimitExcludesExpandedSignatureReferenceTranscripts() {
        var project = OfficeVbaProject.Create();
        project.AddRegisteredReference("Library", new Guid("11111111-2222-3333-4444-555555555555"), path: new string('x', 4000));
        byte[] bytes = project.Write().GetBytes();
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out _));
        Assert.True(OfficeVbaCompression.TryDecompress(compound!.Streams["VBA/dir"], 20000, out byte[] directory, out _));
        var loaded = OfficeVbaProject.Load(bytes, new OfficeVbaReadOptions { MaximumExpandedBytes = directory.Length });
        Assert.Single(loaded.References); Assert.Equal(bytes, loaded.Write(new OfficeVbaWriteOptions { MaximumExpandedBytes = directory.Length }).GetBytes());
        Assert.Throws<InvalidDataException>(() => OfficeVbaProject.Load(bytes, new OfficeVbaReadOptions { MaximumExpandedBytes = directory.Length - 1 }));
    }

    [Theory]
    [InlineData("Attribute VB_Base")]
    [InlineData("  attribute\tVB_Base")]
    public void DuplicateHostAttributesRejectNewAndReplacementSourceBeforeMutation(string secondAttribute) {
        var project = OfficeVbaProject.Create();
        var module = project.AddDocumentModule("ThisWorkbook", "", new Guid("00020819-0000-0000-C000-000000000046"));
        string original = module.Source;
        string duplicate = original + "\r\n" + secondAttribute + " = \"0{00020820-0000-0000-C000-000000000046}\"\r\n";
        Assert.Throws<InvalidDataException>(() => project.SetModuleSource(module.Name, duplicate));
        Assert.Equal(original, module.Source);
        Assert.Throws<InvalidDataException>(() => project.AddModule("Injected", duplicate, OfficeVbaModuleKind.Class));
        Assert.Single(project.Modules);
        Assert.Throws<InvalidDataException>(() => OfficeVbaText.GetBaseIdentity(duplicate));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExcelWorkbookCodeNameCannotCollideWithWorksheetOrChartsheet(bool chartsheet) {
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-vba-code-name-" + Guid.NewGuid().ToString("N") + ".xlsm");
        try {
            using (var created = ExcelDocument.Create(path)) { created.AddWorksheet("Data"); created.Save(); }
            using (var package = SpreadsheetDocument.Open(path, true)) {
                var workbookPart = package.WorkbookPart!;
                workbookPart.WorksheetParts.Single().Worksheet!.SheetProperties = new SheetProperties { CodeName = chartsheet ? "DataCode" : "SheetCode" };
                if (chartsheet) {
                    var chart = workbookPart.AddNewPart<ChartsheetPart>();
                    chart.Chartsheet = new Chartsheet(new ChartSheetProperties { CodeName = "SheetCode" });
                }
            }
            using var excel = ExcelDocument.Load(path);
            var project = OfficeVbaProject.Create();
            project.AddDocumentModule("SheetCode", "", new Guid("00020819-0000-0000-C000-000000000046"));
            Assert.Throws<ArgumentException>(() => excel.SetVbaProject(project)); Assert.Null(excel.ReadVbaProject());
        } finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void HostAdaptersCanReplaceLargerProjectsWithAnOutputBoundedProject() {
        var large = OfficeVbaProject.Create(); large.AddModule("LongModule", "'" + new string('x', 20000));
        var small = OfficeVbaProject.Create(); small.AddModule("Helpers", "'small");
        int limit = small.Write().GetBytes().Length;
        // Opaque bytes distinguish complete CFB sizes even when module compression is small.
        byte[] largeBytes = large.Write().GetBytes();
        Assert.True(OfficeCompoundFileReader.TryRead(largeBytes, out OfficeCompoundFile? compound, out _));
        largeBytes = OfficeCompoundFileWriter.Rewrite(compound!, new Dictionary<string, byte[]> { ["Opaque"] = new byte[limit + 1000] });
        large = OfficeVbaProject.Load(largeBytes);
        using var excel = ExcelDocument.Create(); excel.AddWorksheet("Data"); excel.SetVbaProject(large);
        excel.SetVbaProject(small, new OfficeVbaWriteOptions { MaximumProjectBytes = limit });
        Assert.Equal("Helpers", Assert.Single(excel.ReadVbaProject()!.Modules).Name);
        using var presentation = PowerPointPresentation.Create(); presentation.SetVbaProject(large);
        presentation.SetVbaProject(small, new OfficeVbaWriteOptions { MaximumProjectBytes = limit });
        Assert.Equal("Helpers", Assert.Single(presentation.ReadVbaProject()!.Modules).Name);
        using var word = WordDocument.Create(); word.SetVbaProject(large);
        var prepared = OfficeVbaProject.Create(); prepared.AddDocumentModuleWithIdentity("ThisDocument", "", "1Normal.ThisDocument"); prepared.AddModule("Helpers", "'small");
        byte[] preparedBytes = prepared.Write().GetBytes();
        word.SetVbaProject(prepared, new OfficeVbaWriteOptions { MaximumProjectBytes = preparedBytes.Length });
        Assert.Contains(word.ReadVbaProject()!.Modules, module => module.Name == "Helpers");
    }
}
