using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Core.Internal;
using OfficeIMO.Excel;

namespace OfficeIMO.Security.Tests;

public sealed class OfficeVbaIdentityBoundaryTests {
    [Theory]
    [InlineData(OfficeVbaModuleKind.Standard, "worksheet")]
    [InlineData(OfficeVbaModuleKind.Standard, "chartsheet")]
    [InlineData(OfficeVbaModuleKind.Standard, "workbook")]
    [InlineData(OfficeVbaModuleKind.Class, "worksheet")]
    [InlineData(OfficeVbaModuleKind.Class, "chartsheet")]
    [InlineData(OfficeVbaModuleKind.Class, "workbook")]
    [InlineData(OfficeVbaModuleKind.Designer, "worksheet")]
    [InlineData(OfficeVbaModuleKind.Designer, "chartsheet")]
    public void ExcelRejectsNonHostModuleCodeNameCollisionsBeforeReplacingTheCarrier(OfficeVbaModuleKind kind, string host) {
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-vba-identity-" + Guid.NewGuid().ToString("N") + ".xlsm");
        try {
            var original = OfficeVbaProject.Create(); original.AddModule("Original", "'original");
            byte[] originalBytes = original.Write().GetBytes();
            var candidate = kind == OfficeVbaModuleKind.Designer
                ? OfficeVbaProject.Load(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestData", "Vba", "Excel-nested-form.bin")))
                : OfficeVbaProject.Create();
            string collision = kind == OfficeVbaModuleKind.Designer ? "FrmNested" : "HostCode";
            if (kind != OfficeVbaModuleKind.Designer) candidate.AddModule(collision, "'replacement", kind);
            using (var excel = ExcelDocument.Create(path)) { excel.AddWorksheet("Data"); excel.AddWorksheet("Collision"); excel.SetVbaProject(original); excel.Save(); }
            string workbookXml;
            string sheetXml;
            using (var package = SpreadsheetDocument.Open(path, true)) {
                var workbook = package.WorkbookPart!;
                workbook.Workbook!.WorkbookProperties = new WorkbookProperties { CodeName = host == "workbook" ? collision.ToLowerInvariant() : "ThisWorkbook" };
                // Keep the native form fixture's Sheet1 document binding valid so the
                // designer collision is the only reason for rejecting the candidate.
                workbook.WorksheetParts.First().Worksheet!.SheetProperties = new SheetProperties { CodeName = "Sheet1" };
                workbook.WorksheetParts.Last().Worksheet!.SheetProperties = new SheetProperties { CodeName = host == "worksheet" ? collision.ToLowerInvariant() : "SheetCode" };
                if (host == "chartsheet") {
                    var chart = workbook.AddNewPart<ChartsheetPart>();
                    chart.Chartsheet = new Chartsheet(new ChartSheetProperties { CodeName = collision.ToLowerInvariant() });
                }
                workbookXml = workbook.Workbook.OuterXml;
                sheetXml = workbook.WorksheetParts.Last().Worksheet!.OuterXml;
            }
            using (var excel = ExcelDocument.Load(path)) {
                Assert.Throws<ArgumentException>(() => excel.SetVbaProject(candidate));
                Assert.Equal(originalBytes, excel.ReadVbaProject()!.Write().GetBytes());
                excel.Save();
            }
            using var saved = SpreadsheetDocument.Open(path, false);
            Assert.Equal(DocumentFormat.OpenXml.SpreadsheetDocumentType.MacroEnabledWorkbook, saved.DocumentType);
            Assert.Equal(workbookXml, saved.WorkbookPart!.Workbook!.OuterXml);
            Assert.Equal(sheetXml, saved.WorkbookPart.WorksheetParts.Last().Worksheet!.OuterXml);
            if (host == "chartsheet") Assert.Equal(collision.ToLowerInvariant(), saved.WorkbookPart.ChartsheetParts.Single().Chartsheet!.ChartSheetProperties!.CodeName!.Value);
        } finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Theory]
    [InlineData(OfficeVbaModuleKind.Standard)]
    [InlineData(OfficeVbaModuleKind.Class)]
    [InlineData(OfficeVbaModuleKind.Document)]
    public void AddingAModuleRejectsAPreservedStorageIdentityWithoutChangingTheProject(OfficeVbaModuleKind kind) {
        byte[] bytes = ProjectWithOpaqueStorage();
        var project = OfficeVbaProject.Load(bytes);
        Assert.Throws<ArgumentException>(() => {
            if (kind == OfficeVbaModuleKind.Document) project.AddDocumentModule("fOo", "", new Guid("00020819-0000-0000-C000-000000000046"));
            else project.AddModule("fOo", "", kind);
        });
        Assert.False(project.HasChanges);
        Assert.Equal(bytes, project.Write().GetBytes());
    }

    [Fact]
    public void RenameRejectsAPreservedStorageIdentityWithoutChangingTheProject() {
        byte[] bytes = ProjectWithOpaqueStorage();
        var project = OfficeVbaProject.Load(bytes);
        Assert.Throws<ArgumentException>(() => project.RenameModule("Stable", "fOo"));
        Assert.False(project.HasChanges);
        Assert.Equal(bytes, project.Write().GetBytes());
    }

    [Fact]
    public void ImportValidatesStorageIdentitiesBeforeApplyingEarlierSourceChanges() {
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-vba-import-identity-" + Guid.NewGuid().ToString("N"));
        try {
            byte[] bytes = ProjectWithOpaqueStorage();
            var project = OfficeVbaProject.Load(bytes);
            var donor = OfficeVbaProject.Create(); donor.AddModule("Stable", "'changed"); donor.AddModule("fOo", "'new");
            donor.ExportSources(path);
            Assert.Throws<ArgumentException>(() => project.ImportSources(path));
            Assert.False(project.HasChanges);
            Assert.Equal(bytes, project.Write().GetBytes());
        } finally { if (Directory.Exists(path)) Directory.Delete(path, true); }
    }

    [Theory]
    [InlineData(-1, 0, "majorVersion")]
    [InlineData(65536, 0, "majorVersion")]
    [InlineData(0, -1, "minorVersion")]
    [InlineData(0, 65536, "minorVersion")]
    public void RegisteredReferenceVersionsRejectInvalidFieldsBeforeMutation(int major, int minor, string parameter) {
        byte[] bytes = OfficeVbaProject.Create().Write().GetBytes();
        var project = OfficeVbaProject.Load(bytes);
        var error = Assert.Throws<ArgumentOutOfRangeException>(() => project.AddRegisteredReference("Library", Guid.NewGuid(), major, minor));
        Assert.Equal(parameter, error.ParamName);
        Assert.Empty(project.References);
        Assert.False(project.HasChanges);
        Assert.Equal(bytes, project.Write().GetBytes());
    }

    [Fact]
    public void RegisteredReferenceVersionBoundsSerializeWithinTheNativeGrammar() {
        var project = OfficeVbaProject.Create();
        project.AddRegisteredReference("Maximum", Guid.NewGuid(), ushort.MaxValue, ushort.MaxValue);
        project.AddRegisteredReference("Minimum", Guid.NewGuid(), 0, 0);
        var loaded = OfficeVbaProject.Load(project.Write().GetBytes());
        Assert.Contains("#ffff.ffff#", loaded.References.Single(reference => reference.Name == "Maximum").LibraryId);
        Assert.Contains("#0.0#", loaded.References.Single(reference => reference.Name == "Minimum").LibraryId);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RegisteredReferenceIdentityUsesTheLibraryGuidInsteadOfItsPath(bool reload) {
        Guid first = new("11111111-2222-3333-4444-555555555555");
        Guid second = new("22222222-3333-4444-5555-666666666666");
        var project = OfficeVbaProject.Create();
        project.AddRegisteredReference("First", first, path: "C:\\Libraries\\" + second.ToString("B") + "\\first.tlb");
        if (reload) project = OfficeVbaProject.Load(project.Write().GetBytes());
        project.AddRegisteredReference("Second", second);
        Assert.Equal(new[] { "First", "Second" }, project.References.Select(reference => reference.Name));
        byte[] bytes = project.Write().GetBytes();
        project = OfficeVbaProject.Load(bytes);
        Assert.Equal(new[] { "First", "Second" }, project.References.Select(reference => reference.Name));
        project.AddRegisteredReference("AlreadyPresent", first, path: "C:\\another.tlb");
        Assert.False(project.HasChanges);
        Assert.Equal(bytes, project.Write().GetBytes());
    }

    private static byte[] ProjectWithOpaqueStorage() {
        var project = OfficeVbaProject.Create(); project.AddModule("Stable", "'original");
        Assert.True(OfficeCompoundFileReader.TryRead(project.Write().GetBytes(), out OfficeCompoundFile? compound, out _));
        var storage = new OfficeCompoundFileEntry("Foo", "VBA/Foo", 1, 0, classId: Guid.NewGuid(), stateBits: 7, creationTime: 123, modifiedTime: 456);
        var preserved = new OfficeCompoundFile(compound!.Streams, compound.Entries.Concat(new[] { storage }).ToArray(), compound.RootEntry);
        byte[] bytes = OfficeCompoundFileWriter.Rewrite(preserved, new Dictionary<string, byte[]>());
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? loaded, out _));
        Assert.Contains(loaded!.Entries, entry => entry.IsStorage && entry.Path == "VBA/Foo" && entry.ClassId == storage.ClassId);
        return bytes;
    }
}
