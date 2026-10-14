using System.IO;
using System.Xml.Linq;
using OfficeIMO;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.Security.Tests;

public sealed class OfficeVbaProjectTests {
    [Fact]
    public void DocumentAdaptersRejectForeignDocumentModulesBeforeAddingAProjectPart() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        var wordProject = OfficeVbaProject.Load(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestData", "Vba", "Word-authored.bin")));
        using var excel = ExcelDocument.Create();
        excel.AddWorksheet("Data");
        Assert.Throws<ArgumentException>(() => excel.SetVbaProject(wordProject));
        Assert.Null(excel.ReadVbaProject());
        using var presentation = PowerPointPresentation.Create();
        Assert.Throws<ArgumentException>(() => presentation.SetVbaProject(wordProject));
        Assert.Null(presentation.ReadVbaProject());
        var excelProject = OfficeVbaProject.Create();
        excelProject.AddDocumentModule("ThisWorkbook", "", new Guid("00020819-0000-0000-C000-000000000046"));
        using var word = WordDocument.Create();
        Assert.Throws<ArgumentException>(() => word.SetVbaProject(excelProject));
        Assert.Null(word.ReadVbaProject());
    }

    [Fact]
    public void ExcelSheetModulesRequireExistingUniqueCodeNamesAndMatchingHostTypes() {
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-vba-binding-" + Guid.NewGuid().ToString("N") + ".xlsm");
        var project = OfficeVbaProject.Create();
        project.AddDocumentModule("DataCode", "", new Guid("00020820-0000-0000-C000-000000000046"));
        try {
            using (var excel = ExcelDocument.Create(path)) {
                excel.AddWorksheet("Display Name");
                Assert.Throws<ArgumentException>(() => excel.SetVbaProject(project));
                Assert.Null(excel.ReadVbaProject());
                excel.Save();
            }
            using (var package = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(path, true)) {
                package.WorkbookPart!.WorksheetParts.Single().Worksheet!.SheetProperties =
                    new DocumentFormat.OpenXml.Spreadsheet.SheetProperties { CodeName = "DataCode" };
            }
            using (var excel = ExcelDocument.Load(path)) {
                excel.SetVbaProject(project);
                Assert.Equal("DataCode", Assert.Single(excel.ReadVbaProject()!.Modules).Name);
                var wrongType = OfficeVbaProject.Create();
                wrongType.AddDocumentModule("DataCode", "", new Guid("00020821-0000-0000-C000-000000000046"));
                Assert.Throws<ArgumentException>(() => excel.SetVbaProject(wrongType));
                Assert.Contains("00020820", excel.ReadVbaProject()!.GetModule("DataCode").Source);
            }
        } finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void UnchangedWritesEnforceExpandedBudgetAndRetainBytesWhenWithinIt() {
        var project = OfficeVbaProject.Create();
        Assert.Throws<InvalidDataException>(() => project.Write(new OfficeVbaWriteOptions { MaximumExpandedBytes = 1 }));
        project.AddModule("Helpers", "'" + new string('x', 8000) + "\r\n");
        byte[] bytes = project.Write().GetBytes();
        var loaded = OfficeVbaProject.Load(bytes);
        Assert.False(loaded.HasChanges);
        Assert.Throws<InvalidDataException>(() => loaded.Write(new OfficeVbaWriteOptions { MaximumExpandedBytes = 1000 }));
        Assert.Equal(bytes, loaded.Write().GetBytes());
        using var document = WordDocument.Create();
        Assert.Throws<InvalidDataException>(() => document.SetVbaProject(loaded, new OfficeVbaWriteOptions { MaximumExpandedBytes = 1000 }));
        Assert.Null(document.ReadVbaProject());
    }

    [Fact]
    public void ExactSourceImportBudgetAllowsEmptyTrailingFilesAndRejectsNonemptyOnesAtomically() {
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-vba-budget-" + Guid.NewGuid().ToString("N"));
        var project = OfficeVbaProject.Create();
        project.AddModule("Helpers", "'original\r\n");
        project.AddModule("Empty", "");
        project = OfficeVbaProject.Load(project.Write().GetBytes());
        try {
            project.ExportSources(folder);
            File.WriteAllText(Path.Combine(folder, "module-0001.bas"), "'changed\r\n", new UTF8Encoding(false));
            File.WriteAllText(Path.Combine(folder, "module-0002.bas"), "");
            int budget = checked((int)new FileInfo(Path.Combine(folder, "module-0001.bas")).Length);
            project.ImportSources(folder, budget);
            Assert.Contains("'changed", project.GetModule("Helpers").Source);
            project = OfficeVbaProject.Load(project.Write().GetBytes());
            File.WriteAllText(Path.Combine(folder, "module-0001.bas"), "'another\r\n", new UTF8Encoding(false));
            File.WriteAllText(Path.Combine(folder, "module-0002.bas"), "x");
            budget = checked((int)new FileInfo(Path.Combine(folder, "module-0001.bas")).Length);
            Assert.Throws<InvalidDataException>(() => project.ImportSources(folder, budget));
            Assert.False(project.HasChanges);
        } finally { if (Directory.Exists(folder)) Directory.Delete(folder, true); }
    }

    [Fact]
    public void WordAuthoredDocumentIdentityAndSourceRemainEditable() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        byte[] bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestData", "Vba", "Word-authored.bin"));
        OfficeVbaProject project = OfficeVbaProject.Load(bytes);
        Assert.Equal(OfficeVbaModuleKind.Document, project.GetModule("ThisDocument").Kind);
        Assert.Contains("1Normal.ThisDocument", project.GetModule("ThisDocument").Source);
        Assert.Equal(bytes, project.Write().GetBytes());
        project.SetModuleSource("Module1", "Public Function NativeSum() As Long\r\nNativeSum = 73\r\nEnd Function\r\n");
        OfficeVbaProject changed = OfficeVbaProject.Load(project.Write().GetBytes());
        Assert.Equal(project.GetModule("ThisDocument").Source, changed.GetModule("ThisDocument").Source);
        Assert.Contains("NativeSum = 73", changed.GetModule("Module1").Source);
    }

    [Theory]
    [InlineData(40)]
    [InlineData(50)]
    [InlineData(55)]
    [InlineData(65)]
    [InlineData(130)]
    public void IncompressibleSourceIsPreservedWithoutRawChunksOrPadding(int lines) {
        var random = new Random(432);
        string source = string.Join("\r\n", Enumerable.Range(0, lines).Select(_ => "'" + new string(
            Enumerable.Range(0, 70).Select(__ => (char)random.Next(33, 123)).ToArray()))) + "\r\nPublic Function Value() As Long\r\nValue = 42\r\nEnd Function\r\n";
        var project = OfficeVbaProject.Create();
        project.AddModule("Helpers", source);
        byte[] bytes = project.Write().GetBytes();
        var loaded = OfficeVbaProject.Load(bytes);
        Assert.Equal(project.GetModule("Helpers").Source, loaded.GetModule("Helpers").Source);
        Assert.True(OfficeIMO.Core.Internal.OfficeCompoundFileReader.TryRead(bytes, out OfficeIMO.Core.Internal.OfficeCompoundFile? compound, out _));
        byte[] stream = compound!.Streams["VBA/Helpers"];
        for (int position = 1; position < stream.Length;) {
            int header = stream[position] | stream[position + 1] << 8;
            Assert.NotEqual(0, header & 0x8000);
            position += (header & 0x0fff) + 3;
            Assert.True(position <= stream.Length);
        }
        Assert.Equal(bytes, loaded.Write().GetBytes());
    }

    [Fact]
    public void RenamingMetadataNamedModuleDoesNotRewriteProjectName() {
        var project = OfficeVbaProject.Create("Automation");
        project.AddModule("Name", "'source\r\n");
        project = OfficeVbaProject.Load(project.Write().GetBytes());
        project.RenameModule("Name", "Tools");
        byte[] bytes = project.Write().GetBytes();
        Assert.True(OfficeIMO.Core.Internal.OfficeCompoundFileReader.TryRead(bytes, out OfficeIMO.Core.Internal.OfficeCompoundFile? compound, out _));
        string text = OfficeIMO.Core.Internal.OfficeVbaText.Decode(compound!.Streams["PROJECT"], 1252);
        Assert.Contains("Name=\"Automation\"", text);
        Assert.Contains("Module=Tools", text);
    }

    [Fact]
    public void OfficeAuthoredClassAndHostModulesPreserveMetadataAcrossSourceEdits() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        byte[] bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestData", "Vba", "Excel-authored.bin"));
        OfficeVbaProject project = OfficeVbaProject.Load(bytes);
        Assert.Equal(1250, project.CodePage);
        Assert.Equal(OfficeVbaModuleKind.Class, project.GetModule("Class1").Kind);
        Assert.Equal(OfficeVbaModuleKind.Document, project.GetModule("ThisWorkbook").Kind);
        Assert.Equal(bytes, project.Write().GetBytes());
        project.SetModuleSource("Helpers", "Public Function NativeSum() As Long\r\nNativeSum = 73\r\nEnd Function\r\n");
        project.RenameModule("Helpers", "Tools");
        project.AddModule("NewClass", "Public Count As Long\r\n", OfficeVbaModuleKind.Class);
        OfficeVbaProject edited = OfficeVbaProject.Load(project.Write().GetBytes());
        Assert.Equal(project.GetModule("Class1").Source, edited.GetModule("Class1").Source);
        Assert.Equal(project.GetModule("Sheet1").Source, edited.GetModule("Sheet1").Source);
        Assert.Contains("NativeSum = 73", edited.GetModule("Tools").Source);
        Assert.Contains("FCFB3D2A-A0FA-1068-A738-08002B3371B5", edited.GetModule("NewClass").Source);
        Assert.Throws<InvalidDataException>(() => edited.SetModuleSource("Sheet1", "Attribute VB_Base = \"changed\"\r\n"));
        Assert.False(edited.HasChanges);
    }

    [Fact]
    public void NativeCreationEditsModulesAndReferencesWithoutLosingHostIdentity() {
        var project = OfficeVbaProject.Create("Automation");
        project.AddModule("Helpers", "Public Function Value() As Long\r\nValue = 42\r\nEnd Function\r\n");
        project.AddModule("Counter", "Public Count As Long\r\n", OfficeVbaModuleKind.Class);
        project.AddDocumentModule("ThisWorkbook", "Option Explicit\r\n", new Guid("00020819-0000-0000-C000-000000000046"));
        Guid library = new("2DF8D04C-5BFA-101B-BDE5-00AA0044DE52");
        project.AddRegisteredReference("Office", library, 2, 8);
        OfficeVbaWriteResult initial = project.Write();
        OfficeVbaProject loaded = OfficeVbaProject.Load(initial.GetBytes());
        Assert.Equal("Automation", loaded.Name);
        Assert.False(loaded.IsProtected);
        Assert.Equal(3, loaded.Modules.Count);
        Assert.Equal(OfficeVbaModuleKind.Class, loaded.GetModule("Counter").Kind);
        Assert.Contains("Attribute VB_Creatable = False", loaded.GetModule("Counter").Source);
        loaded.SetModuleSource("ThisWorkbook", "Option Explicit\r\nPublic HostValue As Long\r\n");
        loaded.RenameModule("Helpers", "Tools");
        loaded.SetModuleSource("Tools", "Public Function Value() As Long\r\nValue = 73\r\nEnd Function\r\n");
        loaded.DeleteModule("Counter");
        loaded.AddRegisteredReference("Office", library, 2, 8);
        Assert.Single(loaded.References);
        OfficeVbaProject changed = OfficeVbaProject.Load(loaded.Write().GetBytes());
        Assert.Equal(new[] { "Tools", "ThisWorkbook" }, changed.Modules.Select(module => module.Name));
        Assert.Contains("Value = 73", changed.GetModule("Tools").Source);
        Assert.Contains("00020819-0000-0000-C000-000000000046", changed.GetModule("ThisWorkbook").Source);
        Assert.True(changed.RemoveReference("Office"));
        Assert.Empty(OfficeVbaProject.Load(changed.Write().GetBytes()).References);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(3640)]
    [InlineData(3641)]
    [InlineData(4095)]
    [InlineData(4096)]
    [InlineData(8193)]
    public void NativeWriterHandlesCompressedChunkBoundariesAndStrictEncoding(int characters) {
        var project = OfficeVbaProject.Create();
        string body = "'" + new string('x', characters) + " café €\r\n";
        project.AddModule("Payload", body);
        byte[] bytes = project.Write().GetBytes();
        OfficeVbaProject loaded = OfficeVbaProject.Load(bytes);
        Assert.EndsWith(body, loaded.GetModule("Payload").Source);
        Assert.False(loaded.HasChanges);
        Assert.Equal(bytes, loaded.Write().GetBytes());
        Assert.Throws<EncoderFallbackException>(() => loaded.SetModuleSource("Payload", "'雪"));
        Assert.False(loaded.HasChanges);
    }

    [Fact]
    public void LimitsAndInfrastructureCollisionsAreRejectedWithoutChangingTheProject() {
        var project = OfficeVbaProject.Create();
        project.AddModule("Helpers", new string('x', 5000));
        byte[] bytes = project.Write().GetBytes();
        Assert.Throws<InvalidDataException>(() => OfficeVbaProject.Load(bytes, new OfficeVbaReadOptions { MaximumExpandedBytes = 1000 }));
        Assert.Throws<InvalidDataException>(() => project.Write(new OfficeVbaWriteOptions { MaximumExpandedBytes = 1000 }));
        Assert.Throws<ArgumentException>(() => project.AddModule("dir", ""));
        Assert.Throws<ArgumentException>(() => project.RenameModule("Helpers", "dir"));
        Assert.Equal("Helpers", Assert.Single(project.Modules).Name);
    }

    [Fact]
    public void SourceImportValidatesEveryFileBeforeApplyingChanges() {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO-vba-source-" + Guid.NewGuid().ToString("N"));
        var project = OfficeVbaProject.Create();
        project.AddModule("Helpers", "'original\r\n");
        project.AddModule("Counter", "Public Count As Long\r\n", OfficeVbaModuleKind.Class);
        project = OfficeVbaProject.Load(project.Write().GetBytes());
        try {
            project.ExportSources(directory);
            File.WriteAllText(Path.Combine(directory, "module-0001.bas"), "'changed\r\n");
            File.WriteAllBytes(Path.Combine(directory, "module-0002.cls"), new byte[] { 0xff, 0xff });
            Assert.Throws<DecoderFallbackException>(() => project.ImportSources(directory));
            Assert.False(project.HasChanges);
            File.WriteAllText(Path.Combine(directory, "module-0002.cls"), "Public Count As Long\r\n");
            project.ImportSources(directory);
            var loaded = OfficeVbaProject.Load(project.Write().GetBytes());
            Assert.Contains("'changed", loaded.GetModule("Helpers").Source);
            Assert.Contains("Attribute VB_Creatable = False", loaded.GetModule("Counter").Source);
            XDocument manifest = XDocument.Load(Path.Combine(directory, "vba-project.xml"));
            manifest.Root!.Element("module")!.SetAttributeValue("file", "../outside.bas");
            manifest.Save(Path.Combine(directory, "vba-project.xml"));
            Assert.Throws<InvalidDataException>(() => loaded.ImportSources(directory));
            Assert.False(loaded.HasChanges);
        } finally {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Theory]
    [InlineData("xlsm")]
    [InlineData("docm")]
    [InlineData("pptm")]
    public void DocumentAdaptersSaveAndReopenNativeSource(string extension) {
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-vba-host-" + Guid.NewGuid().ToString("N") + "." + extension);
        var project = OfficeVbaProject.Create("Automation");
        project.AddModule("Helpers", "Public Function Value() As Long\r\nValue = 42\r\nEnd Function\r\n");
        try {
            switch (extension) {
                case "xlsm":
                    using (ExcelDocument document = ExcelDocument.Create(path)) { document.AddWorksheet("Data"); document.SetVbaProject(project); document.Save(); }
                    using (ExcelDocument loaded = ExcelDocument.Load(path)) Assert.Contains("Value = 42", loaded.ReadVbaProject()!.GetModule("Helpers").Source);
                    break;
                case "docm":
                    using (WordDocument document = WordDocument.Create(path)) { document.AddParagraph("VBA source preservation sentinel"); document.SetVbaProject(project); document.Save(); }
                    using (WordDocument loaded = WordDocument.Load(path)) Assert.Contains("Value = 42", loaded.ReadVbaProject()!.GetModule("Helpers").Source);
                    break;
                case "pptm":
                    using (PowerPointPresentation document = PowerPointPresentation.Create(path)) { document.SetVbaProject(project); document.Save(); }
                    using (PowerPointPresentation loaded = PowerPointPresentation.Load(path)) Assert.Contains("Value = 42", loaded.ReadVbaProject()!.GetModule("Helpers").Source);
                    break;
            }
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
