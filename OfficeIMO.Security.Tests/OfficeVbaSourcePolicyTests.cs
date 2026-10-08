using System.IO;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO;
using OfficeIMO.Core.Internal;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.Security.Tests;

public sealed class OfficeVbaSourcePolicyTests {
    [Fact]
    public void EditsPreserveOpaqueStreamsAndUnchangedModuleBytes() {
        var project = OfficeVbaProject.Create();
        project.AddModule("Helpers", "'initial\r\n");
        project.AddModule("Counter", "Public Count As Long\r\n", OfficeVbaModuleKind.Class);
        byte[] original = project.Write().GetBytes();
        Assert.True(OfficeCompoundFileReader.TryRead(original, out OfficeCompoundFile? compound, out _));
        byte[] opaque = { 4, 3, 2, 1 };
        byte[] withOpaque = OfficeCompoundFileWriter.Rewrite(compound!, new Dictionary<string, byte[]> {
            ["VBA/SignatureNotes"] = opaque, ["FormDesigner/o"] = opaque, ["VBA/__SRP_0"] = opaque
        });
        project = OfficeVbaProject.Load(withOpaque);
        project.SetModuleSource("Helpers", "'changed\r\n");
        byte[] written = project.Write().GetBytes();
        Assert.True(OfficeCompoundFileReader.TryRead(written, out OfficeCompoundFile? changed, out _));
        Assert.Equal(opaque, changed!.Streams["VBA/SignatureNotes"]);
        Assert.Equal(opaque, changed.Streams["FormDesigner/o"]);
        Assert.Equal(compound!.Streams["VBA/Counter"], changed.Streams["VBA/Counter"]);
        Assert.False(changed.Streams.ContainsKey("VBA/__SRP_0"));
        Assert.Equal(new byte[] { 0xcc, 0x61, 0xff, 0xff, 0, 1, 0 }, changed.Streams["VBA/_VBA_PROJECT"]);
    }

    [Fact]
    public void ProtectedProjectsAreReadableButEveryMutationIsRejected() {
        var project = OfficeVbaProject.Create();
        project.AddModule("Helpers", "'source\r\n");
        Assert.True(OfficeCompoundFileReader.TryRead(project.Write().GetBytes(), out OfficeCompoundFile? compound, out _));
        string text = OfficeVbaText.Decode(compound!.Streams["PROJECT"], 1252);
        int end = text.IndexOf('"', text.IndexOf("CMG=\"", StringComparison.Ordinal) + 5);
        byte last = byte.Parse(text.Substring(end - 2, 2), System.Globalization.NumberStyles.HexNumber);
        text = text.Substring(0, end - 2) + (last ^ 1).ToString("X2") + text.Substring(end);
        byte[] bytes = OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> { ["PROJECT"] = OfficeVbaText.Encode(text, 1252) });
        project = OfficeVbaProject.Load(bytes);
        Assert.True(project.IsProtected);
        Assert.Contains("'source", project.GetModule("Helpers").Source);
        Assert.Equal(bytes, project.Write().GetBytes());
        Assert.Throws<InvalidOperationException>(() => project.SetModuleSource("Helpers", "'changed"));
        Assert.Throws<InvalidOperationException>(() => project.AddModule("NewModule", ""));
        Assert.Throws<InvalidOperationException>(() => project.RenameModule("Helpers", "Tools"));
        Assert.Throws<InvalidOperationException>(() => project.DeleteModule("Helpers"));
        Assert.Throws<InvalidOperationException>(() => project.AddRegisteredReference("Office", Guid.NewGuid()));
        Assert.False(project.HasChanges);
    }

    [Fact]
    public void SourceImportRejectsReservedNewModulesBeforeChangingExistingSource() {
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-vba-import-" + Guid.NewGuid().ToString("N"));
        var project = OfficeVbaProject.Create();
        project.AddModule("Helpers", "'original\r\n");
        project = OfficeVbaProject.Load(project.Write().GetBytes());
        try {
            project.ExportSources(folder);
            File.WriteAllText(Path.Combine(folder, "module-0001.bas"), "'changed\r\n");
            File.WriteAllText(Path.Combine(folder, "new.bas"), "'new\r\n");
            var manifest = XDocument.Load(Path.Combine(folder, "vba-project.xml"));
            manifest.Root!.Add(new XElement("module", new XAttribute("name", "dir"),
                new XAttribute("kind", "Standard"), new XAttribute("file", "new.bas")));
            manifest.Save(Path.Combine(folder, "vba-project.xml"));
            Assert.Throws<ArgumentException>(() => project.ImportSources(folder));
            Assert.False(project.HasChanges);
        } finally { if (Directory.Exists(folder)) Directory.Delete(folder, true); }
    }

    [Theory]
    [InlineData("docm")]
    [InlineData("xlsm")]
    [InlineData("pptm")]
    public void ApplyingSourceRequiresExplicitSignatureRemovalAndPreservesOtherParts(string extension) {
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-vba-policy-" + Guid.NewGuid().ToString("N") + "." + extension);
        var project = OfficeVbaProject.Create();
        project.AddModule("Helpers", "'original\r\n");
        try {
            switch (extension) {
                case "docm": using (var doc = WordDocument.Create(path)) { doc.SetVbaProject(project); doc.Save(); } break;
                case "xlsm": using (var doc = ExcelDocument.Create(path)) { doc.AddWorksheet("Data"); doc.SetVbaProject(project); doc.Save(); } break;
                case "pptm": using (var doc = PowerPointPresentation.Create(path)) { doc.AddSlide(); doc.SetVbaProject(project); doc.Save(); } break;
            }
            using (OpenXmlPackage package = OpenPackage(path, extension)) {
                VbaProjectPart part = GetProjectPart(package);
                foreach (string suffix in new[] { "", "Agile", "V3" }) {
                    string relationship = suffix.Length == 0 ? "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature"
                        : suffix == "Agile" ? "http://schemas.microsoft.com/office/2014/relationships/vbaProjectSignatureAgile"
                        : "http://schemas.microsoft.com/office/2020/07/relationships/vbaProjectSignatureV3";
                    var signature = part.AddExtendedPart(relationship,
                        "application/vnd.ms-office.vbaProjectSignature" + suffix, ".bin");
                    using var data = new MemoryStream(new byte[] { 1, 2, 3 }); signature.FeedData(data);
                }
            }
            switch (extension) {
                case "docm": using (var doc = WordDocument.Load(path)) { var edit = doc.ReadVbaProject()!; doc.SetVbaProject(edit); edit.SetModuleSource("Helpers", "'changed\r\n"); Assert.Throws<InvalidOperationException>(() => doc.SetVbaProject(edit)); Assert.Contains("'original", doc.ReadVbaProject()!.GetModule("Helpers").Source); doc.SetVbaProject(edit, new OfficeVbaWriteOptions { AllowSignatureRemoval = true }); doc.Save(); } break;
                case "xlsm": using (var doc = ExcelDocument.Load(path)) { var edit = doc.ReadVbaProject()!; doc.SetVbaProject(edit); edit.SetModuleSource("Helpers", "'changed\r\n"); Assert.Throws<InvalidOperationException>(() => doc.SetVbaProject(edit)); Assert.Contains("'original", doc.ReadVbaProject()!.GetModule("Helpers").Source); doc.SetVbaProject(edit, new OfficeVbaWriteOptions { AllowSignatureRemoval = true }); doc.Save(); } break;
                case "pptm": using (var doc = PowerPointPresentation.Load(path)) { var edit = doc.ReadVbaProject()!; doc.SetVbaProject(edit); edit.SetModuleSource("Helpers", "'changed\r\n"); Assert.Throws<InvalidOperationException>(() => doc.SetVbaProject(edit)); Assert.Contains("'original", doc.ReadVbaProject()!.GetModule("Helpers").Source); doc.SetVbaProject(edit, new OfficeVbaWriteOptions { AllowSignatureRemoval = true }); doc.Save(); } break;
            }
            using (OpenXmlPackage package = OpenPackage(path, extension)) {
                VbaProjectPart part = GetProjectPart(package);
                Assert.DoesNotContain(part.Parts, pair => pair.OpenXmlPart.ContentType.Contains("vbaProjectSignature", StringComparison.OrdinalIgnoreCase));
                if (extension == "docm") Assert.NotNull(part.VbaDataPart);
                Assert.Contains("'changed", OfficeVbaProject.Load(Read(part)).GetModule("Helpers").Source);
            }
        } finally { if (File.Exists(path)) File.Delete(path); }
    }

    private static OpenXmlPackage OpenPackage(string path, string extension) => extension switch {
        "docm" => WordprocessingDocument.Open(path, true), "xlsm" => SpreadsheetDocument.Open(path, true),
        _ => PresentationDocument.Open(path, true)
    };

    private static VbaProjectPart GetProjectPart(OpenXmlPackage package) => package switch {
        WordprocessingDocument doc => doc.MainDocumentPart!.VbaProjectPart!,
        SpreadsheetDocument doc => doc.WorkbookPart!.VbaProjectPart!,
        PresentationDocument doc => doc.PresentationPart!.VbaProjectPart!,
        _ => throw new InvalidOperationException()
    };

    private static byte[] Read(VbaProjectPart part) {
        using var stream = part.GetStream(); using var output = new MemoryStream(); stream.CopyTo(output); return output.ToArray();
    }
}
