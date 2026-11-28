using System.IO.Compression;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointVbaProjectReadTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OwnedVbaReadsMatchTheSdkForAnInertProject(bool includeSignature) {
        byte[] project = ReadInertProject();
        byte[] fixture = CreateFixture(project, includeSignature);
        using var baseline = LoadReadOnly(fixture);
        _ = baseline.OpenXmlDocument;
        OfficeVbaProjectInfo expectedMetadata = baseline.InspectVbaProjectMetadata()!;
        OfficeVbaProjectInfo expected = baseline.InspectVbaProject(includeSha256: true)!;
        using var presentation = LoadReadOnly(fixture);

        Assert.True(presentation.HasVbaProject);
        OfficeVbaProjectInfo descriptor = Assert.IsType<OfficeVbaProjectInfo>(presentation.InspectVbaProjectMetadata());
        Assert.Equal(expected.PartUri, descriptor.PartUri);
        Assert.Equal(expected.ContentType, descriptor.ContentType);
        Assert.Equal(includeSignature, expectedMetadata.HasSignature);
        Assert.Equal(expectedMetadata.HasSignature, descriptor.HasSignature);
        Assert.False(descriptor.HasKnownLength);
        Assert.False(descriptor.ModuleInspectionRequested);
        Assert.Empty(descriptor.ModuleNames);
        Assert.Null(descriptor.Sha256);
        Assert.Equal(project, presentation.GetVbaProjectBytes(project.Length));
        Assert.Throws<InvalidDataException>(() => presentation.GetVbaProjectBytes(project.Length - 1));
        Assert.Throws<InvalidDataException>(() => presentation.InspectVbaProject(maxBytes: project.Length - 1));

        OfficeVbaProjectInfo actual = Assert.IsType<OfficeVbaProjectInfo>(presentation.InspectVbaProject(includeSha256: true));
        Assert.Equal(expected.Length, actual.Length);
        Assert.Equal(expected.Sha256, actual.Sha256);
        Assert.Equal(expected.ModuleNames, actual.ModuleNames);
        Assert.NotEmpty(actual.ModuleNames);
        Assert.Equal(expected.HasSignature, actual.HasSignature);
        Assert.True(actual.ModuleInspectionRequested);
        Assert.False(presentation.IsSdkLoaded);
    }

    [Fact]
    public void SameSubtypeSaveRetainsTheVbaProjectAndSignatureBytes() {
        byte[] project = ReadInertProject();
        byte[] fixture = CreateFixture(project, includeSignature: true);
        using var presentation = PowerPointDocument.Load(new MemoryStream(fixture, writable: false));
        Assert.Single(presentation.Slides).AddTextBox("Saved with VBA");
        byte[] saved = presentation.ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Pptm });
        Assert.False(presentation.IsSdkLoaded);

        using var reopened = PowerPointDocument.Load(new MemoryStream(saved, writable: false));
        Assert.Equal(".pptm", reopened.SourceFormatDescriptor.Extension);
        Assert.Equal(project, reopened.GetVbaProjectBytes());
        Assert.True(reopened.InspectVbaProjectMetadata()!.HasSignature);
        Assert.Contains(Assert.Single(reopened.Slides).TextBoxes, box => box.Text == "Saved with VBA");
        Assert.False(reopened.IsSdkLoaded);
        using var originalPackage = PresentationDocument.Open(new MemoryStream(fixture, writable: false), false);
        using var savedPackage = PresentationDocument.Open(new MemoryStream(saved, writable: false), false);
        OpenXmlPart originalSignature = Assert.Single(originalPackage.PresentationPart!.VbaProjectPart!.Parts).OpenXmlPart;
        OpenXmlPart savedSignature = Assert.Single(savedPackage.PresentationPart!.VbaProjectPart!.Parts).OpenXmlPart;
        using var expected = new MemoryStream();
        using var actual = new MemoryStream();
        using (Stream entry = originalSignature.GetStream()) entry.CopyTo(expected);
        using (Stream entry = savedSignature.GetStream()) entry.CopyTo(actual);
        Assert.Equal(expected.ToArray(), actual.ToArray());
    }

    [Fact]
    public void MacroFreePresentationsReportNoProjectWithoutOpeningTheSdk() {
        using var presentation = PowerPointDocument.Create();
        presentation.AddSlide();

        Assert.False(presentation.HasVbaProject);
        Assert.Null(presentation.GetVbaProjectBytes());
        Assert.Null(presentation.InspectVbaProjectMetadata());
        Assert.Null(presentation.InspectVbaProject());
        Assert.False(presentation.IsSdkLoaded);
    }

    [Theory]
    [InlineData("project")]
    [InlineData("signature")]
    [InlineData("slide")]
    public void MalformedRelationshipsKeepTheExistingVbaReadFailure(string damagedOwner) {
        const string projectType = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
        byte[] fixture = CreateFixture(ReadInertProject(), includeSignature: true, includeSlide: damagedOwner != "project");
        using var original = OfficeOpcPackage.Open(new MemoryStream(fixture, writable: false));
        string main = original.GetRelatedPartName(null, OfficeOpcPackage.OfficeDocumentRelationshipType)!;
        string project = original.GetRelatedPartName(main, projectType)!;
        using var package = OfficeOpcEditablePackage.Open(new MemoryStream(fixture, writable: false));
        const int ns = OfficeXmlNamespaces.PackageRelationships;
        switch (damagedOwner) {
            case "project":
                OfficeXmlTree mainRelationships = package.GetRelationships(main);
                string projectId = original.GetRelationships(main).Single(relationship => relationship.Type == projectType).Id;
                for (int node = mainRelationships.FindChild(mainRelationships.Root, ns, "Relationship");
                     node != OfficeXmlTree.NoNode;
                     node = mainRelationships.FindNextSibling(node, ns, "Relationship")) {
                    if (mainRelationships.GetAttribute(node, OfficeXmlNamespaces.None, "Id") == projectId) {
                        mainRelationships.RemoveAttribute(node, OfficeXmlNamespaces.None, "Id");
                        break;
                    }
                }
                break;
            case "signature":
                OfficeXmlTree signatures = package.GetRelationships(project);
                int existing = signatures.FindChild(signatures.Root, ns, "Relationship");
                int duplicate = signatures.CloneElement(existing);
                signatures.Append(signatures.Root, duplicate);
                break;
            default:
                string slide = original.GetRelatedPartName(main, OfficeOpcPackage.SlideRelationshipType)!;
                package.AddRelationship(slide, "urn:officeimo:missing-child", "/ppt/missing-child.xml");
                break;
        }
        using var output = new MemoryStream();
        package.Save(output);
        fixture = output.ToArray();
        using var baseline = LoadReadOnly(fixture);
        using var presentation = LoadReadOnly(fixture);
        Exception? expected = Record.Exception(() => {
            _ = baseline.OpenXmlDocument;
            _ = baseline.InspectVbaProjectMetadata();
        });

        Assert.NotNull(expected);
        Assert.IsType(expected.GetType(), Record.Exception(() => presentation.InspectVbaProjectMetadata()));
    }

    [Fact]
    public void VbaReadsHonorTheConfiguredSdkPartLimit() {
        byte[] fixture = CreateFixture(ReadInertProject(), includeSignature: false);
        using var presentation = PowerPointDocument.Load(new MemoryStream(fixture, writable: false),
            new PowerPointLoadOptions {
                AccessMode = DocumentAccessMode.ReadOnly,
                OpenSettings = new OfficeOpenXmlLoadSettings { MaxCharactersInPart = 1_000_000 }
            });

        Assert.NotNull(presentation.InspectVbaProjectMetadata());
        Assert.True(presentation.IsSdkLoaded);
    }

    private static PowerPointDocument LoadReadOnly(byte[] fixture) =>
        PowerPointDocument.Load(new MemoryStream(fixture, writable: false),
            new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });

    private static byte[] ReadInertProject() {
        string source = Path.Combine(AppContext.BaseDirectory, "Documents", "ExcelPreservationCorpus", "inert-macro-control.xlsm");
        using var archive = new ZipArchive(File.OpenRead(source), ZipArchiveMode.Read);
        using Stream entry = archive.GetEntry("xl/vbaProject.bin")!.Open();
        using var bytes = new MemoryStream();
        entry.CopyTo(bytes);
        return bytes.ToArray();
    }

    private static byte[] CreateFixture(byte[] project, bool includeSignature, bool includeSlide = true) {
        using var presentation = PowerPointDocument.Create();
        if (includeSlide) presentation.AddSlide(PowerPointSlideLayoutType.Blank).AddTextBox("Inert VBA fixture");
        byte[] data = presentation.ToBytes();
        using var stream = new MemoryStream();
        stream.Write(data, 0, data.Length);
        stream.Position = 0;
        using (var document = PresentationDocument.Open(stream, true)) {
            document.ChangeDocumentType(PresentationDocumentType.MacroEnabledPresentation);
            VbaProjectPart part = document.PresentationPart!.AddNewPart<VbaProjectPart>();
            using var source = new MemoryStream(project, writable: false);
            part.FeedData(source);
            if (includeSignature) {
                var signature = part.AddExtendedPart("http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature",
                    "application/vnd.ms-office.vbaProjectSignature", ".bin");
                using var payload = new MemoryStream(new byte[] { 7 }, writable: false);
                signature.FeedData(payload);
            }
        }
        return stream.ToArray();
    }
}
