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

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OwnedVbaMutationRemovesTheOriginalSignatureAndKeepsTheMacroSubtype(bool remove) {
        const string projectType = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
        byte[] bytes = ReadInertProject();
        byte[] fixture = CreateFixture(bytes, includeSignature: true);
        using var original = OfficeOpcPackage.Open(new MemoryStream(fixture, writable: false));
        string main = original.GetRelatedPartName(null, OfficeOpcPackage.OfficeDocumentRelationshipType)!;
        OfficeOpcRelationship project = Assert.Single(original.GetRelationships(main), edge => edge.Type == projectType);
        string signature = Assert.Single(original.GetRelationships(project.TargetPartName)).TargetPartName!;

        using var presentation = PowerPointDocument.Load(new MemoryStream(fixture, writable: false));
        if (remove) {
            Assert.True(presentation.RemoveVbaProject());
            Assert.False(presentation.RemoveVbaProject());
            Assert.False(presentation.HasVbaProject);
        } else {
            byte[] callerOwned = (byte[])bytes.Clone();
            presentation.SetVbaProject(callerOwned);
            callerOwned[callerOwned.Length - 1] ^= 0xFF;
            Assert.Equal(bytes, presentation.GetVbaProjectBytes());
            Assert.False(presentation.InspectVbaProjectMetadata()!.HasSignature);
        }
        byte[] saved = presentation.ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Pptm });
        Assert.False(presentation.IsSdkLoaded);

        using var actual = OfficeOpcPackage.Open(new MemoryStream(saved, writable: false));
        Assert.False(actual.PartExists(signature));
        Assert.Equal(!remove, actual.PartExists(project.TargetPartName!));
        if (remove) {
            Assert.DoesNotContain(actual.GetRelationships(main), edge => edge.Type == projectType);
        } else {
            OfficeOpcRelationship savedProject = Assert.Single(actual.GetRelationships(main), edge => edge.Type == projectType);
            Assert.Equal(project.Id, savedProject.Id);
            Assert.Equal(project.TargetPartName, savedProject.TargetPartName);
        }
        using var reopened = PowerPointDocument.Load(new MemoryStream(saved, writable: false));
        Assert.Equal(".pptm", reopened.SourceFormatDescriptor.Extension);
        Assert.Equal(!remove, reopened.HasVbaProject);
    }

    [Theory]
    [InlineData(false, null)]
    [InlineData(true, null)]
    [InlineData(false, "application/vnd.ms-office.vbaProject")]
    [InlineData(false, "application/vnd.ms-office.VbaProject")]
    public void AddsAndReaddsAVbaProjectWithTheSdkPartNames(bool openSdkBetweenAdds, string? orphanContentType) {
        using var original = PowerPointDocument.Create();
        byte[] fixture = original.ToBytes();
        if (orphanContentType != null) {
            using var package = OfficeOpcEditablePackage.Open(new MemoryStream(fixture, writable: false));
            package.CreateBinaryPart("/ppt/old.bin", orphanContentType, new byte[] { 7 });
            using var edited = new MemoryStream();
            package.Save(edited);
            fixture = edited.ToArray();
        }
        using var baseline = PowerPointDocument.Load(new MemoryStream(fixture, writable: false));
        _ = baseline.OpenXmlDocument;
        using var presentation = PowerPointDocument.Load(new MemoryStream(fixture, writable: false));
        byte[] bytes = ReadInertProject();
        Assert.False(presentation.RemoveVbaProject());

        for (int attempt = 0; attempt < 2; attempt++) {
            baseline.SetVbaProject(bytes);
            presentation.SetVbaProject(bytes);

            Assert.Equal(bytes, presentation.GetVbaProjectBytes());
            Assert.Equal(baseline.InspectVbaProjectMetadata()!.PartUri, presentation.InspectVbaProjectMetadata()!.PartUri);
            Assert.True(baseline.RemoveVbaProject());
            Assert.True(presentation.RemoveVbaProject());
            if (openSdkBetweenAdds && attempt == 0) {
                Assert.False(presentation.IsSdkLoaded);
                _ = presentation.OpenXmlDocument;
            }
        }
        Assert.Equal(openSdkBetweenAdds || orphanContentType != null, presentation.IsSdkLoaded);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MacroFreeSaveOmitsAProjectAddedAfterLoad(bool explicitFormat) {
        using var original = PowerPointDocument.Create();
        using var presentation = PowerPointDocument.Load(new MemoryStream(original.ToBytes(), writable: false));
        byte[] bytes = ReadInertProject();
        presentation.SetVbaProject(bytes);
        Assert.False(presentation.IsSdkLoaded);

        byte[] saved = explicitFormat
            ? presentation.ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Pptx })
            : presentation.ToBytes();

        using var reopened = PowerPointDocument.Load(new MemoryStream(saved, writable: false));
        Assert.Equal(".pptx", reopened.SourceFormatDescriptor.Extension);
        Assert.False(reopened.HasVbaProject);
        Assert.Equal(bytes, presentation.GetVbaProjectBytes());
        Assert.False(presentation.IsSdkLoaded);
    }

    [Theory]
    [InlineData(".potm", true)]
    [InlineData(".pptx", false)]
    public void ConvertingSubtypePreservesVbaOnlyInTheMacroEnabledOutput(string extension, bool preservesVba) {
        byte[] bytes = ReadInertProject();
        byte[] fixture = CreateFixture(bytes, includeSignature: true);
        using var presentation = PowerPointDocument.Load(new MemoryStream(fixture, writable: false));
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-SubtypeVba-" + Guid.NewGuid().ToString("N") + extension);
        try {
            presentation.SaveCopy(path);
            using var saved = PowerPointDocument.Load(path);

            Assert.Equal(extension, saved.SourceFormatDescriptor.Extension);
            Assert.Equal(preservesVba, saved.HasVbaProject);
            if (preservesVba) {
                Assert.Equal(bytes, saved.GetVbaProjectBytes());
                Assert.True(saved.InspectVbaProjectMetadata()!.HasSignature);
            }
            Assert.Equal(".pptm", presentation.SourceFormatDescriptor.Extension);
            Assert.Equal(bytes, presentation.GetVbaProjectBytes());
            Assert.True(presentation.InspectVbaProjectMetadata()!.HasSignature);
            Assert.False(presentation.IsSdkLoaded);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void VbaMutationDistinguishesEscapedCacheAndLivePartNames(bool remove) {
        const string projectType = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
        const string cache = "/ppt/cache%2520copy.bin";
        const string kept = "/ppt/cache%20copy.bin";
        byte[] bytes = ReadInertProject();
        byte[] fixture = CreateFixture(bytes, includeSignature: false);
        using var package = OfficeOpcEditablePackage.Open(new MemoryStream(fixture, writable: false));
        using var packageInfo = OfficeOpcPackage.Open(new MemoryStream(fixture, writable: false));
        string main = packageInfo.GetRelatedPartName(null, OfficeOpcPackage.OfficeDocumentRelationshipType)!;
        string project = packageInfo.GetRelatedPartName(main, projectType)!;
        string slide = packageInfo.GetRelatedPartName(main, OfficeOpcPackage.SlideRelationshipType)!;
        package.CreateBinaryPart(cache, "application/octet-stream", new byte[] { 1 });
        package.CreateBinaryPart(kept, "application/octet-stream", new byte[] { 2 });
        package.AddRelationship(project, "urn:officeimo:vba-cache", cache);
        package.AddRelationship(slide, OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "package", kept);
        using var modified = new MemoryStream();
        package.Save(modified);
        modified.Position = 0;

        using var presentation = PowerPointDocument.Load(modified);
        if (remove) Assert.True(presentation.RemoveVbaProject());
        else presentation.SetVbaProject(bytes);
        byte[] saved = presentation.ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Pptm });
        Assert.False(presentation.IsSdkLoaded);

        using var actual = OfficeOpcPackage.Open(new MemoryStream(saved, writable: false));
        Assert.False(actual.PartExists(cache));
        Assert.True(actual.PartExists(kept));
        Assert.Contains(actual.GetRelationships(slide), relationship => relationship.TargetPartName == kept);
    }

    [Fact]
    public void VbaAddDeclaresTheContentTypeOfANewRelationshipPart() {
        const string main = "/ppt/presentation.xml";
        byte[] project = ReadInertProject();
        using var package = OfficeOpcEditablePackage.Create();
        OfficeXmlTree root = package.CreateXmlPart(main, "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml",
            OfficeXmlNamespaces.PresentationML, "presentation");
        root.DeclareNamespace(root.Root, "p", "http://schemas.openxmlformats.org/presentationml/2006/main");
        package.AddRelationship(null, OfficeOpcPackage.OfficeDocumentRelationshipType, main);
        package.SetOverride("/_rels/.rels", OfficeOpcEditablePackage.RelationshipContentType);
        OfficeXmlTree types = package.ContentTypes;
        int defaultRelationships = types.FindChild(types.Root, OfficeXmlNamespaces.ContentTypes, "Default");
        Assert.Equal("rels", types.GetAttribute(defaultRelationships, OfficeXmlNamespaces.None, "Extension"));
        types.SetAttribute(defaultRelationships, OfficeXmlNamespaces.None, "ContentType", "application/xml");
        using var fixture = new MemoryStream();
        package.Save(fixture);
        fixture.Position = 0;

        using var presentation = PowerPointDocument.Load(fixture);
        presentation.SetVbaProject(project);
        byte[] saved = presentation.ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Pptm });
        Assert.False(presentation.IsSdkLoaded);

        using var sdk = PresentationDocument.Open(new MemoryStream(saved, writable: false), false);
        Assert.NotNull(sdk.PresentationPart!.VbaProjectPart);
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

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EditsVbaSourceWithTheExistingSignatureAndRecoveryPolicy(bool hasCache) {
        const string cachePart = "/ppt/vba/cache.bin";
        const string olePart = "/ppt/embeddings/object1.bin";
        const string cacheRelationship = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "package";
        const string projectRelationship = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
        OfficeVbaProject project = OfficeVbaProject.Create("SourceFixture");
        project.AddModule("Helpers", "' original source");
        byte[] original = project.Write().GetBytes();
        byte[] fixture = CreateFixture(original, includeSignature: true);
        if (hasCache) {
            using var reader = OfficeOpcPackage.Open(new MemoryStream(fixture, writable: false));
            string main = reader.GetRelatedPartName(null, OfficeOpcPackage.OfficeDocumentRelationshipType)!;
            string carrier = reader.GetRelatedPartName(main, projectRelationship)!;
            string slide = reader.GetRelatedPartName(main, OfficeOpcPackage.SlideRelationshipType)!;
            using var editable = OfficeOpcEditablePackage.Open(new MemoryStream(fixture, writable: false));
            editable.CreateBinaryPart(cachePart, "application/octet-stream", new byte[] { 5, 9 });
            editable.AddRelationship(carrier, cacheRelationship, cachePart, id: "sourceCache");
            editable.CreateBinaryPart(olePart, "application/vnd.openxmlformats-officedocument.oleObject", new byte[] { 7 });
            editable.AddRelationship(slide, OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "oleObject", olePart, id: "opaqueObject");
            using var modified = new MemoryStream();
            editable.Save(modified);
            fixture = modified.ToArray();
        }

        using var presentation = PowerPointDocument.Load(new MemoryStream(fixture, writable: false));
        Assert.Throws<InvalidDataException>(() => presentation.ReadVbaProject(new OfficeVbaReadOptions { MaximumProjectBytes = 1 }));
        OfficeVbaProject source = presentation.ReadVbaProject()!;
        Assert.Contains("original source", Assert.Single(source.Modules).Source);
        presentation.SetVbaProject(source, new OfficeVbaWriteOptions { MaximumRecoveryBytes = 1 });
        Assert.True(presentation.InspectVbaProjectMetadata()!.HasSignature);

        source.SetModuleSource("Helpers", "' revised source");
        Assert.Throws<InvalidOperationException>(() => presentation.SetVbaProject(source));
        Assert.Throws<InvalidDataException>(() => presentation.SetVbaProject(source,
            new OfficeVbaWriteOptions { AllowSignatureRemoval = true, MaximumRecoveryBytes = 1 }));
        Assert.True(presentation.InspectVbaProjectMetadata()!.HasSignature);
        Assert.Equal(original, presentation.GetVbaProjectBytes());
        Assert.False(presentation.IsSdkLoaded);

        using var baseline = PowerPointDocument.Load(new MemoryStream(fixture, writable: false));
        _ = baseline.OpenXmlDocument;
        var options = new OfficeVbaWriteOptions { AllowSignatureRemoval = true };
        baseline.SetVbaProject(source, options);
        presentation.SetVbaProject(source, options);
        Assert.Equal(baseline.GetVbaProjectBytes(), presentation.GetVbaProjectBytes());
        byte[] saved = presentation.ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Pptm });
        Assert.False(presentation.IsSdkLoaded);

        using var reopened = PowerPointDocument.Load(new MemoryStream(saved, writable: false));
        Assert.False(reopened.InspectVbaProjectMetadata()!.HasSignature);
        Assert.Contains("revised source", Assert.Single(reopened.ReadVbaProject()!.Modules).Source);
        using var package = OfficeOpcPackage.Open(new MemoryStream(saved, writable: false));
        Assert.Equal(hasCache, package.PartExists(cachePart));
        Assert.Equal(hasCache, package.PartExists(olePart));
        if (hasCache) {
            using var storedCache = package.ReadPart(cachePart);
            Assert.Equal(new byte[] { 5, 9 }, storedCache.Array.Take(storedCache.Length));
        }
    }

    [Fact]
    public void AddsADetachedSourceProjectToAnOwnedPresentation() {
        using var presentation = PowerPointDocument.Create();
        Assert.Null(presentation.ReadVbaProject());
        OfficeVbaProject project = OfficeVbaProject.Create();
        project.AddModule("Helpers", "' source only");

        presentation.SetVbaProject(project);
        byte[] saved = presentation.ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Pptm });

        Assert.False(presentation.IsSdkLoaded);
        using var reopened = PowerPointDocument.Load(new MemoryStream(saved, writable: false));
        Assert.Contains("source only", Assert.Single(reopened.ReadVbaProject()!.Modules).Source);
        Assert.False(reopened.IsSdkLoaded);
    }

    [Theory]
    [InlineData("fixed part type")]
    [InlineData("relationship part type")]
    [InlineData("relationship markup")]
    [InlineData("duplicate relationship part type")]
    public void DetachedVbaSourceOperationsKeepTheSdkErrorForAMalformedPackage(string defect) {
        OfficeVbaProject source = OfficeVbaProject.Create();
        source.AddModule("Helpers", "' original source");
        byte[] original = source.Write().GetBytes();
        byte[] seed = CreateFixture(original, includeSignature: false);
        using var info = OfficeOpcPackage.Open(new MemoryStream(seed, writable: false));
        string main = info.GetRelatedPartName(null, OfficeOpcPackage.OfficeDocumentRelationshipType)!;
        string view = info.GetRelatedPartName(main, OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "viewProps")!;
        using var edit = OfficeOpcEditablePackage.Open(new MemoryStream(seed, writable: false));
        if (defect == "relationship markup") {
            OfficeXmlTree relationships = edit.GetRelationships(main);
            relationships.Append(relationships.Root, relationships.CreateElement(OfficeXmlNamespaces.PackageRelationships, "Unexpected"));
        } else if (defect == "duplicate relationship part type") {
            string relationships = OfficeOpcPartNames.GetRelationshipsPartName(main);
            edit.SetOverride(relationships, "application/xml");
            OfficeXmlTree types = edit.ContentTypes;
            int duplicate = types.CreateElement(OfficeXmlNamespaces.ContentTypes, "Override");
            types.SetAttribute(duplicate, OfficeXmlNamespaces.None, "PartName", relationships);
            types.SetAttribute(duplicate, OfficeXmlNamespaces.None, "ContentType", OfficeOpcEditablePackage.RelationshipContentType);
            types.Append(types.Root, duplicate);
        } else {
            edit.SetOverride(defect == "relationship part type" ? OfficeOpcPartNames.GetRelationshipsPartName(main) : view, "application/xml");
        }
        using var output = new MemoryStream();
        edit.Save(output);
        byte[] fixture = output.ToArray();

        using var baseline = PowerPointDocument.Load(new MemoryStream(fixture, writable: false));
        Exception expected = Assert.ThrowsAny<Exception>(() => {
            _ = baseline.OpenXmlDocument;
            _ = baseline.ReadVbaProject();
        });
        using var reader = PowerPointDocument.Load(new MemoryStream(fixture, writable: false));
        Assert.IsType(expected.GetType(), Record.Exception(() => reader.ReadVbaProject()));

        using var writer = PowerPointDocument.Load(new MemoryStream(fixture, writable: false));
        source.SetModuleSource("Helpers", "' changed source");
        Assert.IsType(expected.GetType(), Record.Exception(() => writer.SetVbaProject(source)));
        Assert.Equal(original, writer.GetVbaProjectBytes());
    }

    [Fact]
    public void SourceVbaAttachmentKeepsItsPartSequenceAcrossSdkHandoff() {
        OfficeVbaProject project = OfficeVbaProject.Create();
        project.AddModule("Helpers", "' source only");
        using var baseline = PowerPointDocument.Create();
        _ = baseline.OpenXmlDocument;
        baseline.SetVbaProject(project);
        Assert.True(baseline.RemoveVbaProject());
        baseline.SetVbaProject(project);

        using var presentation = PowerPointDocument.Create();
        presentation.SetVbaProject(project);
        Assert.True(presentation.RemoveVbaProject());
        Assert.False(presentation.IsSdkLoaded);
        _ = presentation.OpenXmlDocument;
        presentation.SetVbaProject(project);

        Assert.Equal(baseline.InspectVbaProjectMetadata()!.PartUri, presentation.InspectVbaProjectMetadata()!.PartUri);
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
