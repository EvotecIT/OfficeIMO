using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointModernSubtypeLoadTests {
    [Theory]
    [InlineData(PresentationDocumentType.MacroEnabledPresentation, ".pptm")]
    [InlineData(PresentationDocumentType.Template, ".potx")]
    [InlineData(PresentationDocumentType.MacroEnabledTemplate, ".potm")]
    [InlineData(PresentationDocumentType.Slideshow, ".ppsx")]
    [InlineData(PresentationDocumentType.MacroEnabledSlideshow, ".ppsm")]
    [InlineData(PresentationDocumentType.AddIn, ".ppam")]
    public void ReadsSupportedModernSubtypesBeforeOpeningTheSdk(PresentationDocumentType type, string extension) {
        byte[] fixture = CreateFixture(type, signed: false);
        using var readOnly = PowerPointDocument.Load(new MemoryStream(fixture, writable: false),
            new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
        using var presentation = PowerPointDocument.Load(new MemoryStream(fixture, writable: false));

        Assert.Equal(extension, readOnly.SourceFormatDescriptor.Extension);
        Assert.Equal(extension, presentation.SourceFormatDescriptor.Extension);
        PowerPointSlide slide = Assert.Single(presentation.Slides);
        Assert.Equal("Subtype fixture", Assert.Single(slide.TextBoxes).Text);
        Assert.Equal("Subtype fixture", Assert.Single(Assert.Single(readOnly.Slides).TextBoxes).Text);
        Assert.False(presentation.IsSdkLoaded);
        Assert.False(readOnly.IsSdkLoaded);
        slide.AddTextBox("Pending edit");
        Assert.False(presentation.IsSdkLoaded);

        byte[] copy = presentation.ToBytes();
        Assert.False(presentation.IsSdkLoaded);
        Assert.Same(slide, Assert.Single(presentation.Slides));
        using var output = PresentationDocument.Open(new MemoryStream(copy, writable: false), false);
        using var reopened = PowerPointDocument.Load(new MemoryStream(copy, writable: false));

        Assert.Equal(PresentationDocumentType.Presentation, output.DocumentType);
        Assert.Equal(extension, presentation.SourceFormatDescriptor.Extension);
        Assert.Contains(Assert.Single(reopened.Slides).TextBoxes, box => box.Text == "Pending edit");
    }

    [Theory]
    [InlineData(PresentationDocumentType.MacroEnabledPresentation, ".pptm", null)]
    [InlineData(PresentationDocumentType.Template, ".potx", "occupied name")]
    [InlineData(PresentationDocumentType.Template, ".potx", "stale override")]
    [InlineData(PresentationDocumentType.Template, ".potx", "external relationship")]
    [InlineData(PresentationDocumentType.MacroEnabledTemplate, ".potm", null)]
    [InlineData(PresentationDocumentType.Slideshow, ".ppsx", null)]
    [InlineData(PresentationDocumentType.MacroEnabledSlideshow, ".ppsm", null)]
    [InlineData(PresentationDocumentType.AddIn, ".ppam", null)]
    public void SavesAnotherSubtypeWithTheSamePhysicalMainPartAsTheSdk(PresentationDocumentType type, string extension, string? existing) {
        byte[] fixture = CreateFixture(PresentationDocumentType.Presentation, signed: false);
        using var source = OfficeOpcPackage.Open(new MemoryStream(fixture, writable: false));
        OfficeOpcRelationship originalMain = Assert.Single(source.GetRelationships(null),
            edge => edge.Type == OfficeOpcPackage.OfficeDocumentRelationshipType);
        if (existing != null) {
            fixture = EditPackage(fixture, package => {
                const string collision = "/ppt/presentation2.xml";
                switch (existing) {
                    case "occupied name":
                        package.CreateBinaryPart(collision, "application/octet-stream", new byte[] { 1 });
                        package.AddRelationship(originalMain.TargetPartName, OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "package", collision);
                        break;
                    case "stale override":
                        package.SetOverride(collision, "application/octet-stream");
                        break;
                    case "external relationship":
                        package.AddRelationship(null, "urn:officeimo:package-external", "https://example.invalid/document", external: true);
                        break;
                }
            });
        }
        string filename = Path.Combine(Path.GetTempPath(), "OfficeIMO-SubtypeConversion-" + Guid.NewGuid().ToString("N"));
        string expectedPath = filename + "-expected" + extension;
        string actualPath = filename + extension;
        try {
            using var baseline = PowerPointDocument.Load(new MemoryStream(fixture, writable: false));
            _ = baseline.OpenXmlDocument;
            Assert.Single(baseline.Slides).AddTextBox("Saved edit");
            baseline.SaveCopy(expectedPath);
            using var presentation = PowerPointDocument.Load(new MemoryStream(fixture, writable: false));
            Assert.False(presentation.IsSdkLoaded);
            PowerPointSlide slide = Assert.Single(presentation.Slides);
            slide.AddTextBox("Saved edit");
            Assert.False(presentation.IsSdkLoaded);
            presentation.SaveCopy(actualPath);
            Assert.False(presentation.IsSdkLoaded);
            Assert.Same(slide, Assert.Single(presentation.Slides));
            Assert.Equal(".pptx", presentation.SourceFormatDescriptor.Extension);

            using var expected = OfficeOpcPackage.Open(File.OpenRead(expectedPath));
            using var actual = OfficeOpcPackage.Open(File.OpenRead(actualPath));
            OfficeOpcRelationship expectedMain = Assert.Single(expected.GetRelationships(null),
                edge => edge.Type == OfficeOpcPackage.OfficeDocumentRelationshipType);
            OfficeOpcRelationship actualMain = Assert.Single(actual.GetRelationships(null),
                edge => edge.Type == OfficeOpcPackage.OfficeDocumentRelationshipType);
            Assert.Equal(originalMain.Id, actualMain.Id);
            Assert.Equal(expectedMain.TargetPartName, actualMain.TargetPartName);
            Assert.NotEqual(originalMain.TargetPartName, actualMain.TargetPartName);
            Assert.False(actual.PartExists(originalMain.TargetPartName!));
            Assert.DoesNotContain(expected.GetRelationships(null), edge => edge.IsExternal);
            Assert.DoesNotContain(actual.GetRelationships(null), edge => edge.IsExternal);
            using var sdk = PresentationDocument.Open(actualPath, false);
            Assert.Equal(type, sdk.DocumentType);
            using var reopened = PowerPointDocument.Load(actualPath);
            Assert.Contains(Assert.Single(reopened.Slides).TextBoxes, box => box.Text == "Saved edit");
            if (existing == "occupied name") {
                using var again = OfficeOpcPackage.Open(new MemoryStream(presentation.ToBytes(), writable: false));
                Assert.Equal(originalMain.TargetPartName, again.GetRelatedPartName(null, OfficeOpcPackage.OfficeDocumentRelationshipType));
                Assert.False(presentation.IsSdkLoaded);
            }
        } finally {
            if (File.Exists(expectedPath)) File.Delete(expectedPath);
            if (File.Exists(actualPath)) File.Delete(actualPath);
        }
    }

    [Fact]
    public void SubtypeConversionKeepsTheRelationshipsForNewDocumentProperties() {
        const string appType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "extended-properties";
        const int ns = OfficeXmlNamespaces.PackageRelationships;
        byte[] fixture = EditPackage(CreateFixture(PresentationDocumentType.Template, signed: false), package => {
            OfficeXmlTree relationships = package.GetRelationships(null);
            int relationship = relationships.FindChild(relationships.Root, ns, "Relationship");
            while (relationship != OfficeXmlTree.NoNode) {
                int next = relationships.FindNextSibling(relationship, ns, "Relationship");
                if (relationships.GetAttribute(relationship, OfficeXmlNamespaces.None, "Type") is appType or OfficeOpcPackage.CorePropertiesRelationshipType) {
                    string id = relationships.GetAttribute(relationship, OfficeXmlNamespaces.None, "Id")!;
                    package.DeletePart(package.GetRelationshipTarget(null, id)!);
                    relationships.Remove(relationship);
                }
                relationship = next;
            }
        });
        using var presentation = PowerPointDocument.Load(new MemoryStream(fixture, writable: false));
        byte[] converted = presentation.ToBytes();
        Assert.False(presentation.IsSdkLoaded);

        using var package = OfficeOpcPackage.Open(new MemoryStream(converted, writable: false));
        Assert.True(package.PartExists(package.GetRelatedPartName(null, appType)!));
        Assert.True(package.PartExists(package.GetRelatedPartName(null, OfficeOpcPackage.CorePropertiesRelationshipType)!));
    }

    [Fact]
    public void SubtypeConversionRetainsSdkFailureForAMediaPartWithTwoRoles() {
        byte[] fixture = CreateFixture(PresentationDocumentType.Presentation, signed: false);
        using var original = OfficeOpcPackage.Open(new MemoryStream(fixture, writable: false));
        string main = original.GetRelatedPartName(null, OfficeOpcPackage.OfficeDocumentRelationshipType)!;
        string slide = original.GetRelatedPartName(main, OfficeOpcPackage.SlideRelationshipType)!;
        fixture = EditPackage(fixture, package => {
            const string video = "/ppt/media/mixed.mp4";
            package.CreateBinaryPart(video, "video/mp4", new byte[] { 1 });
            package.AddRelationship(slide, OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "package", video);
            package.AddRelationship(slide, OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "video", video);
        });

        AssertTemplateSaveFailure(fixture);
    }

    [Theory]
    [InlineData(PresentationDocumentType.MacroEnabledPresentation, ".pptm")]
    [InlineData(PresentationDocumentType.Template, ".potx")]
    [InlineData(PresentationDocumentType.MacroEnabledTemplate, ".potm")]
    [InlineData(PresentationDocumentType.Slideshow, ".ppsx")]
    [InlineData(PresentationDocumentType.MacroEnabledSlideshow, ".ppsm")]
    [InlineData(PresentationDocumentType.AddIn, ".ppam")]
    public void SavesAnEditedCopyInTheLoadedSubtypeWithoutOpeningTheSdk(PresentationDocumentType type, string extension) {
        byte[] fixture = CreateFixture(type, signed: false);
        using var original = OfficeOpcPackage.Open(new MemoryStream(fixture, writable: false));
        OfficeOpcRelationship originalMain = Assert.Single(original.GetRelationships(null),
            edge => edge.Type == OfficeOpcPackage.OfficeDocumentRelationshipType);
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-SubtypeSave-" + Guid.NewGuid().ToString("N") + extension);
        try {
            using var presentation = PowerPointDocument.Load(new MemoryStream(fixture, writable: false));
            Assert.Single(presentation.Slides).AddTextBox("Saved edit");
            presentation.SaveCopy(path);
            Assert.False(presentation.IsSdkLoaded);
            Assert.Equal(extension, presentation.SourceFormatDescriptor.Extension);

            byte[] saved = File.ReadAllBytes(path);
            using var actual = OfficeOpcPackage.Open(new MemoryStream(saved, writable: false));
            OfficeOpcRelationship actualMain = Assert.Single(actual.GetRelationships(null),
                edge => edge.Type == OfficeOpcPackage.OfficeDocumentRelationshipType);
            Assert.Equal(originalMain.Id, actualMain.Id);
            Assert.Equal(originalMain.TargetPartName, actualMain.TargetPartName);
            using var sdk = PresentationDocument.Open(new MemoryStream(saved, writable: false), false);
            Assert.Equal(type, sdk.DocumentType);
            using var reopened = PowerPointDocument.Load(new MemoryStream(saved, writable: false));
            Assert.Contains(Assert.Single(reopened.Slides).TextBoxes, box => box.Text == "Saved edit");
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Theory]
    [InlineData("missing video")]
    [InlineData("unexpected relationship child")]
    [InlineData("unexpected root attribute")]
    [InlineData("unexpected edge attribute")]
    [InlineData("wrong view content type")]
    public void SameSubtypeSaveRetainsSdkFailuresForInvalidParts(string invalidPart) {
        byte[] fixture = CreateFixture(PresentationDocumentType.Template, signed: false);
        using var original = OfficeOpcPackage.Open(new MemoryStream(fixture, writable: false));
        string main = original.GetRelatedPartName(null, OfficeOpcPackage.OfficeDocumentRelationshipType)!;
        string slide = original.GetRelatedPartName(main, OfficeOpcPackage.SlideRelationshipType)!;
        fixture = EditPackage(fixture, package => {
            const int ns = OfficeXmlNamespaces.PackageRelationships;
            OfficeXmlTree relationships = package.GetRelationships(slide);
            switch (invalidPart) {
                case "missing video":
                    package.AddRelationship(slide, OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "video", "/ppt/media/missing.mp4");
                    break;
                case "unexpected relationship child":
                    relationships.Append(relationships.Root, relationships.CreateElement(ns, "Unexpected"));
                    break;
                case "unexpected root attribute":
                    relationships.SetAttribute(relationships.Root, OfficeXmlNamespaces.None, "Unexpected", "1");
                    break;
                case "unexpected edge attribute":
                    relationships.SetAttribute(relationships.FindChild(relationships.Root, ns, "Relationship"),
                        OfficeXmlNamespaces.None, "Unexpected", "1");
                    break;
                case "wrong view content type":
                    string view = original.GetRelatedPartName(main, OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "viewProps")!;
                    package.SetOverride(view, "application/xml");
                    break;
            }
        });

        AssertTemplateSaveFailure(fixture);
    }

    [Fact]
    public void SameSubtypeSaveRetainsSdkValidationOfTheExistingCreatedDate() {
        byte[] fixture = CreateFixture(PresentationDocumentType.Template, signed: false);
        using var original = OfficeOpcPackage.Open(new MemoryStream(fixture, writable: false));
        string core = original.GetRelatedPartName(null, OfficeOpcPackage.CorePropertiesRelationshipType)!;
        fixture = EditPackage(fixture, package => {
            OfficeXmlTree tree = package.GetXml(core);
            const int ns = OfficeXmlNamespaces.DublinCoreTerms;
            int created = tree.FindChild(tree.Root, ns, "created");
            Assert.NotEqual(OfficeXmlTree.NoNode, tree.FindChild(tree.Root, ns, "modified"));
            tree.RemoveAttribute(created, package.Namespaces.GetId("http://www.w3.org/2001/XMLSchema-instance"), "type");
        });

        AssertTemplateSaveFailure(fixture);
    }

    [Fact]
    public void SameSubtypeSaveDropsAnUnreferencedVbaPartFromAMacroFreeTemplate() {
        const string vba = "/ppt/vbaProject.bin";
        byte[] fixture = EditPackage(CreateFixture(PresentationDocumentType.Template, signed: false),
            package => package.CreateBinaryPart(vba, "application/vnd.ms-office.vbaProject", new byte[] { 1, 2, 3, 4 }));
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-OrphanVba-" + Guid.NewGuid().ToString("N") + ".potx");
        try {
            using var presentation = PowerPointDocument.Load(new MemoryStream(fixture, writable: false));
            Assert.False(presentation.HasVbaProject);
            Assert.False(presentation.IsSdkLoaded);
            presentation.SaveCopy(path);

            using var saved = OfficeOpcPackage.Open(File.OpenRead(path));
            Assert.False(saved.PartExists(vba));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    private static void AssertTemplateSaveFailure(byte[] fixture) {
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-InvalidSubtype-" + Guid.NewGuid().ToString("N") + ".potx");
        try {
            using var sdk = PowerPointDocument.Load(new MemoryStream(fixture, writable: false));
            Exception expected = Assert.ThrowsAny<Exception>(() => {
                _ = sdk.OpenXmlDocument;
                sdk.SaveCopy(path);
            });
            using var presentation = PowerPointDocument.Load(new MemoryStream(fixture, writable: false));
            Assert.False(presentation.IsSdkLoaded);
            Assert.IsType(expected.GetType(), Record.Exception(() => presentation.SaveCopy(path)));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void SignedTemplatesKeepTheExistingLoadPath() {
        byte[] fixture = CreateFixture(PresentationDocumentType.Template, signed: true);

        using var presentation = PowerPointDocument.Load(new MemoryStream(fixture, writable: false));

        Assert.Equal(".potx", presentation.SourceFormatDescriptor.Extension);
        Assert.True(presentation.IsSdkLoaded);
    }

    [Fact]
    public void NoncanonicalMacroMimeKeepsTheSdkSourceFormatBeforeAndAfterAccess() {
        byte[] fixture = CreateFixture(PresentationDocumentType.MacroEnabledPresentation, signed: false);
        using var package = OfficeOpcPackage.Open(new MemoryStream(fixture, writable: false));
        string main = package.GetRelatedPartName(null, OfficeOpcPackage.OfficeDocumentRelationshipType)!;
        fixture = EditPackage(fixture, edit => edit.SetOverride(main,
            "application/vnd.ms-powerpoint.presentation.macroenabled.main+xml"));

        using var presentation = PowerPointDocument.Load(new MemoryStream(fixture, writable: false));

        Assert.Equal(PowerPointFileFormat.Pptx, presentation.SourceFormat);
        Assert.Equal(".pptx", presentation.SourceFormatDescriptor.Extension);
        _ = presentation.OpenXmlDocument;
        Assert.Equal(".pptx", presentation.SourceFormatDescriptor.Extension);
    }

    [Fact]
    public void MissingRootPropertiesTargetFailsDuringSubtypeLoad() {
        byte[] fixture = EditPackage(CreateFixture(PresentationDocumentType.Template, signed: false), package =>
            package.AddRelationship(null, OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "custom-properties", "/docProps/missing.xml"));

        Assert.Throws<InvalidOperationException>(() => PowerPointDocument.Load(new MemoryStream(fixture, writable: false)));
    }

    [Fact]
    public void InvalidRootTargetModeFailsDuringSubtypeLoad() {
        byte[] fixture = EditPackage(CreateFixture(PresentationDocumentType.Template, signed: false), package => {
            OfficeXmlTree relationships = package.GetRelationships(null);
            const int ns = OfficeXmlNamespaces.PackageRelationships;
            for (int node = relationships.FindChild(relationships.Root, ns, "Relationship");
                 node != OfficeXmlTree.NoNode;
                 node = relationships.FindNextSibling(node, ns, "Relationship")) {
                if (relationships.GetAttribute(node, OfficeXmlNamespaces.None, "Type") == OfficeOpcPackage.OfficeDocumentRelationshipType) {
                    relationships.SetAttribute(node, OfficeXmlNamespaces.None, "TargetMode", "Bogus");
                }
            }
        });

        Assert.ThrowsAny<System.Xml.XmlException>(() => PowerPointDocument.Load(new MemoryStream(fixture, writable: false)));
    }

    [Fact]
    public void InvalidApplicationPropertiesRootFailsDuringSubtypeLoad() {
        byte[] fixture = EditPackage(CreateFixture(PresentationDocumentType.Template, signed: false), package => {
            byte[] xml = Encoding.UTF8.GetBytes(
                "<Bogus xmlns='http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'/>");
            package.ReplacePart("/docProps/app.xml", xml, xml.Length);
        });

        Assert.Throws<InvalidDataException>(() => PowerPointDocument.Load(new MemoryStream(fixture, writable: false)));
    }

    [Fact]
    public void NegativeXmlPartLimitFailsDuringSubtypeLoad() {
        byte[] fixture = CreateFixture(PresentationDocumentType.Template, signed: false);

        Assert.Throws<ArgumentOutOfRangeException>(() => PowerPointDocument.Load(new MemoryStream(fixture, writable: false),
            new PowerPointLoadOptions {
                OpenSettings = new OfficeOpenXmlLoadSettings { MaxCharactersInPart = -1 }
            }));
    }

    private static byte[] EditPackage(byte[] data, Action<OfficeOpcEditablePackage> action) {
        using var package = OfficeOpcEditablePackage.Open(new MemoryStream(data, writable: false));
        action(package);
        using var output = new MemoryStream();
        package.Save(output);
        return output.ToArray();
    }

    private static byte[] CreateFixture(PresentationDocumentType type, bool signed) {
        using var document = PowerPointDocument.Create();
        document.AddSlide(PowerPointSlideLayoutType.Blank).AddTextBox("Subtype fixture");
        byte[] data = document.ToBytes();
        using var stream = new MemoryStream();
        stream.Write(data, 0, data.Length);
        stream.Position = 0;
        using (var package = PresentationDocument.Open(stream, true)) {
            package.ChangeDocumentType(type);
            if (signed) {
                var part = package.ExtendedFilePropertiesPart ?? package.AddExtendedFilePropertiesPart();
                part.Properties ??= new DocumentFormat.OpenXml.ExtendedProperties.Properties();
                part.Properties.DigitalSignature = new DocumentFormat.OpenXml.ExtendedProperties.DigitalSignature();
                part.Properties.Save();
            }
        }
        if (signed) return stream.ToArray();
        return AddDecoySlideList(stream.ToArray());
    }

    private static byte[] AddDecoySlideList(byte[] data) {
        using var package = OfficeOpcPackage.Open(new MemoryStream(data, writable: false));
        string main = package.GetRelatedPartName(null, OfficeOpcPackage.OfficeDocumentRelationshipType)!;
        using var edit = OfficeOpcEditablePackage.Open(new MemoryStream(data, writable: false));
        OfficeXmlTree tree = edit.GetXml(main);
        const int ns = OfficeXmlNamespaces.PresentationML;
        int extensions = tree.FindChild(tree.Root, ns, "extLst");
        if (extensions == OfficeXmlTree.NoNode) {
            extensions = tree.CreateElement(ns, "extLst");
            tree.Append(tree.Root, extensions);
        }
        int extension = tree.CreateElement(ns, "ext");
        tree.SetAttribute(extension, OfficeXmlNamespaces.None, "uri", "urn:officeimo:subtype-fixture");
        tree.Append(extensions, extension);
        int list = tree.CreateElement(ns, "sldIdLst");
        tree.Append(extension, list);
        int slide = tree.CreateElement(ns, "sldId");
        tree.SetAttribute(slide, OfficeXmlNamespaces.None, "id", "512");
        tree.SetAttribute(slide, OfficeXmlNamespaces.Relationships, "id", "rIdThisDoesNotExist");
        tree.Append(list, slide);
        using var output = new MemoryStream();
        edit.Save(output);
        return output.ToArray();
    }
}
