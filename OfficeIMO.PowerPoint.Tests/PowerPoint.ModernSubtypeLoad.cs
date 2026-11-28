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
        using var output = PresentationDocument.Open(new MemoryStream(copy, writable: false), false);
        using var reopened = PowerPointDocument.Load(new MemoryStream(copy, writable: false));

        Assert.Equal(PresentationDocumentType.Presentation, output.DocumentType);
        Assert.Equal(extension, presentation.SourceFormatDescriptor.Extension);
        Assert.Contains(Assert.Single(reopened.Slides).TextBoxes, box => box.Text == "Pending edit");
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
