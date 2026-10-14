using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public class VisioPackageMetadataContractTests {
    [Theory]
    [InlineData(VisioPackageType.Drawing)]
    [InlineData(VisioPackageType.Template)]
    [InlineData(VisioPackageType.Stencil)]
    public void GeneratedAndLegacyConvertedPackagesHaveReadableMetadata(VisioPackageType family) {
        VisioDocument document = VisioDocument.Create(family);
        document.Title = "Metadata control";
        document.Author = "Test author";
        if (family == VisioPackageType.Stencil) {
            document.RegisterMaster("Symbol", new VisioShape("1", 1, 1, 1, 1, "Label"));
        } else {
            document.AddPage("First page").Shapes.Add(new VisioShape("1", 1, 1, 1, 1, "Label"));
        }

        AssertMetadata(document.ToBytes(), family);
        VisioDocument converted = VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value), family).Value;
        AssertMetadata(converted.ToBytes(), family);
        VisioDocument reopened = VisioDocument.Load(new MemoryStream(converted.ToBytes()));
        Assert.Equal(family, reopened.PackageType);
        Assert.Equal(document.Title, reopened.Title);
        Assert.Equal(document.Author, reopened.Author);
        AssertMetadata(reopened.ToBytes(), family);
    }

    private static void AssertMetadata(byte[] bytes, VisioPackageType family) {
        using var archive = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        foreach (ZipArchiveEntry entry in archive.Entries.Where(entry => entry.FullName.EndsWith(".xml", StringComparison.Ordinal))) {
            using Stream input = entry.Open();
            Assert.NotNull(XDocument.Load(input).Root);
        }

        XNamespace extended = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
        XNamespace custom = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
        Assert.Equal(extended + "Properties", ReadXml(archive, "docProps/app.xml").Root!.Name);
        Assert.Equal(custom + "Properties", ReadXml(archive, "docProps/custom.xml").Root!.Name);

        XNamespace visio = "http://schemas.microsoft.com/office/visio/2012/main";
        XElement windows = ReadXml(archive, "visio/windows.xml").Root!;
        Assert.Equal(visio + "Windows", windows.Name);
        // Client dimensions describe a display area in unsigned integer units,
        // not the drawing's fractional inch dimensions.
        Assert.Null(windows.Attribute("ClientWidth"));
        Assert.Null(windows.Attribute("ClientHeight"));
        if (family == VisioPackageType.Stencil) {
            Assert.Empty(windows.Elements());
        } else {
            XElement window = Assert.Single(windows.Elements(visio + "Window"));
            Assert.True(uint.TryParse((string?)window.Attribute("ID"), out _));
            Assert.Equal("Drawing", (string?)window.Attribute("WindowType"));
            Assert.Equal("Page", (string?)window.Attribute("ContainerType"));
            string pageId = ReadXml(archive, "visio/pages/pages.xml").Root!.Elements(visio + "Page").First().Attribute("ID")!.Value;
            Assert.Equal(pageId, (string?)window.Attribute("Page"));
            Assert.Equal(pageId, (string?)window.Attribute("Container"));
        }

        // A generated package has no thumbnail image. Do not advertise an empty EMF.
        Assert.Null(archive.GetEntry("docProps/thumbnail.emf"));
        XNamespace rels = "http://schemas.openxmlformats.org/package/2006/relationships";
        Assert.DoesNotContain(ReadXml(archive, "_rels/.rels").Root!.Elements(rels + "Relationship"),
            row => ((string?)row.Attribute("Type"))?.EndsWith("/thumbnail", StringComparison.Ordinal) == true);
        XNamespace types = "http://schemas.openxmlformats.org/package/2006/content-types";
        Assert.DoesNotContain(ReadXml(archive, "[Content_Types].xml").Root!.Elements(types + "Override"),
            row => (string?)row.Attribute("PartName") == "/docProps/thumbnail.emf");
    }

    private static XDocument ReadXml(ZipArchive archive, string name) {
        using Stream input = archive.GetEntry(name)!.Open();
        return XDocument.Load(input);
    }
}
