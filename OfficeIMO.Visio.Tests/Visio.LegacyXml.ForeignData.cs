using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public class VisioLegacyForeignDataTests {
    private static readonly XNamespace Ns = "http://schemas.microsoft.com/visio/2003/core";
    private static readonly byte[] Pixel = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PreservesForeignBytesThroughXmlPackageAndTypedEdits(bool stencil) {
        var family = stencil ? VisioPackageType.Stencil : VisioPackageType.Drawing;
        using var input = Fixture(stencil, false);
        var imported = VisioDocument.LoadLegacyXml(input, family);
        Assert.DoesNotContain(imported.Report.FidelityDiagnostics, d => d.Code == "VDX_FOREIGN_DATA");
        if (stencil) imported.Value.Masters.Single().Shape.Text = "Edited master";
        else imported.Value.Pages[0].Shapes[0].Text = "Edited diagram";
        byte[] package = imported.Value.ToBytes();
        using (var archive = new ZipArchive(new MemoryStream(package))) {
            var binary = Assert.Single(archive.Entries, entry => entry.FullName.StartsWith("visio/media/foreign", StringComparison.Ordinal));
            using var bytes = new MemoryStream(); using (var stream = binary.Open()) stream.CopyTo(bytes);
            Assert.Equal(Pixel, bytes.ToArray());
        }
        var reopened = VisioDocument.Load(new MemoryStream(package));
        var legacy = reopened.ToLegacyXmlResult();
        Assert.DoesNotContain(legacy.Report.FidelityDiagnostics, d => d.Code == "VDX_PACKAGE_PART" && d.Location!.Contains("media"));
        var xml = XDocument.Load(new MemoryStream(legacy.Value));
        var foreign = Assert.Single(xml.Descendants(Ns + "ForeignData"));
        Assert.Equal("Foreign", (string?)foreign.Parent!.Attribute("Type"));
        Assert.Equal("PNG", (string?)foreign.Attribute("CompressionType"));
        Assert.Equal(Pixel, Convert.FromBase64String(foreign.Value));
        Assert.Contains(stencil ? "Edited master" : "Edited diagram", xml.ToString());
        Assert.Equal(Pixel, Convert.FromBase64String(XDocument.Load(new MemoryStream(VisioDocument.LoadLegacyXml(new MemoryStream(legacy.Value), family).Value.ToLegacyXmlResult().Value)).Descendants(Ns + "ForeignData").Single().Value));
    }

    [Theory]
    [InlineData(VisioPackageType.Stencil, false)]
    [InlineData(VisioPackageType.Template, true)]
    public void ImageBearingMasterEditsKeepUntouchedFormulasAndRemoveDeletedResources(VisioPackageType family, bool embedded) {
        using var fixture = Fixture(true, embedded);
        var input = XDocument.Load(fixture);
        XElement sourceShape = input.Descendants(Ns + "Master").Single().Element(Ns + "Shapes")!.Element(Ns + "Shape")!;
        sourceShape.Element(Ns + "XForm")!.Element(Ns + "Height")!.SetAttributeValue("F", "GUARD(3)");
        XElement child = sourceShape.Element(Ns + "Shapes")!.Element(Ns + "Shape")!;
        child.Element(Ns + "XForm")!.Element(Ns + "Width")!.SetAttributeValue("F", "GUARD(2)");
        child.Add(new XElement(Ns + "Connection", new XAttribute("IX", "7"),
            new XElement(Ns + "X", new XAttribute("F", "Width*0.5"), "1"),
            new XElement(Ns + "Y", "0.5")));
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(input.ToString()));
        var document = VisioDocument.LoadLegacyXml(stream, family).Value;
        var root = document.Masters.Single().Shape;
        root.Text = "Edited image master";
        root.Width = 4;
        foreach (var candidate in new[] { document, VisioDocument.Load(new MemoryStream(document.ToBytes())) }) {
            var xml = XDocument.Load(new MemoryStream(candidate.ToLegacyXmlResult().Value));
            XElement savedRoot = xml.Descendants(Ns + "Master").Single().Element(Ns + "Shapes")!.Element(Ns + "Shape")!;
            Assert.Equal("4", savedRoot.Element(Ns + "XForm")!.Element(Ns + "Width")!.Value);
            Assert.Equal("GUARD(3)", (string?)savedRoot.Element(Ns + "XForm")!.Element(Ns + "Height")!.Attribute("F"));
            XElement savedChild = savedRoot.Element(Ns + "Shapes")!.Element(Ns + "Shape")!;
            Assert.Equal("GUARD(2)", (string?)savedChild.Element(Ns + "XForm")!.Element(Ns + "Width")!.Attribute("F"));
            Assert.Equal("Width*0.5", (string?)savedChild.Element(Ns + "Connection")!.Element(Ns + "X")!.Attribute("F"));
            Assert.Equal("Edited image master", savedRoot.Element(Ns + "Text")!.Value);
            Assert.Equal(Pixel, Convert.FromBase64String(savedChild.Element(Ns + "ForeignData")!.Value));
        }
        // Membership edits use the same native merge, and orphaned bytes must not reappear.
        root.Children.Clear();
        using var package = new ZipArchive(new MemoryStream(document.ToBytes()));
        Assert.DoesNotContain(package.Entries, entry => entry.FullName.StartsWith("visio/media/foreign") || entry.FullName.StartsWith("visio/embeddings/foreign"));
        Assert.Empty(XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value)).Descendants(Ns + "ForeignData"));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void ChangedMasterExtentsRetainTheModeledLocalPinWhenNativeCellsAreAbsent(bool explicitPin, bool foreignContent) {
        using var fixture = Fixture(true, false);
        var input = XDocument.Load(fixture);
        if (!foreignContent) input.Descendants(Ns + "ForeignData").Single().Parent!.Remove();
        var transform = input.Descendants(Ns + "Master").Single().Element(Ns + "Shapes")!.Element(Ns + "Shape")!.Element(Ns + "XForm")!;
        if (explicitPin) transform.Add(new XElement(Ns + "LocPinX", "0"), new XElement(Ns + "LocPinY", "0"));
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(input.ToString()));
        var document = VisioDocument.LoadLegacyXml(stream, VisioPackageType.Stencil).Value;
        var shape = document.Masters.Single().Shape;
        double localX = shape.LocPinX, localY = shape.LocPinY;
        shape.Width = 4; shape.Height = 5;
        var reopened = VisioDocument.Load(new MemoryStream(document.ToBytes())).Masters.Single().Shape;
        Assert.Equal(localX, reopened.LocPinX);
        Assert.Equal(localY, reopened.LocPinY);
    }

    [Fact]
    public void AppliesEmbeddedPayloadPolicyBeforeLoadingForeignObjects() {
        using var input = Fixture(false, true);
        Assert.ThrowsAny<Exception>(() => VisioDocument.LoadLegacyXml(input, options: VisioLoadOptions.UntrustedDefaults));
        input.Position = 0;
        var imported = VisioDocument.LoadLegacyXml(input);
        using var disguisedInput = Fixture(false, false);
        var disguised = XDocument.Load(disguisedInput);
        disguised.Descendants(Ns + "ForeignData").Single().Value = Convert.ToBase64String(Encoding.ASCII.GetBytes("opaque object bytes"));
        using var disguisedStream = new MemoryStream(Encoding.UTF8.GetBytes(disguised.ToString()));
        Assert.ThrowsAny<Exception>(() => VisioDocument.LoadLegacyXml(disguisedStream, options: VisioLoadOptions.UntrustedDefaults));
        Assert.Equal(Pixel, Convert.FromBase64String(XDocument.Load(new MemoryStream(imported.Value.ToLegacyXmlResult().Value)).Descendants(Ns + "ForeignData").Single().Value));
    }

    [Fact]
    public void RejectsExcessForeignResourcesBeforeCreatingAnUnboundedPackage() {
        using var input = Fixture(false, false);
        XDocument xml = XDocument.Load(input);
        XElement shape = xml.Descendants(Ns + "ForeignData").Single().Parent!;
        XElement shapes = shape.Parent!;
        for (int index = 0; index < 1024; index++) {
            var copy = new XElement(shape); copy.SetAttributeValue("ID", (index + 3).ToString()); shapes.Add(copy);
        }
        using var excessive = new MemoryStream(Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting)));
        Assert.Throws<InvalidDataException>(() => VisioDocument.LoadLegacyXml(excessive));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void ForeignBytesFollowMovedShapesAndAreRemovedWithTheirLastReference(bool crossDocument, bool occupied) {
        using var sourceXml = Fixture(false, false);
        XDocument xml = XDocument.Load(sourceXml);
        XElement secondPage = new XElement(xml.Descendants(Ns + "Page").Single());
        secondPage.SetAttributeValue("ID", "1"); secondPage.SetAttributeValue("NameU", "Destination");
        foreach (XElement shape in secondPage.Descendants(Ns + "Shape")) shape.SetAttributeValue("ID", (int)shape.Attribute("ID")! + 10);
        byte[] otherBytes = Encoding.UTF8.GetBytes("another embedded object");
        secondPage.Descendants(Ns + "ForeignData").Single().Value = Convert.ToBase64String(otherBytes);
        if (!occupied) secondPage.Element(Ns + "Shapes")!.RemoveNodes();
        xml.Root!.Element(Ns + "Pages")!.Add(secondPage);
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(xml.ToString()));
        var source = VisioDocument.LoadLegacyXml(input).Value;
        foreach (VisioPage page in source.Pages) page.AutoResizeDrawing = false; // Auto-size is a separate unsupported legacy setting.
        var destination = crossDocument ? VisioDocument.Load(new MemoryStream(source.ToBytes())) : source;
        var moved = source.Pages[0].Shapes[0];
        source.Pages[0].Shapes.Remove(moved); destination.Pages[1].Shapes.Add(moved);
        if (crossDocument) destination.Pages[0].Shapes.Clear();
        XDocument saved = XDocument.Load(new MemoryStream(destination.ToLegacyXmlResult().Value));
        var payloads = saved.Descendants(Ns + "Page").Last().Descendants(Ns + "ForeignData").Select(foreign => Convert.FromBase64String(foreign.Value)).ToList();
        Assert.Contains(payloads, bytes => bytes.SequenceEqual(Pixel));
        Assert.Equal(occupied ? 2 : 1, payloads.Count);
        if (occupied) Assert.Contains(payloads, bytes => bytes.SequenceEqual(otherBytes));
        // Moving a child out of the group exercises the same payload ownership contract.
        var child = moved.Children[0]; moved.Children.Remove(child); destination.Pages[1].Shapes.Add(child);
        destination.Pages[1].Shapes.Remove(moved);
        Assert.Contains(XDocument.Load(new MemoryStream(destination.ToLegacyXmlResult().Value)).Descendants(Ns + "ForeignData"), foreign => Convert.FromBase64String(foreign.Value).SequenceEqual(Pixel));
        destination.Pages[1].Shapes.Clear();
        using var deletedPackage = new ZipArchive(new MemoryStream(destination.ToBytes()));
        Assert.DoesNotContain(deletedPackage.Entries, entry => entry.FullName.StartsWith("visio/media/foreign") || entry.FullName.StartsWith("visio/embeddings/foreign"));
        var deletionResult = destination.ToLegacyXmlResult();
        Assert.False(deletionResult.Report.HasOmissions, string.Join("; ", deletionResult.Report.FidelityDiagnostics.Select(diagnostic => diagnostic.Code + " " + diagnostic.Location + " " + diagnostic.Message)));
        using var output = new MemoryStream(); destination.SaveLegacyXml(output);
        Assert.Empty(XDocument.Load(new MemoryStream(output.ToArray())).Descendants(Ns + "ForeignData"));
    }

    [Fact]
    public void DuplicatedShapesAndPagesKeepTheirOwnForeignContentAfterOriginalDeletion() {
        using var input = Fixture(false, false);
        var document = VisioDocument.LoadLegacyXml(input).Value;
        var page = document.Pages[0];
        page.DuplicateShapes(new[] { page.Shapes[0] });
        page.Shapes.RemoveAt(0);
        var duplicate = document.DuplicatePage(page);
        page.Shapes.Clear();
        XDocument xml = XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value));
        Assert.Equal(Pixel, Convert.FromBase64String(Assert.Single(xml.Descendants(Ns + "ForeignData")).Value));
        Assert.Equal("0.25", Assert.Single(xml.Descendants(Ns + "Foreign")).Element(Ns + "ImgOffsetX")!.Value);
        duplicate.Shapes.Clear();
        using var archive = new ZipArchive(new MemoryStream(document.ToBytes()));
        Assert.DoesNotContain(archive.Entries, entry => entry.FullName.StartsWith("visio/media/foreign"));
    }

    private static MemoryStream Fixture(bool stencil, bool embedded) {
        var foreignShape = new XElement(Ns + "Shape", new XAttribute("ID", "2"), new XAttribute("Type", "Foreign"),
            new XElement(Ns + "XForm", new XElement(Ns + "PinX", 1), new XElement(Ns + "PinY", 1), new XElement(Ns + "Width", 2), new XElement(Ns + "Height", 1)),
            new XElement(Ns + "Foreign", new XElement(Ns + "ImgOffsetX", "0.25"), new XElement(Ns + "ImgOffsetY", "0.1"), new XElement(Ns + "ImgWidth", "2"), new XElement(Ns + "ImgHeight", "1")),
            new XElement(Ns + "ForeignData", new XAttribute("ForeignType", embedded ? "Object" : "Bitmap"), new XAttribute("CompressionType", "PNG"), Convert.ToBase64String(Pixel)));
        var shape = new XElement(Ns + "Shape", new XAttribute("ID", "1"), new XAttribute("Type", "Group"),
            new XElement(Ns + "XForm", new XElement(Ns + "PinX", 2), new XElement(Ns + "PinY", 2), new XElement(Ns + "Width", 3), new XElement(Ns + "Height", 3)),
            new XElement(Ns + "Text", "Original"), new XElement(Ns + "Shapes", foreignShape));
        var item = new XElement(Ns + (stencil ? "Master" : "Page"), new XAttribute("ID", "0"), new XAttribute("NameU", "Diagram"), new XElement(Ns + "Shapes", shape));
        var xml = new XDocument(new XElement(Ns + "VisioDocument", new XElement(Ns + (stencil ? "Masters" : "Pages"), item)));
        return new MemoryStream(Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting)));
    }
}
