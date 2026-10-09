using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentOdgImageFormatTests {
    private static readonly byte[] Png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

    [Theory]
    [InlineData(OfficeImageFormat.Svg, false)]
    [InlineData(OfficeImageFormat.Svg, true)]
    [InlineData(OfficeImageFormat.Wmf, false)]
    [InlineData(OfficeImageFormat.Emf, false)]
    public void UnprojectedCommonFillResourcesRetainTheirBytesAcrossFlatAndPackageSaves(OfficeImageFormat format, bool dimensionless) {
        var document = Drawing(); byte[] bytes = Resource(format);
        if (dimensionless) bytes = WithoutIntrinsicDimensions(bytes);
        string path = AddResource(document, bytes, format);
        AddFill(document, path); string source = document.GetXml("styles.xml").ToString();
        using var flat = new MemoryStream(); var saved = document.SaveFlatXml(flat);
        Assert.DoesNotContain(path, saved.Report.LossyEntries); Assert.Equal(source, document.GetXml("styles.xml").ToString());
        flat.Position = 0; var reopened = OdgDocument.LoadFlatXml(flat);
        using var packaged = new MemoryStream(reopened.ToBytes(new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.PreserveSource }));
        foreach (var read in new[] { reopened, OdgDocument.Load(packaged) }) {
            string stored = (string)read.GetXml("styles.xml").Descendants(OdfNamespaces.Draw + "fill-image").Single().Attribute(OdfNamespaces.XLink + "href")!;
            Assert.Equal(bytes, read.GetPackageEntryBytes(stored));
            Assert.Equal(OfficeImageInfo.GetMimeType(format), read.Package.GetRequiredEntry(stored).MediaType);
        }
    }

    [Theory]
    [InlineData(OfficeImageFormat.Wmf, false)]
    [InlineData(OfficeImageFormat.Wmf, true)]
    [InlineData(OfficeImageFormat.Emf, false)]
    [InlineData(OfficeImageFormat.Emf, true)]
    [InlineData(OfficeImageFormat.Bmp, false)]
    [InlineData(OfficeImageFormat.Bmp, true)]
    [InlineData(OfficeImageFormat.Svg, false)]
    public void ImagesOutsideTheSharedRenderProfileRemainPreservedButCannotPassStrictProjection(OfficeImageFormat format, bool background) {
        var document = Drawing(); byte[] bytes = Resource(format);
        if (format == OfficeImageFormat.Svg) bytes = WithoutIntrinsicDimensions(bytes);
        string path = AddResource(document, bytes, format);
        Assert.True(OfficeImageReader.TryValidateContent(bytes, null, out OfficeImageInfo info)); Assert.Equal(format, info.Format);
        string feature;
        if (background) {
            AddFill(document, path);
            var properties = new XElement(OdfNamespaces.Style + "drawing-page-properties",
                new XAttribute(OdfNamespaces.Draw + "fill", "bitmap"), new XAttribute(OdfNamespaces.Draw + "fill-image-name", "Native"),
                new XAttribute(OdfNamespaces.Style + "repeat", "stretch"));
            document.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "automatic-styles")!.Add(new XElement(OdfNamespaces.Style + "style",
                new XAttribute(OdfNamespaces.Style + "name", "Background"), new XAttribute(OdfNamespaces.Style + "family", "drawing-page"), properties));
            document.Pages[0].Master!.SetAttributeValue(OdfNamespaces.Draw + "style-name", "Background"); document.MarkPartDirty("styles.xml");
            feature = "page-background";
        } else {
            var frame = document.Pages[0].Shapes.AddImage(Png, "caption.png", new OdfRect(P(10), P(10), P(20), P(10)), "Metafile");
            frame.Element.Element(OdfNamespaces.Draw + "image")!.SetAttributeValue(OdfNamespaces.XLink + "href", path);
            frame.AddParagraph("Retained caption"); document.MarkPartDirty("content.xml"); feature = "shape:Metafile";
        }
        using var package = new MemoryStream(document.ToBytes(new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.PreserveSource }));
        foreach (var read in new[] { document, OdgDocument.Load(package) }) {
            string[] before = Parts(read); var result = read.Pages[0].ToDrawing();
            Assert.Empty(result.Value.Elements.OfType<OfficeDrawingImage>());
            Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == feature && mapping.Status == OdfConversionMappingStatus.Skipped);
            if (!background) Assert.Equal("Retained caption", Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>()).PlainText);
            Assert.Equal(bytes, read.GetPackageEntryBytes(path)); Assert.Equal(before, Parts(read));
            Assert.Throws<OdfConversionLossException>(() => read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        }
    }

    [Fact]
    public void PartiallySupportedSvgPaintRetainsItsCaptionAndReportsContentLoss() {
        var document = Drawing(); byte[] bytes = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'><script>alert(1)</script><rect width='10' height='10' fill='red'/></svg>");
        var frame = document.Pages[0].Shapes.AddImage(bytes, "partial.svg", new OdfRect(P(10), P(10), P(20), P(10)), "Partial");
        frame.AddParagraph("Retained caption"); string[] before = Parts(document); var result = document.Pages[0].ToDrawing();
        Assert.Single(result.Value.Elements.OfType<OfficeDrawingImage>());
        Assert.Equal("Retained caption", Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>()).PlainText);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "shape:Partial:svg-content" && mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Equal(before, Parts(document));
        Assert.Throws<OdfConversionLossException>(() => document.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CommonSvgFillPreservationStillEnforcesXmlAndViewportBounds(bool oversized) {
        var document = Drawing(); var flat = document.ToFlatXml();
        byte[] bytes = Encoding.UTF8.GetBytes(oversized
            ? "<svg xmlns='http://www.w3.org/2000/svg' width='9000' height='10'><rect width='10' height='10'/></svg>"
            : "<!DOCTYPE svg [<!ENTITY sample 'text'>]><svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'><text>&sample;</text></svg>");
        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(bytes));
        if (oversized) Assert.True(OfficeImageReader.TryValidateContent(bytes, null, out _));
        flat.Root!.Element(OdfNamespaces.Office + "styles")!.Add(new XElement(OdfNamespaces.Draw + "fill-image",
            new XAttribute(OdfNamespaces.Draw + "name", "Native"), new XElement(OdfNamespaces.Office + "binary-data", Convert.ToBase64String(bytes))));
        using var stream = new MemoryStream(); flat.Save(stream); stream.Position = 0;
        Assert.Throws<InvalidDataException>(() => OdgDocument.LoadFlatXml(stream));
    }

    [Fact]
    public void CommonSvgFillPreservationRetainsAnIgnoredLegacyDoctype() {
        byte[] bytes = Encoding.UTF8.GetBytes("<!DOCTYPE svg SYSTEM 'file:///officeimo-unused-test-svg.dtd'>"
            + "<svg xmlns='http://www.w3.org/2000/svg' width='20' height='10'><rect width='20' height='10'/></svg>");
        Assert.True(OfficeSvgDrawingReader.IsWithinSafetyLimits(bytes));
        var document = Drawing(); string path = AddResource(document, bytes, OfficeImageFormat.Svg);
        AddFill(document, path);
        using var flat = new MemoryStream(); document.SaveFlatXml(flat); flat.Position = 0;
        var reopened = OdgDocument.LoadFlatXml(flat);
        using var packaged = new MemoryStream(reopened.ToBytes(new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.PreserveSource }));
        foreach (var read in new[] { reopened, OdgDocument.Load(packaged) }) {
            string stored = (string)read.GetXml("styles.xml").Descendants(OdfNamespaces.Draw + "fill-image").Single().Attribute(OdfNamespaces.XLink + "href")!;
            Assert.Equal(bytes, read.GetPackageEntryBytes(stored));
        }
    }

    private static OdgDocument Drawing() { var document = OdgDocument.Create(); document.AddPage("Resources", P(100), P(80)); return document; }
    private static OdfLength P(double value) => OdfLength.Points(value);
    private static byte[] WithoutIntrinsicDimensions(byte[] svg) => Encoding.UTF8.GetBytes(Encoding.UTF8.GetString(svg)
        .Replace("<svg xmlns='http://www.w3.org/2000/svg' width='20' height='10'", "<svg xmlns='http://www.w3.org/2000/svg'"));
    private static string[] Parts(OdgDocument document) => new[] { document.GetXml("content.xml").ToString(), document.GetXml("styles.xml").ToString() };
    private static string AddResource(OdgDocument document, byte[] bytes, OfficeImageFormat format) {
        string path = "Pictures/native" + OfficeImageInfo.GetDefaultExtension(format);
        document.Package.AddOrReplaceEntry(path, bytes, OfficeImageInfo.GetMimeType(format)); return path;
    }
    private static void AddFill(OdgDocument document, string path) {
        document.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!.Add(new XElement(OdfNamespaces.Draw + "fill-image",
            new XAttribute(OdfNamespaces.Draw + "name", "Native"), new XAttribute(OdfNamespaces.XLink + "href", path)));
        document.MarkPartDirty("styles.xml");
    }
    private static byte[] Resource(OfficeImageFormat format) => format switch {
        OfficeImageFormat.Svg => Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='20' height='10'><metadata><producer xmlns='urn:example:producer'>Retain this producer metadata</producer></metadata><rect width='20' height='10' fill='#e07020'/></svg>"),
        // Complete bounded metafiles use the shared image-reader identity controls.
        OfficeImageFormat.Wmf => Convert.FromBase64String("183GmgAAAAAAAEALoAWgBQAAAABRXAEACQAAAxEAAAAAAAUAAAAAAAUAAAABAgAAAAADAAAAAAA="),
        OfficeImageFormat.Emf => Convert.FromBase64String("AQAAAFgAAAAAAAAAAAAAABQAAAAKAAAAAAAAAAAAAAAAAAAAAAAAACBFTUYAAAEAbAAAAAIAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA4AAAAUAAAAAAAAAAAAAAAUAAAA"),
        OfficeImageFormat.Bmp => OversizedTranscodeBitmap(),
        _ => throw new ArgumentOutOfRangeException(nameof(format))
    };
    private static byte[] OversizedTranscodeBitmap() {
        const int width = 4001, height = 2000; int stride = (width * 3 + 3) & ~3;
        byte[] bytes = new byte[54 + stride * height]; bytes[0] = 66; bytes[1] = 77;
        Write(2, bytes.Length); Write(10, 54); Write(14, 40); Write(18, width); Write(22, height);
        bytes[26] = 1; bytes[28] = 24; Write(34, stride * height);
        // Deterministic pixel data keeps package compression inside the normal load ratio.
        uint seed = 12345;
        for (int index = 54; index < bytes.Length; index++) { seed = unchecked(seed * 1664525 + 1013904223); bytes[index] = (byte)(seed >> 24); }
        return bytes;
        void Write(int offset, int value) { for (int index = 0; index < 4; index++) bytes[offset + index] = (byte)(value >> (index * 8)); }
    }
}
