using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.OpenDocument.Testing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentFlatXmlImageTests {
    private static readonly byte[] TinyPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

    [Theory]
    [InlineData(OdfDocumentKind.Text)]
    [InlineData(OdfDocumentKind.Spreadsheet)]
    [InlineData(OdfDocumentKind.Presentation)]
    [InlineData(OdfDocumentKind.Graphics)]
    public void FlatImagePayloadPrecedesPreservedCaptionsAcrossDocumentKinds(OdfDocumentKind kind) {
        AssertCaptionProjection(CreateDocumentWithImage(kind), "content.xml");
    }

    [Fact]
    public void FlatHeaderImagePayloadPrecedesPreservedCaptions() {
        var document = OdtDocument.Create();
        document.PageLayout.Header.AddParagraph().AddImage(TinyPng, "caption.png",
            OdfLength.Centimeters(1), OdfLength.Centimeters(1));
        AssertCaptionProjection(document, "styles.xml");
    }

    [Theory]
    [InlineData(OdfDocumentKind.Graphics)]
    [InlineData(OdfDocumentKind.Presentation)]
    public void CommonFillImageBytesSurviveFlatExportAndRepackaging(OdfDocumentKind kind) {
        var document = CreateDocumentWithImage(kind);
        string path = (string)document.GetXml("content.xml").Descendants(OdfNamespaces.Draw + "image").Single().Attribute(OdfNamespaces.XLink + "href")!;
        document.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!.Add(new XElement(OdfNamespaces.Draw + "fill-image",
            new XAttribute(OdfNamespaces.Draw + "name", "Background"), new XAttribute(OdfNamespaces.XLink + "href", path)));
        document.MarkPartDirty("styles.xml"); string before = document.GetXml("styles.xml").ToString();
        using var flat = new MemoryStream(); var saved = document.SaveFlatXml(flat);
        Assert.DoesNotContain(path, saved.Report.LossyEntries); Assert.Equal(before, document.GetXml("styles.xml").ToString());
        flat.Position = 0; var xml = XDocument.Load(flat); var definition = xml.Descendants(OdfNamespaces.Draw + "fill-image").Single();
        Assert.Equal(TinyPng, Convert.FromBase64String(definition.Element(OdfNamespaces.Office + "binary-data")!.Value));
        Assert.Null(definition.Attribute(OdfNamespaces.Draw + "mime-type")); Assert.Null(definition.Attribute(OdfNamespaces.XLink + "href"));
        flat.Position = 0; var imported = OdfDocument.LoadFlatXml(flat);
        var reopened = OdfDocument.Load(new MemoryStream(imported.ToBytes()));
        string stored = (string)reopened.GetXml("styles.xml").Descendants(OdfNamespaces.Draw + "fill-image").Single().Attribute(OdfNamespaces.XLink + "href")!;
        Assert.Equal(TinyPng, reopened.GetPackageEntryBytes(stored));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FlatFillImagesEnforcePayloadValidationAndEntryLimits(bool excessive) {
        var document = OdgDocument.Create(); document.AddPage();
        string path = OdfImageStore.Add(document, TinyPng, "background.png");
        document.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!.Add(new XElement(OdfNamespaces.Draw + "fill-image",
            new XAttribute(OdfNamespaces.Draw + "name", "Background"), new XAttribute(OdfNamespaces.XLink + "href", path)));
        document.MarkPartDirty("styles.xml"); var xml = document.ToFlatXml();
        if (!excessive) xml.Descendants(OdfNamespaces.Office + "binary-data").Single().Value = "%%%%";
        using var stream = new MemoryStream(); xml.Save(stream); stream.Position = 0;
        var error = Assert.Throws<InvalidDataException>(() => OdfDocument.LoadFlatXml(stream,
            excessive ? new OdfLoadOptions { MaxEntryUncompressedBytes = TinyPng.Length - 1 } : null));
        Assert.Contains(excessive ? "MaxEntryUncompressedBytes" : "invalid base64", error.Message);
    }

    private static void AssertCaptionProjection(OdfDocument source, string partPath) {
        // Image captions can arrive from native producers even without a typed caption API.
        byte[] package = OdfTestPackageRewriter.Rewrite(source.ToBytes(), (name, bytes) => {
            if (name != partPath) return bytes;
            XDocument xml = XDocument.Load(new MemoryStream(bytes));
            xml.Descendants(OdfNamespaces.Draw + "image").Single().Add(
                new XElement(OdfNamespaces.Text + "p", "Before ",
                    new XElement(OdfNamespaces.Text + "span", "caption"),
                    new XElement(OdfNamespaces.Text + "line-break"), "after"),
                new XElement(OdfNamespaces.Text + "p", "Second caption"));
            using var output = new MemoryStream();
            xml.Save(output);
            return output.ToArray();
        });
        OdfDocument document = OdfDocument.Load(new MemoryStream(package));
        XElement image = document.GetXml(partPath).Descendants(OdfNamespaces.Draw + "image").Single();
        XElement[] captions = image.Elements().Select(element => new XElement(element)).ToArray();
        string imagePath = (string)image.Attribute(OdfNamespaces.XLink + "href")!;
        string originalXml = document.GetXml(partPath).ToString();

        AssertFlatImage(document.ToFlatXml(), captions);
        using var flat = new MemoryStream();
        OdfSaveResult saved = document.SaveFlatXml(flat);
        Assert.DoesNotContain(imagePath, saved.Report.LossyEntries);
        flat.Position = 0;
        AssertFlatImage(XDocument.Load(flat), captions);
        Assert.Equal(originalXml, document.GetXml(partPath).ToString());
        Assert.Equal(TinyPng, document.GetPackageEntryBytes(imagePath));

        flat.Position = 0;
        OdfDocument imported = OdfDocument.LoadFlatXml(flat);
        XElement importedImage = imported.GetXml(partPath).Descendants(OdfNamespaces.Draw + "image").Single();
        AssertCaptions(captions, importedImage.Elements().ToArray());
        Assert.Equal(TinyPng, imported.GetPackageEntryBytes((string)importedImage.Attribute(OdfNamespaces.XLink + "href")!));
        OdfDocument reopened = OdfDocument.Load(new MemoryStream(imported.ToBytes()));
        AssertFlatImage(reopened.ToFlatXml(), captions);
    }

    private static void AssertFlatImage(XDocument flat, XElement[] captions) {
        XElement image = flat.Descendants(OdfNamespaces.Draw + "image").Single();
        XElement binary = Assert.Single(image.Elements(OdfNamespaces.Office + "binary-data"));
        Assert.Equal(OdfNamespaces.Office + "binary-data", image.Elements().First().Name);
        Assert.Equal(TinyPng, Convert.FromBase64String(binary.Value));
        Assert.Equal("image/png", (string?)image.Attribute(OdfNamespaces.Draw + "mime-type"));
        Assert.Null(image.Attribute(OdfNamespaces.XLink + "href"));
        AssertCaptions(captions, image.Elements().Skip(1).ToArray());
    }

    private static void AssertCaptions(XElement[] expected, XElement[] actual) {
        Assert.Equal(expected.Length, actual.Length);
        for (int index = 0; index < expected.Length; index++)
            Assert.True(XNode.DeepEquals(expected[index], actual[index]));
    }

    private static OdfDocument CreateDocumentWithImage(OdfDocumentKind kind) {
        OdfRect bounds = OdfRect.FromCentimeters(1, 1, 1, 1);
        switch (kind) {
            case OdfDocumentKind.Text:
                var text = OdtDocument.Create();
                text.AddParagraph().AddImage(TinyPng, "caption.png", bounds.Width, bounds.Height);
                return text;
            case OdfDocumentKind.Spreadsheet:
                var spreadsheet = OdsDocument.Create();
                string path = OdfImageStore.Add(spreadsheet, TinyPng, "caption.png");
                var frame = new XElement(OdfNamespaces.Draw + "frame",
                    new XElement(OdfNamespaces.Draw + "image", new XAttribute(OdfNamespaces.XLink + "href", path),
                        new XAttribute(OdfNamespaces.XLink + "type", "simple"),
                        new XAttribute(OdfNamespaces.XLink + "show", "embed"),
                        new XAttribute(OdfNamespaces.XLink + "actuate", "onLoad")));
                OdfShape.ApplyBounds(frame, bounds);
                spreadsheet.AddSheet("Images").Element.AddFirst(new XElement(OdfNamespaces.Table + "shapes", frame));
                spreadsheet.MarkPartDirty("content.xml");
                return spreadsheet;
            case OdfDocumentKind.Presentation:
                var presentation = OdpPresentation.Create();
                presentation.AddSlide("Images").AddImage(TinyPng, "caption.png", bounds);
                return presentation;
            case OdfDocumentKind.Graphics:
                var drawing = OdgDocument.Create();
                drawing.AddPage().Shapes.AddImage(TinyPng, "caption.png", bounds);
                return drawing;
            default:
                throw new ArgumentOutOfRangeException(nameof(kind));
        }
    }
}
