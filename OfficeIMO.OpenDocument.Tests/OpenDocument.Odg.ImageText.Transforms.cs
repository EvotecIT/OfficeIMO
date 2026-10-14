using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentOdgTransformedImageTextTests {
    private const string Family = "Translated Caption Proof";

    [Theory]
    [InlineData("center", false, false)]
    [InlineData("right", false, false)]
    [InlineData("center", true, false)]
    [InlineData("right", true, false)]
    [InlineData("center", false, true)]
    [InlineData("right", false, true)]
    [InlineData("center", true, true)]
    [InlineData("right", true, true)]
    public void TranslatedImageCaptionMatchesItsAbsoluteFrameThroughNativeRoundTrips(string alignment, bool fitting, bool grouped) {
        string caption = fitting ? "Fit" : "OVERWIDE CAPTION TEXT";
        byte[] resource = OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.CornflowerBlue));
        byte[] font = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "FontReferences", "SourceSansPro-Regular.otf"));
        var actualDocument = Create(caption, alignment, grouped, true, resource);
        var expectedDocument = Create(caption, alignment, false, false, resource);
        OfficeDrawing expected = expectedDocument.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Value;
        expected.Fonts.Add(Family, font);
        OfficeRasterImage reference = OfficeDrawingRasterRenderer.Render(expected);
        bool outsideInk = Enumerable.Range(0, reference.Height).Any(y => Enumerable.Range(0, reference.Width)
            .Any(x => (x < 140 || x >= 160) && reference.GetPixel(x, y).A > 0));
        Assert.Equal(!fitting, outsideInk);
        foreach (OdgDocument read in RoundTrips(actualDocument)) {
            OdgShape image = grouped ? read.Pages[0].Shapes[0].Children[0] : read.Pages[0].Shapes[0];
            string content = read.GetXml("content.xml").ToString(), styles = read.GetXml("styles.xml").ToString();
            OfficeDrawing actual = read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Value;
            // Post-projection registration is part of the drawing's render-time font contract.
            actual.Fonts.Add(Family, font);
            OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(actual);
            Assert.Equal((reference.Width, reference.Height), (raster.Width, raster.Height));
            for (int y = 0; y < raster.Height; y++) for (int x = 0; x < raster.Width; x++) Assert.Equal(reference.GetPixel(x, y), raster.GetPixel(x, y));
            var effect = Assert.IsType<OfficeDrawingEffectGroup>(Assert.Single(actual.Elements));
            var frame = Assert.IsType<OfficeDrawingRichText>(effect.Drawing.Elements.Last());
            Assert.Equal((0D, 0D, 20D, 40D), (frame.X, frame.Y, frame.Width, frame.Height));
            Assert.False(frame.WrapText);
            Assert.Equal(content, read.GetXml("content.xml").ToString());
            Assert.Equal(styles, read.GetXml("styles.xml").ToString());
            Assert.Equal(resource, image.GetImageBytes());
            Assert.Equal(caption, image.Text);
        }
    }

    private static OdgDocument Create(string caption, string alignment, bool grouped, bool translated, byte[] resource) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        page.Width = OdfLength.Points(300); page.Height = OdfLength.Points(160);
        OdgShapes shapes = page.Shapes; OdgShape? group = null;
        if (grouped) { group = shapes.AddGroup("caption-group"); shapes = group.Children; }
        var image = shapes.AddImage(resource, "caption.png", new OdfRect(OdfLength.Points(translated ? 100 : 140),
            OdfLength.Points(40), OdfLength.Points(20), OdfLength.Points(40)), "caption-image");
        image.FontSize = OdfLength.Points(10);
        string name = (string)image.ToXml().Attribute(OdfNamespaces.Draw + "style-name")!;
        document.Styles.FindInPart(OdfStyleFamily.Graphic, name, "content.xml")!
            .SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "wrap-option", "no-wrap");
        var paragraph = image.AddParagraph(caption); paragraph.FontFamily = Family;
        paragraph.FontSize = OdfLength.Points(10); paragraph.TextAlign = alignment;
        if (translated) {
            if (group == null) image.Transform = "translate(40pt 0pt)";
            else group.TransformChildren("translate(40pt 0pt)");
        }
        return document;
    }

    private static IEnumerable<OdgDocument> RoundTrips(OdgDocument document) {
        yield return document;
        yield return OdgDocument.Load(new MemoryStream(document.ToBytes()));
        using var flat = new MemoryStream(); document.SaveFlatXml(flat); flat.Position = 0;
        yield return OdgDocument.LoadFlatXml(flat);
    }
}
