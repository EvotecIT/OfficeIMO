using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.OpenDocument.Testing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentOdgImageTextTests {
    [Fact]
    public void ProducerImageCaptionsUseTheExistingTextStoryWithoutChangingPayloads() {
        byte[] bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", "libreoffice-network-connectors.odg"));
        var document = OdgDocument.Load(new MemoryStream(OdfTestPackageRewriter.Rewrite(bytes)));
        string content = document.GetXml("content.xml").ToString();
        string styles = document.GetXml("styles.xml").ToString();
        OdgShape image = document.Pages[0].Shapes.First(shape => shape.Name == "xenarthra");
        byte[] resource = image.GetImageBytes()!;
        Assert.Equal("Xenarthra", image.Text);
        Assert.Equal("Xenarthra", Assert.Single(image.Paragraphs).Text);
        var projection = document.Pages[0].ToDrawing();
        Assert.Contains(Texts(projection.Value), text => text.PlainText == "Xenarthra");
        Assert.Equal(content, document.GetXml("content.xml").ToString());
        Assert.Equal(styles, document.GetXml("styles.xml").ToString());
        Assert.Equal(resource, image.GetImageBytes());
    }

    [Fact]
    public void ImageCaptionEditingRetainsResourceRichRunsAndOtherTextStories() {
        var (document, image, resource) = CreateImage();
        var paragraph = image.AddParagraph("Before ");
        paragraph.AddRun("Bold").Bold = true;
        paragraph.AddHyperlink("Link", "https://example.test/docs");
        var field = paragraph.AddField(OdfTextFieldKind.Date, "2026-10-07"); field.IsFixed = true;
        paragraph.Element.Add(new XElement(OdfNamespaces.Office + "annotation", new XElement(OdfNamespaces.Text + "p", "Hidden note")));
        image.AddList(true).AddItem("Item");
        foreach (OdgDocument read in RoundTrips(document)) {
            OdgShape saved = read.Pages[0].Shapes[0];
            Assert.Equal("Before BoldLink2026-10-07\nItem", saved.Text);
            Assert.True(saved.Paragraphs[0].Runs.Single(run => run.Text == "Bold").Bold);
            Assert.Equal("https://example.test/docs", Assert.Single(saved.Paragraphs[0].Hyperlinks).Href);
            Assert.True(Assert.Single(saved.Paragraphs[0].Fields).IsFixed);
            Assert.Equal("Hidden note", saved.Paragraphs[0].ToXml().Element(OdfNamespaces.Office + "annotation")!.Value);
            Assert.Equal(resource, saved.GetImageBytes());
            Assert.True(Assert.Single(saved.Lists).IsOrdered);
        }
        image.Text = "Replacement\nSecond";
        Assert.Empty(image.Lists);
        Assert.Equal(new[] { "Replacement", "Second" }, image.Paragraphs.Select(p => p.Text));
        Assert.Equal(resource, image.GetImageBytes());
        Assert.Empty(image.ToXml().Elements(OdfNamespaces.Text + "p"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ImageCropAndMirrorDoNotMoveOrReflectTheCaption(bool mirror) {
        var (document, image, resource) = CreateImage();
        OdfTextParagraph paragraph = image.AddParagraph("Caption");
        paragraph.FontFamily = "Arial"; paragraph.FontSize = OdfLength.Points(10);
        var original = Assert.Single(Texts(document.Pages[0].ToDrawing().Value));
        image.Crop = new OdfInsets(OdfLength.Points(0), OdfLength.Points(10), OdfLength.Points(0), OdfLength.Points(0));
        image.Mirror = mirror ? OdfImageMirror.Horizontal : OdfImageMirror.None;
        foreach (OdgDocument read in RoundTrips(document)) {
            var projection = read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
            var caption = Assert.Single(Texts(projection.Value));
            Assert.Equal((original.X, original.Y, original.Width, original.Height), (caption.X, caption.Y, caption.Width, caption.Height));
            Assert.False(caption.FlipHorizontal); Assert.False(caption.FlipVertical);
            Assert.Equal("Caption", caption.PlainText);
            Assert.Contains("Caption", OfficeDrawingSvgExporter.ToSvg(projection.Value));
            Assert.Equal(resource, read.Pages[0].Shapes[0].GetImageBytes());
        }
    }

    [Fact]
    public void MissingImageBytesStillProjectTheCaptionAndRejectTheImageOmission() {
        var (document, image, _) = CreateImage();
        var paragraph = image.AddParagraph("Available caption");
        paragraph.FontFamily = "Arial"; paragraph.FontSize = OdfLength.Points(10);
        image.Element.Element(OdfNamespaces.Draw + "image")!.SetAttributeValue(OdfNamespaces.XLink + "href", "Pictures/missing.png");
        string before = document.GetXml("content.xml").ToString();
        var projection = document.Pages[0].ToDrawing();
        Assert.Equal("Available caption", Assert.Single(Texts(projection.Value)).PlainText);
        Assert.Contains(projection.Report.Mappings, mapping => mapping.Feature == "shape:Image" && mapping.Status == OdfConversionMappingStatus.Skipped);
        Assert.Throws<OdfConversionLossException>(() => document.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        Assert.Equal(before, document.GetXml("content.xml").ToString());
    }

    [Fact]
    public void EmbeddedObjectAndAnnotationTextAreNotImageCaptions() {
        var (document, image, resource) = CreateImage();
        image.Element.AddFirst(new XElement(OdfNamespaces.Draw + "object", new XElement(OdfNamespaces.Text + "p", "Object story")));
        image.Element.Element(OdfNamespaces.Draw + "image")!.Add(new XElement(OdfNamespaces.Text + "p",
            new XElement(OdfNamespaces.Office + "annotation", new XElement(OdfNamespaces.Text + "p", "Annotation story"))));
        Assert.Equal(string.Empty, image.Text);
        Assert.Empty(Texts(document.Pages[0].ToDrawing().Value));
        Assert.Equal(resource, image.GetImageBytes());
    }

    [Fact]
    public void ImageFrameTextBoxKeepsItsSelectedStoryAndReportsTheOtherCaption() {
        var (document, image, resource) = CreateImage();
        image.AddParagraph("Image story");
        string captionXml = image.ToXml().Element(OdfNamespaces.Draw + "image")!.ToString();
        image.Element.AddFirst(new XElement(OdfNamespaces.Draw + "text-box", new XElement(OdfNamespaces.Text + "p", "Text box story")));
        Assert.Equal("Text box story", image.Text);
        image.Text = "Selected replacement";
        Assert.Equal(captionXml, image.ToXml().Element(OdfNamespaces.Draw + "image")!.ToString());
        foreach (OdgDocument read in RoundTrips(document)) {
            var projection = read.Pages[0].ToDrawing();
            Assert.Equal("Selected replacement", Assert.Single(Texts(projection.Value)).PlainText);
            Assert.Contains(projection.Report.Mappings, mapping => mapping.Feature == "shape:Image:image-text" && mapping.Status == OdfConversionMappingStatus.Skipped);
            Assert.Throws<OdfConversionLossException>(() => read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
            Assert.Equal(resource, read.Pages[0].Shapes[0].GetImageBytes());
        }
    }

    [Fact]
    public void AlternativeImageCaptionRemainsSeparateFromTheSelectedStory() {
        var (document, image, resource) = CreateImage();
        image.AddParagraph("Selected caption");
        var alternative = new XElement(image.Element.Element(OdfNamespaces.Draw + "image")!);
        alternative.Element(OdfNamespaces.Text + "p")!.Value = "Alternative caption";
        image.Element.Add(alternative);
        image.Text = "Edited caption";
        foreach (OdgDocument read in RoundTrips(document)) {
            var projection = read.Pages[0].ToDrawing();
            Assert.Equal("Edited caption", Assert.Single(Texts(projection.Value)).PlainText);
            Assert.Contains(projection.Report.Mappings, mapping => mapping.Feature == "shape:Image:image-text" && mapping.Status == OdfConversionMappingStatus.Skipped);
            Assert.Equal("Alternative caption", read.Pages[0].Shapes[0].ToXml().Elements(OdfNamespaces.Draw + "image").Last().Element(OdfNamespaces.Text + "p")!.Value);
            Assert.Throws<OdfConversionLossException>(() => read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
            Assert.Equal(resource, read.Pages[0].Shapes[0].GetImageBytes());
        }
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void AlternativeTextBoxStoryIsPreservedAndReported(bool hasImage, bool selectedStoryIsEmpty) {
        var (document, image, resource) = CreateImage();
        OdgShape frame = hasImage ? image : document.Pages[0].Shapes.AddTextBox(
            OdfRect.FromCentimeters(8, 1, 6, 3), string.Empty, "TextFrame");
        if (hasImage)
            frame.Element.Add(new XElement(OdfNamespaces.Draw + "text-box"));
        if (!selectedStoryIsEmpty) frame.AddParagraph("Selected story");
        var alternative = new XElement(OdfNamespaces.Draw + "text-box",
            new XElement(OdfNamespaces.Text + "p", "Alternative story"));
        frame.Element.Add(alternative);
        var preservedAlternative = new XElement(alternative);
        if (!selectedStoryIsEmpty) frame.Text = "Edited selected story";
        string expectedText = selectedStoryIsEmpty ? string.Empty : "Edited selected story";

        foreach (OdgDocument read in RoundTrips(document)) {
            OdgShape saved = read.Pages[0].Shapes.First(shape => shape.Name == frame.Name);
            string before = read.GetXml("content.xml").ToString();
            var projection = read.Pages[0].ToDrawing();
            Assert.Equal(expectedText, saved.Text);
            Assert.True(XNode.DeepEquals(preservedAlternative, saved.ToXml().Elements(OdfNamespaces.Draw + "text-box").Last()));
            Assert.DoesNotContain(Texts(projection.Value), text => text.PlainText == "Alternative story");
            if (!selectedStoryIsEmpty)
                Assert.Contains(Texts(projection.Value), text => text.PlainText == expectedText);
            Assert.Contains(projection.Report.Mappings, mapping =>
                mapping.Feature == "shape:" + frame.Name + (hasImage ? ":image-text" : ":frame-text") &&
                mapping.Status == OdfConversionMappingStatus.Skipped);
            Assert.Throws<OdfConversionLossException>(() => read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
            Assert.Equal(before, read.GetXml("content.xml").ToString());
            Assert.Equal(resource, read.Pages[0].Shapes.First(shape => shape.Name == image.Name).GetImageBytes());
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FrameTableStoriesRemainNativeAndRejectSilentOmission(bool selectedStory) {
        var document = OdgDocument.Create();
        var frame = document.AddPage().Shapes.AddTextBox(OdfRect.FromCentimeters(1, 1, 6, 3), string.Empty, "TableFrame");
        XElement story = frame.Element.Element(OdfNamespaces.Draw + "text-box")!;
        if (!selectedStory) {
            story = new XElement(OdfNamespaces.Draw + "text-box");
            frame.Element.Add(story);
        }
        var table = new XElement(OdfNamespaces.Table + "table",
            new XAttribute(OdfNamespaces.Table + "name", "StoryTable"),
            new XElement(OdfNamespaces.Table + "table-row",
                new XElement(OdfNamespaces.Table + "table-cell", new XElement(OdfNamespaces.Text + "p", "Native table story"))));
        story.Add(table);
        foreach (OdgDocument read in RoundTrips(document)) {
            string before = read.GetXml("content.xml").ToString();
            var projection = read.Pages[0].ToDrawing();
            Assert.Empty(Texts(projection.Value));
            Assert.Contains(projection.Report.Mappings, mapping =>
                mapping.Feature == "shape:TableFrame" + (selectedStory ? ":text:unmapped-text-container" : ":frame-text") &&
                mapping.Status == (selectedStory ? OdfConversionMappingStatus.Unsupported : OdfConversionMappingStatus.Skipped));
            Assert.Throws<OdfConversionLossException>(() => read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
            Assert.True(XNode.DeepEquals(new XElement(table), read.Pages[0].Shapes[0].ToXml().Descendants(OdfNamespaces.Table + "table").Single()));
            Assert.Equal(before, read.GetXml("content.xml").ToString());
        }
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void FrameStructuredStoriesRemainNativeAndRejectSilentOmission(bool selectedStory, bool numberedParagraph) {
        var document = OdgDocument.Create();
        var frame = document.AddPage().Shapes.AddTextBox(OdfRect.FromCentimeters(1, 1, 6, 3), string.Empty, "StructuredFrame");
        XElement story = frame.Element.Element(OdfNamespaces.Draw + "text-box")!;
        if (!selectedStory) {
            story = new XElement(OdfNamespaces.Draw + "text-box");
            frame.Element.Add(story);
        }
        var container = new XElement(OdfNamespaces.Text + (numberedParagraph ? "numbered-paragraph" : "section"),
            new XAttribute(OdfNamespaces.Text + (numberedParagraph ? "list-id" : "name"), "NativeStory"),
            new XElement(OdfNamespaces.Text + "p", "Native structured story"));
        story.Add(container);
        foreach (OdgDocument read in RoundTrips(document)) {
            string before = read.GetXml("content.xml").ToString();
            var projection = read.Pages[0].ToDrawing();
            Assert.Empty(Texts(projection.Value));
            Assert.Contains(projection.Report.Mappings, mapping =>
                mapping.Feature == "shape:StructuredFrame" + (selectedStory ? ":text:unmapped-text-container" : ":frame-text") &&
                mapping.Status == (selectedStory ? OdfConversionMappingStatus.Unsupported : OdfConversionMappingStatus.Skipped));
            Assert.Throws<OdfConversionLossException>(() => read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
            Assert.True(XNode.DeepEquals(container, read.Pages[0].Shapes[0].ToXml().Descendants(container.Name).Single()));
            Assert.Equal(before, read.GetXml("content.xml").ToString());
        }
    }

    private static (OdgDocument Document, OdgShape Image, byte[] Resource) CreateImage() {
        var document = OdgDocument.Create();
        var raster = new OfficeRasterImage(40, 20);
        new OfficeRasterCanvas(raster).FillRectangle(0, 0, 40, 20, OfficeColor.FromRgb(80, 160, 220));
        byte[] resource = OfficePngWriter.Encode(raster, new OfficePngEncodeOptions { DpiX = 96, DpiY = 96 });
        var image = document.AddPage().Shapes.AddImage(resource, "caption.png", OdfRect.FromCentimeters(1, 1, 6, 3), "Image");
        return (document, image, resource);
    }

    private static IEnumerable<OfficeDrawingRichText> Texts(OfficeDrawing drawing) {
        foreach (var element in drawing.Elements) {
            if (element is OfficeDrawingRichText text) yield return text;
            else if (element is OfficeDrawingEffectGroup group)
                foreach (var nested in Texts(group.Drawing)) yield return nested;
        }
    }

    private static IEnumerable<OdgDocument> RoundTrips(OdgDocument document) {
        yield return document;
        yield return OdgDocument.Load(new MemoryStream(document.ToBytes()));
        using var flat = new MemoryStream(); document.SaveFlatXml(flat); flat.Position = 0;
        yield return OdgDocument.LoadFlatXml(flat);
    }
}
