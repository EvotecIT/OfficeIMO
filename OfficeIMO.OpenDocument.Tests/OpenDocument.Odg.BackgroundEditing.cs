using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed partial class OpenDocumentOdgBackgroundTests {
    [Fact]
    public void PageBackgroundEditsCopySharedStylesAndNoFillRestoresSharedMasterPaint() {
        var document = Drawing(); var page = document.Pages[0];
        page.MasterBackground.FillColor = OdfColor.Parse("#e07020"); page.MasterBackgroundSize = OdgBackgroundSize.Full;
        page.Background.FillColor = OdfColor.Parse("#2040e0"); page.Background.FillOpacity = .5;
        var copy = document.ClonePage(0, "Copy");
        copy.Background.FillColor = OdfColor.Parse("#00ff00"); copy.Background.FillOpacity = 1;
        foreach (var read in RoundTrips(document)) {
            Assert.Equal(OdfColor.Parse("#2040e0"), read.Pages[0].Background.FillColor); Assert.Equal(.5, read.Pages[0].Background.FillOpacity);
            Assert.Equal(OdfColor.Parse("#00ff00"), read.Pages[1].Background.FillColor); Assert.Equal(1, read.Pages[1].Background.FillOpacity);
            read.Pages[1].Background.UseNoFill(); Assert.Equal("none", read.Pages[1].Background.FillMode);
            var paint = Assert.Single(read.Pages[1].ToDrawing().Value.Elements.OfType<OfficeDrawingShape>());
            Assert.Equal(OfficeColor.Parse("#e07020"), paint.Shape.FillColor);
        }
    }

    [Fact]
    public void MasterEditsAreSharedUntilTheMasterIsClonedIncludingItsPaintArea() {
        var document = Drawing(); var page = document.Pages[0];
        page.MasterBackground.FillColor = OdfColor.Parse("#e07020"); var copy = document.ClonePage(0, "Copy");
        copy.MasterBackground.FillColor = OdfColor.Parse("#2040e0"); copy.MasterBackgroundSize = OdgBackgroundSize.Full;
        Assert.Equal(copy.MasterBackground.FillColor, page.MasterBackground.FillColor); Assert.Equal(OdgBackgroundSize.Full, page.MasterBackgroundSize);
        copy.MasterPageName = document.CloneMasterPage(page.MasterPageName, "Independent");
        copy.MasterBackground.FillColor = OdfColor.Parse("#00ff00"); copy.MasterBackgroundSize = OdgBackgroundSize.Border;
        foreach (var read in RoundTrips(document)) {
            Assert.Equal(OdfColor.Parse("#2040e0"), read.Pages[0].MasterBackground.FillColor); Assert.Equal(OdgBackgroundSize.Full, read.Pages[0].MasterBackgroundSize);
            Assert.Equal(OdfColor.Parse("#00ff00"), read.Pages[1].MasterBackground.FillColor); Assert.Equal(OdgBackgroundSize.Border, read.Pages[1].MasterBackgroundSize);
            read.Pages[1].MasterBackground.FillColor = null; Assert.Empty(read.Pages[1].ToDrawing().Value.Elements);
        }
    }

    [Fact]
    public void BackgroundEditsRetainNamedInheritanceAndUnrelatedAutomaticStyleXml() {
        var document = Drawing(); var page = document.Pages[0];
        var parent = document.Styles.CreateNamed("PaintParent", OdfStyleFamily.DrawingPage);
        parent.SetProperty(OdfNamespaces.Style + "drawing-page-properties", OdfNamespaces.Draw + "opacity", "40%");
        var properties = Bind(document, page.Element, "content.xml", "Paint", "solid", "#2040e0");
        XNamespace custom = "urn:officeimo:test:paint";
        properties.SetAttributeValue(custom + "metadata", "retained");
        properties.Parent!.SetAttributeValue(OdfNamespaces.Style + "parent-style-name", parent.Name);
        properties.Parent.Add(new XElement(custom + "opaque", "payload")); document.MarkPartDirty("content.xml");
        var copy = document.ClonePage(0); copy.Background.FillOpacity = .8; copy.Background.FillColor = OdfColor.Parse("#00ff00");
        foreach (var read in RoundTrips(document)) {
            var selected = read.Pages[1].Element;
            var style = read.Styles.FindInPart(OdfStyleFamily.DrawingPage, (string)selected.Attribute(OdfNamespaces.Draw + "style-name")!, "content.xml")!;
            Assert.Equal("PaintParent", style.ParentStyleName); Assert.Equal("payload", style.Element.Element(custom + "opaque")!.Value);
            Assert.Equal("retained", (string?)style.Element.Element(OdfNamespaces.Style + "drawing-page-properties")!.Attribute(custom + "metadata"));
            Assert.Equal(.4, read.Pages[0].Background.FillOpacity); Assert.Equal(.8, read.Pages[1].Background.FillOpacity);
            read.Pages[1].Background.FillOpacity = null; Assert.Equal(.4, read.Pages[1].Background.FillOpacity);
        }
    }

    [Fact]
    public void NamedBackgroundStyleEditsCreateLocalOverridesWithoutChangingTheParent() {
        var document = Drawing(); var page = document.Pages[0];
        var style = document.Styles.CreateNamed("CommonPaint", OdfStyleFamily.DrawingPage);
        style.SetProperty(OdfNamespaces.Style + "drawing-page-properties", OdfNamespaces.Draw + "fill", "solid");
        style.SetProperty(OdfNamespaces.Style + "drawing-page-properties", OdfNamespaces.Draw + "fill-color", "#e07020");
        page.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", style.Name); document.MarkPartDirty("content.xml");
        var copy = document.ClonePage(0); copy.Background.FillColor = OdfColor.Parse("#2040e0");
        foreach (var read in RoundTrips(document)) {
            Assert.Equal(OdfColor.Parse("#e07020"), read.Pages[0].Background.FillColor); Assert.Equal(OdfColor.Parse("#2040e0"), read.Pages[1].Background.FillColor);
            Assert.Equal("#e07020", (string?)read.Styles.Find(OdfStyleFamily.DrawingPage, "CommonPaint")!.Element
                .Element(OdfNamespaces.Style + "drawing-page-properties")!.Attribute(OdfNamespaces.Draw + "fill-color"));
        }
    }

    [Fact]
    public void GradientBindingsEditIndependentlyWhileDefinitionsRemainShared() {
        var document = Drawing(); var page = document.Pages[0];
        var pattern = new OdfGradientPattern(OdfGradientStyle.Linear, OdfColor.Parse("#e07020"), OdfColor.Parse("#2040e0"), 90);
        var gradient = document.Styles.CreateGradient("Paint", pattern); document.Styles.CreateGradient("Other", pattern);
        page.Background.FillGradientName = gradient.Name; page.Background.GradientStepCount = 0; var copy = document.ClonePage(0);
        copy.Background.FillGradientName = "Other"; copy.Background.GradientStepCount = 12;
        gradient.Pattern = new OdfGradientPattern(OdfGradientStyle.Radial, pattern.StartColor, pattern.EndColor, centerX: .5, centerY: .5);
        foreach (var read in RoundTrips(document)) {
            Assert.Equal("Paint", read.Pages[0].Background.FillGradientName); Assert.Equal(0, read.Pages[0].Background.GradientStepCount);
            Assert.Equal(OdfGradientStyle.Radial, read.Styles.FindGradient("Paint")!.Pattern.Style);
            Assert.NotNull(Assert.Single(read.Pages[0].ToDrawing().Value.Elements.OfType<OfficeDrawingShape>()).Shape.FillRadialGradient);
            Assert.Equal("Other", read.Pages[1].Background.FillGradientName); Assert.Equal(12, read.Pages[1].Background.GradientStepCount);
            Assert.Throws<OdfConversionLossException>(() => read.Pages[1].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
            read.Pages[1].Background.FillGradientName = null; Assert.Equal("none", read.Pages[1].Background.FillMode);
        }
    }

    [Fact]
    public void BitmapReplacementIsIndependentDeduplicatedAndRetainsOpacityAcrossContainers() {
        var document = Drawing(); var page = document.Pages[0];
        byte[] replacement = OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.CornflowerBlue));
        page.Background.SetBitmap(Png); page.Background.FillOpacity = .5;
        var copy = document.ClonePage(0); copy.Background.SetBitmap(Png); Assert.Single(document.PackageEntries, path => path.StartsWith("Pictures/", StringComparison.Ordinal));
        Assert.NotEqual(page.Background.FillImageName, copy.Background.FillImageName); copy.Background.SetBitmap(replacement);
        Assert.Equal(2, document.PackageEntries.Count(path => path.StartsWith("Pictures/", StringComparison.Ordinal)));
        foreach (var read in RoundTrips(document)) {
            Assert.Equal(Png, read.Pages[0].Background.GetBitmapBytes()); Assert.Equal(replacement, read.Pages[1].Background.GetBitmapBytes());
            Assert.Equal(.5, read.Pages[1].Background.FillOpacity);
            Assert.Equal(replacement, Assert.Single(read.Pages[1].ToDrawing().Value.Elements.OfType<OfficeDrawingImage>()).Bytes);
            read.Pages[1].Background.UseNoFill(); Assert.Throws<InvalidOperationException>(() => read.Pages[1].Background.GetBitmapBytes());
        }
    }

    [Fact]
    public void AuthoredBackgroundResourcesAndMasterPaintImportWithoutChangingTheirSource() {
        var source = Drawing(); var original = source.Pages[0]; original.Background.SetBitmap(Png); original.MasterBackground.FillColor = OdfColor.Parse("#e07020");
        original.MasterBackgroundSize = OdgBackgroundSize.Full; string[] before = Parts(source);
        var destination = Drawing(); destination.Pages[0].Background.SetBitmap(Png); var imported = destination.ImportPage(source, 0);
        Assert.Equal(before, Parts(source)); imported.Background.UseNoFill();
        foreach (var read in RoundTrips(destination)) {
            Assert.Equal(Png, read.Pages[0].Background.GetBitmapBytes()); Assert.Equal(OdfColor.Parse("#e07020"), read.Pages[1].MasterBackground.FillColor);
            Assert.Equal(OdgBackgroundSize.Full, read.Pages[1].MasterBackgroundSize);
            Assert.Equal(OfficeColor.Parse("#e07020"), Assert.Single(read.Pages[1].ToDrawing().Value.Elements.OfType<OfficeDrawingShape>()).Shape.FillColor);
        }
    }

    [Theory]
    [InlineData("gradient")]
    [InlineData("gradientName")]
    [InlineData("opacity")]
    [InlineData("bands")]
    [InlineData("area")]
    [InlineData("extension")]
    [InlineData("payload")]
    [InlineData("binding")]
    public void InvalidBackgroundEditsDoNotChangeXmlOrResourceEntries(string kind) {
        var document = Drawing(); var page = document.Pages[0];
        if (kind == "binding") { page.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", "Missing"); document.MarkPartDirty("content.xml"); }
        string[] before = Parts(document), entries = document.PackageEntries.ToArray();
        Assert.ThrowsAny<Exception>(() => {
            switch (kind) {
                case "gradient": page.Background.FillGradientName = "Missing"; break;
                case "gradientName": page.Background.FillGradientName = "Invalid Name"; break;
                case "opacity": page.Background.FillOpacity = double.NaN; break;
                case "bands": page.Background.GradientStepCount = 2; break;
                case "area": page.MasterBackgroundSize = (OdgBackgroundSize)99; break;
                case "extension": page.Background.SetBitmap(Png, "background.jpg"); break;
                case "payload": page.Background.SetBitmap(new byte[] { 1, 2, 3 }); break;
                case "binding": page.Background.SetBitmap(Png); break;
            }
        });
        Assert.Equal(before, Parts(document)); Assert.Equal(entries, document.PackageEntries);
    }

    [Fact]
    public void ActiveOpacityGradientsAndLinkedBitmapsArePreservedWithoutReplacementOrFetch() {
        var document = Drawing(); var properties = Bitmap(document); var page = document.Pages[0];
        properties.SetAttributeValue(OdfNamespaces.Draw + "opacity-name", "Opacity");
        document.GetXml("styles.xml").Descendants(OdfNamespaces.Draw + "fill-image").Single().SetAttributeValue(OdfNamespaces.XLink + "href", "https://example.invalid/image.png");
        document.MarkPartDirty("styles.xml"); string[] before = Parts(document);
        Assert.Throws<NotSupportedException>(() => page.MasterBackground.FillOpacity = .5);
        Assert.Throws<NotSupportedException>(() => page.MasterBackground.GetBitmapBytes()); Assert.Equal(before, Parts(document));
    }

    [Fact]
    public void BackgroundEditsCreateOmittedStyleContainersBeforeBodyAndMasterStyles() {
        var document = Drawing();
        document.GetXml("content.xml").Root!.Element(OdfNamespaces.Office + "automatic-styles")!.Remove();
        document.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!.Remove();
        document.MarkPartDirty("content.xml"); document.MarkPartDirty("styles.xml");
        // Saving/loading exercises a valid producer input with optional containers omitted.
        document = OdgDocument.Load(new MemoryStream(document.ToBytes())); document.Pages[0].Background.SetBitmap(Png);
        foreach (var read in RoundTrips(document)) {
            Assert.Equal(Png, read.Pages[0].Background.GetBitmapBytes());
            XName[] content = read.GetXml("content.xml").Root!.Elements().Select(element => element.Name).ToArray();
            XName[] styles = read.GetXml("styles.xml").Root!.Elements().Select(element => element.Name).ToArray();
            Assert.True(Array.IndexOf(content, OdfNamespaces.Office + "automatic-styles") < Array.IndexOf(content, OdfNamespaces.Office + "body"));
            Assert.True(Array.IndexOf(styles, OdfNamespaces.Office + "styles") < Array.IndexOf(styles, OdfNamespaces.Office + "automatic-styles"));
            Assert.True(Array.IndexOf(styles, OdfNamespaces.Office + "styles") < Array.IndexOf(styles, OdfNamespaces.Office + "master-styles"));
        }
    }
}
