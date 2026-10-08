using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed partial class OpenDocumentOdgImportTests {
    [Fact]
    public void ImportsConnectedImagePageWithIndependentResourcesStylesLayoutAndExistingWrappers() {
        var source = OdgDocument.Create(); var page = source.AddPage("Source", OdfLength.Points(300), OdfLength.Points(200));
        var a = page.Shapes.AddRectangle(Rect(10, 10, 40, 30), "Start"); a.FillColor = OdfColor.Parse("123456"); a.Text = "Original";
        var b = page.Shapes.AddRectangle(Rect(150, 10, 40, 30), "End");
        page.Shapes.AddConnector(a.AddGluePoint(OdgGluePointAlignment.Right), b.AddGluePoint(OdgGluePointAlignment.Left), "Link");
        byte[] pixel = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        page.Shapes.AddImage(pixel, "pixel.png", Rect(10, 70, 25, 25)); page.Shapes.AddImage(pixel, "pixel.png", Rect(50, 70, 25, 25));
        source.Layers.Add("Hidden", OdgLayerDisplay.None); source.Layers.Find("Hidden")!.IsProtected = true; b.Layer = "Hidden";
        var destination = OdgDocument.Create(); var existing = destination.AddPage("Source");
        var existingShape = existing.Shapes.AddRectangle(Rect(0, 0, 50, 20)); existingShape.FillColor = OdfColor.Parse("abcdef");
        destination.Package.AddOrReplaceEntry("Pictures/odgImport1.png", new byte[] { 1, 2, 3 }, "image/png");
        string[] sourceParts = Parts(source);
        string before = Svg(page);
        var imported = destination.ImportPage(source, 0);
        Assert.Equal("SourceCopy1", imported.Name); Assert.Equal(sourceParts, Parts(source));
        Assert.Equal(before, Svg(imported)); Assert.NotEqual(page.MasterPageName, imported.MasterPageName);
        Assert.Equal(new byte[] { 1, 2, 3 }, destination.Package.GetRequiredEntry("Pictures/odgImport1.png").GetBytesForSave());
        Assert.Equal(2, destination.Package.Entries.Count(entry => entry.Name.StartsWith("Pictures/", StringComparison.Ordinal)));
        Assert.Equal(OdgLayerDisplay.None, imported.EffectiveLayers.Find("Hidden")!.Display); Assert.True(imported.EffectiveLayers.Find("Hidden")!.IsProtected);
        var link = imported.Shapes[2]; Assert.Equal(imported.Shapes[0].XmlId, link.StartShapeId); Assert.Equal(imported.Shapes[1].XmlId, link.EndShapeId);
        imported.Width = OdfLength.Points(420); imported.Shapes[0].FillColor = OdfColor.Parse("ff0000"); imported.Shapes[0].Text = "Updated";
        existingShape.Text = "Still attached";
        foreach (var read in RoundTrips(destination)) {
            Assert.Equal("Still attached", read.Pages[0].Shapes[0].Text); Assert.Equal("Updated", read.Pages[1].Shapes[0].Text);
            Assert.Equal(420, read.Pages[1].Width.ToPoints(), 3); Assert.Equal(OdfColor.Parse("ff0000"), read.Pages[1].Shapes[0].FillColor);
            Assert.Equal(OdfColor.Parse("abcdef"), read.Pages[0].Shapes[0].FillColor);
        }
        Assert.Equal(300, page.Width.ToPoints(), 3); Assert.Equal("Original", a.Text); Assert.Equal(sourceParts, Parts(source));
    }

    [Fact]
    public void RetainsPartScopedAutomaticStylesNamedParentsAndSourceDefaultsWithoutChangingDestinationDefaults() {
        var source = OdgDocument.Create(); var page = source.AddPage();
        var shape = page.Shapes.AddRectangle(Rect(0, 0, 40, 30));
        SetFill(source.Styles.CreateNamed("Collision", OdfStyleFamily.Graphic), "#0000ff");
        var contentAutomatic = new XElement(OdfNamespaces.Style + "style", new XAttribute(OdfNamespaces.Style + "name", "Collision"),
            new XAttribute(OdfNamespaces.Style + "family", "graphic"), new XAttribute(OdfNamespaces.Style + "parent-style-name", "Parent"),
            new XElement(OdfNamespaces.Style + "graphic-properties", new XAttribute(OdfNamespaces.Draw + "fill", "solid"), new XAttribute(OdfNamespaces.Draw + "fill-color", "#ff0000")));
        source.Styles.CreateNamed("Parent", OdfStyleFamily.Graphic).SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Svg + "stroke-width", "2pt");
        source.GetXml("content.xml").Root!.Element(OdfNamespaces.Office + "automatic-styles")!.Add(contentAutomatic);
        shape.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", "Collision");
        source.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "automatic-styles")!.Add(new XElement(contentAutomatic));
        var masterArt = new XElement(shape.Element); masterArt.SetAttributeValue(OdfNamespaces.Draw + "name", "Art");
        page.Master!.Add(masterArt);
        Default(source, "graphic", new XAttribute(OdfNamespaces.Svg + "stroke-color", "#123456"));
        source.Styles.FindDefault(OdfStyleFamily.Graphic)!.Element.Add(new XElement(OdfNamespaces.Style + "text-properties", new XAttribute(OdfNamespaces.Style + "font-name", "Face")));
        foreach (string part in new[] { "content.xml", "styles.xml" }) source.GetXml(part).Root!.Element(OdfNamespaces.Office + "font-face-decls")!.Add(
            new XElement(OdfNamespaces.Style + "font-face", new XAttribute(OdfNamespaces.Style + "name", "Face"),
                new XAttribute(OdfNamespaces.Svg + "font-family", part == "content.xml" ? "'DejaVu Sans'" : "Arial")));
        shape.Text = "Scoped font";
        var destination = OdgDocument.Create(); var existing = destination.AddPage();
        Default(destination, "graphic", new XAttribute(OdfNamespaces.Svg + "stroke-color", "#abcdef"));
        var old = existing.Shapes.AddRectangle(Rect(0, 0, 10, 10)); old.Element.Attribute(OdfNamespaces.Draw + "style-name")?.Remove();
        var named = destination.Styles.CreateNamed("Collision", OdfStyleFamily.Graphic); SetFill(named, "#ffff00");
        source.MarkPartDirty("content.xml"); source.MarkPartDirty("styles.xml"); destination.MarkPartDirty("content.xml");
        string[] beforeSource = Parts(source);
        var imported = destination.ImportPage(source, 0);
        Assert.Equal(beforeSource, Parts(source)); Assert.Equal(Svg(page), Svg(imported));
        Assert.Equal("#ffff00", (string?)named.Element.Element(OdfNamespaces.Style + "graphic-properties")!.Attribute(OdfNamespaces.Draw + "fill-color"));
        string pageStyle = (string)imported.Shapes[0].Element.Attribute(OdfNamespaces.Draw + "style-name")!;
        string masterStyle = (string)imported.Master!.Element(OdfNamespaces.Draw + "rect")!.Attribute(OdfNamespaces.Draw + "style-name")!;
        Assert.NotEqual(pageStyle, masterStyle);
        Assert.Equal("#123456", (string?)destination.Styles.Find(OdfStyleFamily.Graphic, pageStyle)!.Element.Element(OdfNamespaces.Style + "graphic-properties")!.Attribute(OdfNamespaces.Svg + "stroke-color")
            ?? (string?)destination.Styles.Resolve(destination.Styles.Find(OdfStyleFamily.Graphic, pageStyle)!).Last().Element.Element(OdfNamespaces.Style + "graphic-properties")!.Attribute(OdfNamespaces.Svg + "stroke-color"));
        Assert.Equal("#abcdef", (string?)destination.Styles.FindDefault(OdfStyleFamily.Graphic)!.Element.Element(OdfNamespaces.Style + "graphic-properties")!.Attribute(OdfNamespaces.Svg + "stroke-color"));
    }

    [Fact]
    public void ImportsReachablePaintListAndFontDefinitionsWithoutUnusedStyles() {
        var source = OdgDocument.Create(); var page = source.AddPage();
        source.Styles.CreateGradient("Paint", new OdfGradientPattern(OdfGradientStyle.Linear, OdfColor.Parse("ff0000"), OdfColor.Parse("0000ff"), 45));
        source.Styles.CreateMarker("Arrow", new OdfMarkerGeometry(new OdfViewBox(0, 0, 20, 30), "M10 0L0 30H20Z"));
        SetFill(source.Styles.CreateNamed("Unused", OdfStyleFamily.Graphic), "#bbbbbb");
        var shape = page.Shapes.AddRectangle(Rect(0, 0, 40, 30)); shape.FillGradientName = "Paint";
        var line = page.Shapes.AddLine(OdfLength.Points(0), OdfLength.Points(80), OdfLength.Points(80), OdfLength.Points(80)); line.StrokeEndMarkerName = "Arrow";
        var text = page.Shapes.AddTextBox(Rect(0, 100, 100, 30), "");
        string listName = OdfListStyleStore.Create(source, true);
        text.Element.Element(OdfNamespaces.Draw + "text-box")!.RemoveNodes();
        text.Element.Element(OdfNamespaces.Draw + "text-box")!.Add(new XElement(OdfNamespaces.Text + "list", new XAttribute(OdfNamespaces.Text + "style-name", listName),
            new XElement(OdfNamespaces.Text + "list-item", new XElement(OdfNamespaces.Text + "p", new XElement(OdfNamespaces.Text + "span", new XAttribute(OdfNamespaces.Text + "style-name", "Run"), "Item")))));
        var run = source.Styles.CreateNamed("Run", OdfStyleFamily.Text); run.FontSize = OdfLength.Points(14);
        run.Element.Element(OdfNamespaces.Style + "text-properties")!.SetAttributeValue(OdfNamespaces.Style + "font-name", "Face");
        source.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "font-face-decls")!.Add(new XElement(OdfNamespaces.Style + "font-face",
            new XAttribute(OdfNamespaces.Style + "name", "Face"), new XAttribute(OdfNamespaces.Svg + "font-family", "'DejaVu Sans'")));
        source.MarkPartDirty("content.xml"); source.MarkPartDirty("styles.xml");
        var destination = OdgDocument.Create(); destination.Styles.CreateGradient("Paint", new OdfGradientPattern(OdfGradientStyle.Radial, OdfColor.Parse("ffffff"), OdfColor.Parse("000000")));
        var imported = destination.ImportPage(source, 0);
        Assert.Equal(Svg(page), Svg(imported));
        Assert.Equal(2, destination.Styles.Gradients.Count); Assert.Single(destination.Styles.Markers);
        Assert.DoesNotContain(destination.Styles.Named, style => (string?)style.Element.Element(OdfNamespaces.Style + "graphic-properties")?.Attribute(OdfNamespaces.Draw + "fill-color") == "#bbbbbb");
        foreach (var read in RoundTrips(destination)) Assert.Equal(Svg(page), Svg(read.Pages[0]));
    }

    [Theory]
    [InlineData("libreoffice-routed-glue.odg")]
    [InlineData("libreoffice-sheared-geometry.odg")]
    [InlineData("libreoffice-transparent-text.fodg")]
    public void ImportsIndependentProducerPageWithoutChangingProjection(string fixture) {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", fixture);
        var source = fixture.EndsWith(".fodg", StringComparison.Ordinal) ? OdgDocument.LoadFlatXml(path) : OdgDocument.Load(path);
        string[] sourceParts = Parts(source);
        var destination = OdgDocument.Create(); var imported = destination.ImportPage(source, 0);
        Assert.Equal(sourceParts, Parts(source)); Assert.Equal(Svg(source.Pages[0]), Svg(imported));
        var reopened = OdgDocument.Load(new MemoryStream(destination.ToBytes())); Assert.Equal(Svg(source.Pages[0]), Svg(reopened.Pages[0]));
    }

    [Theory]
    [InlineData("missingStyle")]
    [InlineData("duplicateStyle")]
    [InlineData("parentCycle")]
    [InlineData("missingImage")]
    [InlineData("externalFragment")]
    [InlineData("externalNote")]
    [InlineData("relativeLink")]
    [InlineData("destinationDefault")]
    [InlineData("otherMaster")]
    [InlineData("legacyAngle")]
    [InlineData("destinationVersion")]
    [InlineData("styleDepth")]
    [InlineData("classes")]
    [InlineData("conditionalText")]
    public void UnsupportedDependenciesFailBeforeEitherDocumentChanges(string failure) {
        var source = OdgDocument.Create(); var page = source.AddPage();
        var shape = page.Shapes.AddTextBox(Rect(0, 0, 20, 20), "");
        var destination = OdgDocument.Create(); destination.AddPage();
        switch (failure) {
            case "missingStyle": shape.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", "Missing"); break;
            case "duplicateStyle":
                var style = source.Styles.CreateNamed("Duplicate", OdfStyleFamily.Graphic); shape.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", "Duplicate");
                style.Element.AddAfterSelf(new XElement(style.Element)); break;
            case "parentCycle": source.Styles.CreateNamed("A", OdfStyleFamily.Graphic, "B"); source.Styles.CreateNamed("B", OdfStyleFamily.Graphic, "A"); shape.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", "A"); break;
            case "missingImage": shape.Element.Add(new XElement(OdfNamespaces.Draw + "image", new XAttribute(OdfNamespaces.XLink + "href", "Pictures/missing.png"))); break;
            case "externalFragment": shape.Element.Add(new XElement(OdfNamespaces.Text + "a", new XAttribute(OdfNamespaces.XLink + "href", "#Elsewhere"))); break;
            case "externalNote": shape.Element.Add(new XElement(OdfNamespaces.Text + "note-ref", new XAttribute(OdfNamespaces.Text + "ref-name", "external"))); break;
            case "relativeLink": shape.Element.Add(new XElement(OdfNamespaces.Text + "a", new XAttribute(OdfNamespaces.XLink + "href", "../other.odg"))); break;
            case "destinationDefault": Default(destination, "graphic", new XAttribute(OdfNamespaces.Draw + "fill-color", "#ff0000")); break;
            case "otherMaster": page.Master!.SetAttributeValue(OdfNamespaces.Style + "next-style-name", source.AddPage().MasterPageName); break;
            case "legacyAngle":
                source.Styles.CreateGradient("Angle", new OdfGradientPattern(OdfGradientStyle.Linear, OdfColor.Parse("ff0000"), OdfColor.Parse("0000ff")));
                source.GetXml("styles.xml").Descendants(OdfNamespaces.Draw + "gradient").Single().SetAttributeValue(OdfNamespaces.Draw + "angle", "450"); shape.FillGradientName = "Angle";
                using (var flat = new MemoryStream()) { source.SaveFlatXml(flat); string xml = System.Text.Encoding.UTF8.GetString(flat.ToArray()).Replace("office:version=\"1.4\"", "office:version=\"1.2\""); source = OdgDocument.LoadFlatXml(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(xml))); } break;
            case "destinationVersion": destination = OdgDocument.Load(new MemoryStream(destination.ToBytes(new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.Odf13 }))); break;
            case "styleDepth":
                for (int i = 0; i < 270; i++) source.Styles.CreateNamed("Depth" + i, OdfStyleFamily.Graphic, i == 269 ? null : "Depth" + (i + 1));
                shape.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", "Depth0"); break;
            case "classes": shape.Element.SetAttributeValue(OdfNamespaces.Draw + "class-names", "Class"); break;
            case "conditionalText":
                var paragraphStyle = source.Styles.CreateNamed("Paragraph", OdfStyleFamily.Paragraph); source.Styles.CreateNamed("Alternative", OdfStyleFamily.Paragraph);
                paragraphStyle.Element.Add(new XElement(OdfNamespaces.Style + "map", new XAttribute(OdfNamespaces.Style + "condition", "outline-level()=1"), new XAttribute(OdfNamespaces.Style + "apply-style-name", "Alternative")));
                shape.Element.Descendants(OdfNamespaces.Text + "p").Single().SetAttributeValue(OdfNamespaces.Text + "style-name", "Paragraph"); break;
        }
        string[] beforeSource = Parts(source), beforeDestination = Parts(destination);
        string[] paths = destination.Package.Entries.Select(entry => entry.Name).ToArray();
        if (failure is "relativeLink" or "destinationDefault" or "otherMaster" or "legacyAngle" or "destinationVersion" or "styleDepth" or "classes" or "conditionalText") Assert.Throws<NotSupportedException>(() => destination.ImportPage(source, 0));
        else Assert.Throws<InvalidDataException>(() => destination.ImportPage(source, 0));
        Assert.Equal(beforeSource, Parts(source)); Assert.Equal(beforeDestination, Parts(destination)); Assert.Equal(paths, destination.Package.Entries.Select(entry => entry.Name));
    }

    [Fact]
    public void RemapsIdentifiersAcrossMasterAndPageAndRetainsAbsoluteLinks() {
        var source = OdgDocument.Create(); var page = source.AddPage("Page One");
        var shape = page.Shapes.AddTextBox(Rect(0, 0, 20, 20), ""); shape.XmlId = "pageShape";
        var art = new XElement(OdfNamespaces.Draw + "rect", new XAttribute(XNamespace.Xml + "id", "art"), new XAttribute(OdfNamespaces.Draw + "name", "Logo"));
        page.Master!.Add(art);
        shape.Element.Add(new XElement(OdfNamespaces.Text + "a", new XAttribute(OdfNamespaces.XLink + "href", "#art"), "Master"),
            new XElement(OdfNamespaces.Text + "a", new XAttribute(OdfNamespaces.XLink + "href", "#Page%20One"), "Page"),
            new XElement(OdfNamespaces.Text + "a", new XAttribute(OdfNamespaces.XLink + "href", "https://example.org/drawing"), "Web"));
        var destination = OdgDocument.Create(); var imported = destination.ImportPage(source, 0, "Page Two");
        var links = imported.Shapes[0].Element.Elements(OdfNamespaces.Text + "a").ToArray();
        Assert.Equal("#" + (string?)imported.Master!.Element(OdfNamespaces.Draw + "rect")!.Attribute(XNamespace.Xml + "id"), (string?)links[0].Attribute(OdfNamespaces.XLink + "href"));
        Assert.Equal("#Page%20Two", (string?)links[1].Attribute(OdfNamespaces.XLink + "href")); Assert.Equal("https://example.org/drawing", (string?)links[2].Attribute(OdfNamespaces.XLink + "href"));
    }

    [Theory]
    [InlineData("plain")]
    [InlineData("paragraph")]
    [InlineData("inline")]
    public void SourceFamilyDefaultsDoNotOverrideExplicitEnclosingGraphicText(string mode) {
        var source = OdgDocument.Create(); var page = source.AddPage();
        var box = page.Shapes.AddTextBox(Rect(0, 0, 200, 60), "Large text"); box.FontSize = OdfLength.Points(28);
        XElement paragraph = box.Element.Descendants(OdfNamespaces.Text + "p").Single();
        foreach (string family in new[] { "paragraph", "text", "graphic" }) source.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!.Add(
            new XElement(OdfNamespaces.Style + "default-style", new XAttribute(OdfNamespaces.Style + "family", family),
                new XElement(OdfNamespaces.Style + "text-properties", new XAttribute(OdfNamespaces.Fo + "font-size", family == "paragraph" ? "12pt" : "8pt"))));
        if (mode == "paragraph") { source.Styles.CreateNamed("Paragraph", OdfStyleFamily.Paragraph).SetProperty(OdfNamespaces.Style + "text-properties", OdfNamespaces.Fo + "color", "#123456"); paragraph.SetAttributeValue(OdfNamespaces.Text + "style-name", "Paragraph"); }
        if (mode == "inline") { source.Styles.CreateNamed("Inline", OdfStyleFamily.Text).Bold = true; paragraph.RemoveNodes(); paragraph.Add(new XElement(OdfNamespaces.Text + "span", new XAttribute(OdfNamespaces.Text + "style-name", "Inline"), "Large text")); }
        source.MarkPartDirty("content.xml"); source.MarkPartDirty("styles.xml");
        var target = OdgDocument.Create(); var imported = target.ImportPage(source, 0);
        Assert.Equal(Svg(page), Svg(imported));
        foreach (var read in RoundTrips(target)) Assert.Equal(Svg(page), Svg(read.Pages[0]));
        imported.Shapes[0].FontSize = OdfLength.Points(35);
        Assert.All(imported.ToDrawing().Value.Elements.OfType<OfficeDrawingRichText>().SelectMany(text => text.Paragraphs).SelectMany(paragraph => paragraph.Runs), run => Assert.Equal(35, run.FontSize));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AutomaticParentsResolveCommonStylesEvenWhenAnAutomaticStyleShadowsTheName(bool sameName) {
        var source = OdgDocument.Create(); var page = source.AddPage();
        var shape = page.Shapes.AddRectangle(Rect(0, 0, 100, 60));
        SetFill(source.Styles.CreateNamed("Base", OdfStyleFamily.Graphic), "#ff0000");
        XElement automatic = source.GetXml("content.xml").Root!.Element(OdfNamespaces.Office + "automatic-styles")!;
        automatic.Add(new XElement(OdfNamespaces.Style + "style", new XAttribute(OdfNamespaces.Style + "name", "Base"), new XAttribute(OdfNamespaces.Style + "family", "graphic"),
            new XElement(OdfNamespaces.Style + "graphic-properties", new XAttribute(OdfNamespaces.Draw + "fill", "solid"), new XAttribute(OdfNamespaces.Draw + "fill-color", "#0000ff"))));
        var active = source.Styles.FindInPart(OdfStyleFamily.Graphic, (string)shape.Element.Attribute(OdfNamespaces.Draw + "style-name")!, "content.xml")!;
        active.Element.RemoveNodes(); active.ParentStyleName = "Base"; source.MarkPartDirty("content.xml");
        if (sameName) {
            automatic.Elements(OdfNamespaces.Style + "style").Single(element => (string?)element.Attribute(OdfNamespaces.Style + "name") == "Base").Remove();
            active.Element.SetAttributeValue(OdfNamespaces.Style + "name", "Base"); shape.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", "Base"); source.MarkPartDirty("content.xml");
        }
        Assert.Equal(OdfColor.Parse("ff0000"), shape.FillColor);
        var target = OdgDocument.Create(); var imported = target.ImportPage(source, 0);
        Assert.Equal(OdfColor.Parse("ff0000"), imported.Shapes[0].FillColor);
        string importedStyle = (string)imported.Shapes[0].Element.Attribute(OdfNamespaces.Draw + "style-name")!;
        var parent = target.Styles.Find(OdfStyleFamily.Graphic, target.Styles.Find(OdfStyleFamily.Graphic, importedStyle)!.ParentStyleName!)!;
        Assert.False(parent.IsAutomatic);
    }

    [Fact]
    public void ImportsHatchAndStackedDashDefinitionsWithCollisionRemapping() {
        var source = OdgDocument.Create(); var page = source.AddPage(); var shape = page.Shapes.AddRectangle(Rect(0, 0, 100, 60));
        XElement definitions = source.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!;
        definitions.Add(new XElement(OdfNamespaces.Draw + "hatch", new XAttribute(OdfNamespaces.Draw + "name", "Paint"), new XAttribute(OdfNamespaces.Draw + "style", "single"),
            new XAttribute(OdfNamespaces.Draw + "color", "#123456"), new XAttribute(OdfNamespaces.Draw + "distance", "2pt"), new XAttribute(OdfNamespaces.Draw + "rotation", "0")));
        definitions.Add(new XElement(OdfNamespaces.Draw + "stroke-dash", new XAttribute(OdfNamespaces.Draw + "name", "Overlay"), new XAttribute(OdfNamespaces.Draw + "style", "rect"),
            new XAttribute(OdfNamespaces.Draw + "dots1", "1"), new XAttribute(OdfNamespaces.Draw + "dots1-length", "3pt"), new XAttribute(OdfNamespaces.Draw + "distance", "2pt")));
        var style = source.Styles.FindInPart(OdfStyleFamily.Graphic, (string)shape.Element.Attribute(OdfNamespaces.Draw + "style-name")!, "content.xml")!;
        style.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "fill", "hatch"); style.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "fill-hatch-name", "Paint");
        style.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "stroke-dash-names", "Overlay Overlay"); source.MarkPartDirty("styles.xml");
        var target = OdgDocument.Create(); target.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!.Add(new XElement(definitions.Element(OdfNamespaces.Draw + "hatch")!)); target.MarkPartDirty("styles.xml");
        var imported = target.ImportPage(source, 0); var copiedStyle = target.Styles.FindInPart(OdfStyleFamily.Graphic, (string)imported.Shapes[0].Element.Attribute(OdfNamespaces.Draw + "style-name")!, "content.xml")!;
        XElement properties = copiedStyle.Element.Element(OdfNamespaces.Style + "graphic-properties")!;
        string hatch = (string)properties.Attribute(OdfNamespaces.Draw + "fill-hatch-name")!; Assert.NotEqual("Paint", hatch);
        Assert.Contains(target.GetXml("styles.xml").Descendants(OdfNamespaces.Draw + "hatch"), element => (string?)element.Attribute(OdfNamespaces.Draw + "name") == hatch);
        string[] dashes = ((string)properties.Attribute(OdfNamespaces.Draw + "stroke-dash-names")!).Split(' '); Assert.Equal(dashes[0], dashes[1]); Assert.NotEqual("Overlay", dashes[0]);
        Assert.Contains(target.GetXml("styles.xml").Descendants(OdfNamespaces.Draw + "stroke-dash"), element => (string?)element.Attribute(OdfNamespaces.Draw + "name") == dashes[0]);
    }

    [Fact]
    public void UnnamedSourcePageGetsRequestedUniqueDestinationName() {
        var source = OdgDocument.Create(); source.AddPage().Element.Attribute(OdfNamespaces.Draw + "name")!.Remove(); source.MarkPartDirty("content.xml");
        var destination = OdgDocument.Create(); Assert.Equal("Review", destination.ImportPage(source, 0, "Review").Name);
    }

    [Fact]
    public void InlineScriptsAreRejectedBeforeImportMutation() {
        var source = OdgDocument.Create(); var page = source.AddPage(); var shape = page.Shapes.AddTextBox(Rect(0, 0, 100, 50), "Text");
        shape.Element.Descendants(OdfNamespaces.Text + "p").Single().Add(new XElement(OdfNamespaces.Text + "script", new XAttribute(OdfNamespaces.Script + "language", "JavaScript"), "void 0"));
        source.MarkPartDirty("content.xml"); var destination = OdgDocument.Create(); string[] before = Parts(destination);
        Assert.Throws<NotSupportedException>(() => destination.ImportPage(source, 0)); Assert.Equal(before, Parts(destination));
    }

    private static void SetFill(OdfStyle style, string color) {
        style.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "fill", "solid");
        style.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "fill-color", color);
    }
    private static void Default(OdgDocument document, string family, params XAttribute[] attributes) {
        document.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!.Add(new XElement(OdfNamespaces.Style + "default-style",
            new XAttribute(OdfNamespaces.Style + "family", family), new XElement(OdfNamespaces.Style + "graphic-properties", attributes)));
        document.MarkPartDirty("styles.xml");
    }
    private static string[] Parts(OdgDocument document) => new[] { document.GetXml("content.xml").ToString(), document.GetXml("styles.xml").ToString() };
    private static string Svg(OdgPage page) => OfficeDrawingSvgExporter.ToSvg(page.ToDrawing().Value);
    private static OdfRect Rect(double x, double y, double width, double height) => new OdfRect(OdfLength.Points(x), OdfLength.Points(y), OdfLength.Points(width), OdfLength.Points(height));
    private static OdgDocument[] RoundTrips(OdgDocument document) {
        using var flat = new MemoryStream(); document.SaveFlatXml(flat); flat.Position = 0;
        return new[] { OdgDocument.Load(new MemoryStream(document.ToBytes())), OdgDocument.LoadFlatXml(flat) };
    }
}
