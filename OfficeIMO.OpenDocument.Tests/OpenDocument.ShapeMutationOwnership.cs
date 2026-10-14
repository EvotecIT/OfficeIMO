using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentShapeMutationOwnershipTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MasterMutationsPersistWithPartScopedPaintAndClonedMasterIsolation(bool flat) {
        var source = OdgDocument.Create(); var page = source.AddPage("Source");
        AddPaint(source, "styles.xml", "Paint", "#ff0000", "#112233", "2pt");
        AddPaint(source, "content.xml", "Paint", "#00ff00", "#445566", "6pt");
        XElement artwork = Rectangle("Master", 10); artwork.SetAttributeValue(OdfNamespaces.Draw + "style-name", "Paint");
        page.Master!.Add(artwork);
        var local = page.Shapes.AddRectangle(Rect(30), "Local"); local.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", "Paint");
        var clone = source.ClonePage(0, "Independent"); clone.MasterPageName = source.CloneMasterPage(page.MasterPageName);
        source.MarkPartDirty("content.xml"); source.MarkPartDirty("styles.xml");
        var document = OdgDocument.Load(new MemoryStream(source.ToBytes()));
        byte[] beforeContent = document.GetPackageEntryBytes("content.xml");
        var edited = document.Pages[0].MasterShapes[0];

        edited.Name = "Edited master"; edited.Bounds = Rect(100); edited.Transform = "translate(5pt 6pt)";
        edited.FillColor = OdfColor.Parse("#0000ff");
        Assert.Equal(beforeContent, document.GetPackageEntryBytes("content.xml"));

        var reopened = Reopen(document, flat); var actual = reopened.Pages[0].MasterShapes[0];
        Assert.Equal("Edited master", actual.Name); Assert.Equal(100, actual.Bounds.X.ToPoints(), 6);
        Assert.Equal("translate(5pt 6pt)", actual.Transform); Assert.Equal(OdfColor.Parse("#0000FF"), actual.FillColor);
        Assert.Equal(OdfColor.Parse("#112233"), actual.StrokeColor); Assert.Equal(2, actual.StrokeWidth!.Value.ToPoints(), 6);
        Assert.Equal(OdfColor.Parse("#FF0000"), reopened.Pages[1].MasterShapes[0].FillColor);
        Assert.Equal(OdfColor.Parse("#00FF00"), reopened.Pages[0].Shapes[0].FillColor);
        Assert.Equal(OdfColor.Parse("#445566"), reopened.Pages[0].Shapes[0].StrokeColor);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MasterRenamePreservesIndependentPageConnectorUsingTheSameNativeIdentifier(bool legacyId) {
        var source = OdgDocument.Create(); var page = source.AddPage();
        XElement masterTarget = Rectangle("Master target", 100); XElement pageTarget = Rectangle("Page target", 10);
        SetNativeId(masterTarget, "shared", legacyId); SetNativeId(pageTarget, "shared", legacyId);
        page.Master!.Add(masterTarget, Connector("Master link", "shared"));
        page.Element.Add(pageTarget, Connector("Page link", "shared"));
        source.MarkPartDirty("styles.xml"); source.MarkPartDirty("content.xml");
        var document = OdgDocument.Load(new MemoryStream(source.ToBytes()));
        byte[] beforeContent = document.GetPackageEntryBytes("content.xml");
        var master = document.Pages[0].MasterShapes[0]; master.XmlId = "renamed";
        Assert.Throws<InvalidOperationException>(() => master.XmlId = null);
        Assert.Equal(beforeContent, document.GetPackageEntryBytes("content.xml"));
        Assert.Equal("renamed", (string?)master.Element.Attribute(OdfNamespaces.Draw + "id"));

        foreach (var reopened in new[] { Reopen(document, false), Reopen(document, true) }) {
            var actual = reopened.Pages[0];
            Assert.Equal("renamed", actual.MasterShapes[0].XmlId); Assert.Equal("renamed", actual.MasterShapes[1].StartShapeId);
            Assert.Equal("shared", actual.Shapes[0].XmlId); Assert.Equal("shared", actual.Shapes[1].StartShapeId);
            Assert.Equal(120, actual.MasterShapes[1].X1.ToPoints(), 6); Assert.Equal(30, actual.Shapes[1].X1.ToPoints(), 6);
            string before = reopened.GetXml("styles.xml").ToString();
            Assert.Throws<ArgumentException>(() => actual.MasterShapes[0].XmlId = "shared");
            Assert.Equal(before, reopened.GetXml("styles.xml").ToString());
        }
    }

    [Fact]
    public void AttachedShapeIdentitiesReserveBothMainPartsAndLegacyTextAliases() {
        var source = OdgDocument.Create(); var page = source.AddPage();
        XElement firstNative = Rectangle("Reserved", 100); SetNativeId(firstNative, "shape1", false);
        XElement secondNative = Rectangle("Legacy reserved", 130); SetNativeId(secondNative, "shape2", true);
        secondNative.Add(new XElement(OdfNamespaces.Text + "p", new XAttribute(OdfNamespaces.Text + "id", "shape3"), "Label"));
        page.Master!.Add(firstNative, secondNative); source.MarkPartDirty("styles.xml");
        var first = page.Shapes.AddRectangle(Rect(10)); var second = page.Shapes.AddRectangle(Rect(40));
        var connector = page.Shapes.AddConnector(first.AddGluePoint(), second.AddGluePoint());
        var masterFirst = page.MasterShapes.AddRectangle(Rect(160)); var masterSecond = page.MasterShapes.AddRectangle(Rect(190));
        var masterConnector = page.MasterShapes.AddConnector(masterFirst.AddGluePoint(), masterSecond.AddGluePoint());
        string[] generated = { first.XmlId!, second.XmlId!, masterFirst.XmlId!, masterSecond.XmlId! };
        Assert.Equal(generated.Length, generated.Distinct(StringComparer.Ordinal).Count());
        Assert.All(generated, id => Assert.DoesNotContain(id, new[] { "shape1", "shape2", "shape3" }));
        Assert.Equal(first.XmlId, connector.StartShapeId); Assert.Equal(masterFirst.XmlId, masterConnector.StartShapeId);
        string[] before = { source.GetXml("content.xml").ToString(), source.GetXml("styles.xml").ToString() };
        Assert.Throws<ArgumentException>(() => first.XmlId = "shape2");
        Assert.Throws<ArgumentException>(() => masterFirst.XmlId = first.XmlId);
        Assert.Equal(before, new[] { source.GetXml("content.xml").ToString(), source.GetXml("styles.xml").ToString() });
        foreach (var reopened in new[] { Reopen(source, false), Reopen(source, true) }) {
            var actual = reopened.Pages[0];
            Assert.Equal(generated, new[] { actual.Shapes[0].XmlId, actual.Shapes[1].XmlId, actual.MasterShapes[2].XmlId, actual.MasterShapes[3].XmlId });
            Assert.Equal(actual.Shapes[0].XmlId, actual.Shapes[2].StartShapeId);
            Assert.Equal(actual.MasterShapes[2].XmlId, actual.MasterShapes[4].StartShapeId);
        }
    }

    [Fact]
    public void AnimationTargetFollowsRenameAndRejectsIdentifierRemovalAfterReopening() {
        var presentation = OdpPresentation.Create(); var slide = presentation.AddSlide("Animated");
        var shape = slide.AddRectangle(Rect(10));
        var animation = slide.AddFadeInAnimation(shape, TimeSpan.FromSeconds(1));
        shape.XmlId = "renamed";
        Assert.Equal("renamed", animation.TargetElement);
        Assert.Throws<InvalidOperationException>(() => shape.XmlId = null);
        foreach (bool flat in new[] { false, true }) {
            OdpPresentation reopened;
            if (flat) { using var stream = new MemoryStream(); presentation.SaveFlatXml(stream); stream.Position = 0; reopened = OdpPresentation.LoadFlatXml(stream); }
            else reopened = OdpPresentation.Load(new MemoryStream(presentation.ToBytes()));
            var actual = reopened.Slides[0]; var target = actual.Shapes[0];
            Assert.Equal("renamed", target.XmlId); Assert.Equal(target.XmlId, Assert.Single(actual.Animations).TargetElement);
            Assert.Throws<InvalidOperationException>(() => target.XmlId = null);
            target.XmlId = "renamed-again";
            Assert.Equal(target.XmlId, Assert.Single(actual.Animations).TargetElement);
        }
    }

    private static OdfRect Rect(double x) => new OdfRect(OdfLength.Points(x), OdfLength.Points(10), OdfLength.Points(20), OdfLength.Points(20));
    private static XElement Rectangle(string name, double x) {
        var element = new XElement(OdfNamespaces.Draw + "rect", new XAttribute(OdfNamespaces.Draw + "name", name));
        OdfShape.ApplyBounds(element, Rect(x)); return element;
    }
    private static void SetNativeId(XElement element, string id, bool legacyOnly) {
        element.SetAttributeValue(OdfNamespaces.Draw + "id", id);
        if (!legacyOnly) element.SetAttributeValue(XNamespace.Xml + "id", id);
    }
    private static XElement Connector(string name, string target) => new XElement(OdfNamespaces.Draw + "connector",
        new XAttribute(OdfNamespaces.Draw + "name", name), new XAttribute(OdfNamespaces.Draw + "type", "line"),
        new XAttribute(OdfNamespaces.Draw + "start-shape", target), new XAttribute(OdfNamespaces.Draw + "start-glue-point", "1"),
        new XAttribute(OdfNamespaces.Svg + "x2", "220pt"), new XAttribute(OdfNamespaces.Svg + "y2", "20pt"));
    private static void AddPaint(OdgDocument document, string part, string name, string fill, string stroke, string width) =>
        document.GetXml(part).Root!.Element(OdfNamespaces.Office + "automatic-styles")!.Add(new XElement(OdfNamespaces.Style + "style",
            new XAttribute(OdfNamespaces.Style + "name", name), new XAttribute(OdfNamespaces.Style + "family", "graphic"),
            new XElement(OdfNamespaces.Style + "graphic-properties", new XAttribute(OdfNamespaces.Draw + "fill", "solid"),
                new XAttribute(OdfNamespaces.Draw + "fill-color", fill), new XAttribute(OdfNamespaces.Draw + "stroke", "solid"),
                new XAttribute(OdfNamespaces.Svg + "stroke-color", stroke), new XAttribute(OdfNamespaces.Svg + "stroke-width", width))));
    private static OdgDocument Reopen(OdgDocument document, bool flat) {
        if (!flat) return OdgDocument.Load(new MemoryStream(document.ToBytes(new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.PreserveSource })));
        using var stream = new MemoryStream(); document.SaveFlatXml(stream); stream.Position = 0; return OdgDocument.LoadFlatXml(stream);
    }
}
