using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public class OpenDocumentOdgConnectorTransformTests {
    [Theory]
    [InlineData("translate(20.17pt 30.005pt)", 1, 0, 0, 1, 20.17, 30.005)]
    [InlineData("scale(1.3 0.7) translate(20pt 30pt)", 1.3, 0, 0, 0.7, 20, 30)]
    [InlineData("rotate(1.5707963267948966) translate(40pt 280pt)", 0, -1, 1, 0, 40, 280)]
    [InlineData("skewX(0.2) translate(60pt 30pt)", 1, 0, -0.2027100355086725, 1, 60, 30)]
    [InlineData("matrix(1.1 0.2 0.15 0.8 30pt 20pt)", 1.1, 0.2, 0.15, 0.8, 30, 20)]
    [InlineData("scale(-1 1) translate(320pt 20pt)", -1, 0, 0, 1, 320, 20)]
    public void BakesEveryCurveControlAndEndpointWithoutChangingKindStylesOrMetadata(string expression,
        double m11, double m12, double m21, double m22, double x, double y) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var connector = page.Shapes.AddConnector(new OfficePoint(30, 40), new OfficePoint(250, 100), OdgConnectorKind.Curve);
        connector.SetConnectorRoute(new[] { OfficePathCommand.MoveTo(30, 40), OfficePathCommand.LineTo(30, 140),
            OfficePathCommand.QuadraticBezierTo(70, 160, 120, 80), OfficePathCommand.CubicBezierTo(140, 0, 200, 160, 250, 100) });
        connector.Name = "Retained connector"; connector.StrokeColor = OdfColor.Parse("BB2233"); connector.StrokeWidth = OdfLength.Points(2);
        connector.Element.Add(new System.Xml.Linq.XElement(OdfNamespaces.Svg + "title", "Route metadata"));
        connector.Element.SetAttributeValue(OdfNamespaces.Draw + "line-skew", "2cm 3cm");
        connector.Transform = expression;
        var transform = new OfficeTransform(m11, m12, m21, m22, x, y);
        var expected = connector.ConnectorRouteCommands.Select(command => Map(command, transform)).ToArray();
        string styles = document.GetXml("styles.xml").ToString();
        connector.BakeConnectorTransform();
        Assert.Equal(styles, document.GetXml("styles.xml").ToString());
        foreach (var read in RoundTrips(document)) {
            var saved = read.Pages[0].Shapes[0]; Assert.Null(saved.Transform);
            Assert.Equal(OdgConnectorKind.Curve, saved.ConnectorKind); Assert.Null(saved.StartShapeId); Assert.Null(saved.EndShapeId);
            Assert.Equal("Retained connector", saved.Name); Assert.Equal(OdfLength.Points(2), saved.StrokeWidth);
            Assert.Equal(OdfColor.Parse("BB2233"), saved.StrokeColor);
            Assert.Equal("Route metadata", saved.Element.Element(OdfNamespaces.Svg + "title")!.Value);
            Assert.Null(saved.Element.Attribute(OdfNamespaces.Draw + "line-skew"));
            AssertCommands(expected, saved.ConnectorRouteCommands);
            Assert.False(read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Report.HasSkippedOrUnsupported);
            string before = saved.ToXml().ToString(); saved.BakeConnectorTransform(); Assert.Equal(before, saved.ToXml().ToString());
        }
    }

    [Fact]
    public void BakesStraightConnectorWithoutAnExistingRouteCache() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var connector = page.Shapes.AddConnector(new OfficePoint(10, 20), new OfficePoint(100, 80));
        connector.Transform = "translate(30pt 40pt)";
        connector.BakeConnectorTransform();
        Assert.Equal(OdgConnectorKind.Line, connector.ConnectorKind); Assert.Null(connector.Transform);
        AssertCommands(new[] { OfficePathCommand.MoveTo(40, 60), OfficePathCommand.LineTo(130, 120) }, connector.ConnectorRouteCommands);
    }

    [Theory]
    [InlineData("partial-attachment")]
    [InlineData("transformed-label")]
    [InlineData("empty-field-label")]
    [InlineData("transformed-parent")]
    [InlineData("missing-cache")]
    [InlineData("coordinate-overflow")]
    public void UnsupportedBakeLeavesEveryDocumentPartUnchanged(string failure) {
        var document = OdgDocument.Create(); var page = document.AddPage(); var group = page.Shapes.AddGroup();
        var connector = group.Children.AddConnector(new OfficePoint(10, 20), new OfficePoint(100, 80));
        connector.Transform = "translate(30pt 40pt)";
        switch (failure) {
            case "partial-attachment": connector.AttachEndToShape(page.Shapes.AddRectangle(Rect(100, 70, 20, 20))); break;
            case "transformed-label": connector.Text = "Keep label placement"; connector.Transform = "scale(2 1)"; break;
            case "empty-field-label": connector.Element.Add(new System.Xml.Linq.XElement(OdfNamespaces.Text + "p",
                new System.Xml.Linq.XElement(OdfNamespaces.Text + "page-number"))); connector.Transform = "scale(2 1)"; break;
            case "transformed-parent": group.Element.SetAttributeValue(OdfNamespaces.Draw + "transform", "translate(20pt 30pt)"); break;
            case "missing-cache": connector.ConnectorKind = OdgConnectorKind.Standard; break;
            case "coordinate-overflow": connector.Transform = "translate(1000000000pt 0pt)"; break;
        }
        string[] before = Parts(document);
        if (failure == "coordinate-overflow") Assert.Throws<ArgumentOutOfRangeException>(() => connector.BakeConnectorTransform());
        else Assert.Throws<NotSupportedException>(() => connector.BakeConnectorTransform());
        Assert.Equal(before, Parts(document));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ImportedEmptyParagraphDoesNotBlockFreeOrAttachedTransformBaking(bool attached) {
        var document = OdgDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", "libreoffice-routed-glue.odg"));
        var connector = document.Pages[0].Shapes.First(shape => shape.IsConnector);
        string textBefore = connector.Element.Element(OdfNamespaces.Text + "p")!.ToString();
        if (attached) {
            connector.Transform = "translate(20pt 30pt)";
            connector.RouteOrthogonal(OdgConnectorRouteOrientation.VerticalFirst, 64);
            connector.UseNativeThreeSegmentRouting();
        } else {
            var route = connector.ConnectorRouteCommands;
            connector.AttachStart(null); connector.AttachEnd(null); connector.SetConnectorRoute(route);
            connector.Transform = "translate(20pt 30pt)";
            connector.BakeConnectorTransform();
        }
        Assert.Null(connector.Transform);
        Assert.Equal(textBefore, connector.Element.Element(OdfNamespaces.Text + "p")!.ToString());
        Assert.Equal(4, connector.ConnectorRouteCommands.Count);
    }

    [Fact]
    public void GroupBakeRejectsLabeledConnectorBeforeChangingEarlierSiblings() {
        var document = OdgDocument.Create(); var group = document.AddPage().Shapes.AddGroup();
        group.Children.AddRectangle(Rect(20, 30, 40, 40));
        group.Children.AddConnector(new OfficePoint(10, 20), new OfficePoint(100, 80));
        var labeled = group.Children.AddGroup().Children.AddConnector(new OfficePoint(20, 30), new OfficePoint(120, 90));
        labeled.Text = "Preserve label transform";
        string[] before = Parts(document);
        Assert.Throws<NotSupportedException>(() => group.TransformChildren("scale(2 1) translate(10pt 20pt)"));
        Assert.Equal(before, Parts(document));
    }

    [Fact]
    public void SharedRouteWriterRejectsNativeCoordinateOverflowBeforeReplacingSavedRoute() {
        var document = OdgDocument.Create(); var connector = document.AddPage().Shapes.AddConnector(new OfficePoint(10, 20), new OfficePoint(100, 80));
        connector.SetConnectorRoute(new[] { OfficePathCommand.MoveTo(10, 20), OfficePathCommand.LineTo(100, 80) });
        string[] before = Parts(document);
        Assert.Throws<ArgumentOutOfRangeException>(() => connector.SetConnectorRoute(new[] {
            OfficePathCommand.MoveTo(1e9, 20), OfficePathCommand.LineTo(1e9 + 90, 80) }));
        Assert.Equal(before, Parts(document));
    }

    private static OfficePathCommand Map(OfficePathCommand command, OfficeTransform transform) => command.Kind switch {
        OfficePathCommandKind.MoveTo => OfficePathCommand.MoveTo(transform.TransformPoint(command.Point)),
        OfficePathCommandKind.LineTo => OfficePathCommand.LineTo(transform.TransformPoint(command.Point)),
        OfficePathCommandKind.QuadraticBezierTo => OfficePathCommand.QuadraticBezierTo(transform.TransformPoint(command.ControlPoint1), transform.TransformPoint(command.Point)),
        OfficePathCommandKind.CubicBezierTo => OfficePathCommand.CubicBezierTo(transform.TransformPoint(command.ControlPoint1), transform.TransformPoint(command.ControlPoint2), transform.TransformPoint(command.Point)),
        _ => throw new InvalidOperationException()
    };
    private static void AssertCommands(IReadOnlyList<OfficePathCommand> expected, IReadOnlyList<OfficePathCommand> actual) {
        Assert.Equal(expected.Count, actual.Count);
        for (int i = 0; i < expected.Count; i++) {
            Assert.Equal(expected[i].Kind, actual[i].Kind); AssertPoint(expected[i].Point, actual[i].Point);
            if (expected[i].Kind is OfficePathCommandKind.QuadraticBezierTo or OfficePathCommandKind.CubicBezierTo) AssertPoint(expected[i].ControlPoint1, actual[i].ControlPoint1);
            if (expected[i].Kind == OfficePathCommandKind.CubicBezierTo) AssertPoint(expected[i].ControlPoint2, actual[i].ControlPoint2);
        }
    }
    private static void AssertPoint(OfficePoint expected, OfficePoint actual) {
        Assert.InRange(Math.Abs(expected.X - actual.X), 0, 0.02); Assert.InRange(Math.Abs(expected.Y - actual.Y), 0, 0.02);
    }
    private static string[] Parts(OdgDocument document) => new[] { document.GetXml("content.xml").ToString(), document.GetXml("styles.xml").ToString(), document.GetXml("settings.xml").ToString() };
    private static IEnumerable<OdgDocument> RoundTrips(OdgDocument document) {
        yield return document; yield return OdgDocument.Load(new MemoryStream(document.ToBytes()));
        using var stream = new MemoryStream(); document.SaveFlatXml(stream); stream.Position = 0; yield return OdgDocument.LoadFlatXml(stream);
    }
    private static OdfRect Rect(double x, double y, double w, double h) => new(OdfLength.Points(x), OdfLength.Points(y), OdfLength.Points(w), OdfLength.Points(h));
}
