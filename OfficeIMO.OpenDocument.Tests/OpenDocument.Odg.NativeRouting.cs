using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public class OpenDocumentOdgNativeRoutingTests {
    [Fact]
    public void BottomCenterRelativePointFollowsResizeAndRejectsAbsoluteOffsetsAtomically() {
        var document = OdgDocument.Create(); var shape = document.AddPage().Shapes.AddRectangle(Rect(10, 20, 40, 40));
        var point = shape.AddGluePoint(OdgGluePointAlignment.Bottom);
        AssertClose(new OfficePoint(30, 60), point.Position);
        shape.Bounds = Rect(10, 20, 80, 20);
        AssertClose(new OfficePoint(50, 40), point.Position);
        string before = shape.ToXml().ToString();
        Assert.Throws<NotSupportedException>(() => shape.AddGluePoint(OdgGluePointAlignment.Bottom, offsetX: OdfLength.Points(5)));
        Assert.Equal(before, shape.ToXml().ToString());
    }

    [Fact]
    public void NormalizesPercentageGlueInDirectionalCopyAndRetainsOriginalPercentagePoint() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var start = page.Shapes.AddRectangle(Rect(30, 30, 40, 40)); var end = page.Shapes.AddRectangle(Rect(300, 100, 40, 40));
        var point = start.AddGluePoint(); point.Element.Attribute(OdfNamespaces.Draw + "align")!.Remove();
        point.Element.SetAttributeValue(OdfNamespaces.Svg + "x", "0%"); point.Element.SetAttributeValue(OdfNamespaces.Svg + "y", "50%");
        var connector = page.Shapes.AddConnector(point, end.AddGluePoint(OdgGluePointAlignment.Left));
        string before = point.Element.ToString();
        connector.ConnectorKind = OdgConnectorKind.Standard;
        connector.SetConnectorRoute(Route(new OfficePoint(50, 70), new OfficePoint(50, 160), new OfficePoint(300, 160), new OfficePoint(300, 120)));
        connector.UseNativeThreeSegmentRouting();
        Assert.Equal(before, point.Element.ToString());
        AssertClose(new OfficePoint(50, 70), start.GluePoints.Last().Position);
        Assert.Equal("5cm", (string?)start.GluePoints.Last().Element.Attribute(OdfNamespaces.Svg + "y"));
        connector.UseNativeThreeSegmentRouting(); Assert.Equal(2, start.GluePoints.Count);
    }

    [Fact]
    public void EncodesDetourWithoutChangingSharedGlueOrGraphicStylesAndReusesDirectionalPoints() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var start = page.Shapes.AddRectangle(Rect(30, 80, 40, 40)); var end = page.Shapes.AddEllipse(Rect(340, 80, 40, 40));
        var a = start.AddGluePoint(OdgGluePointAlignment.Right); var b = end.AddGluePoint(OdgGluePointAlignment.Left);
        var connector = page.Shapes.AddConnector(a, b); connector.StrokeColor = OdfColor.Parse("#BB2233");
        var other = page.Shapes.AddConnector(a, b);
        string styleName = (string)connector.Element.Attribute(OdfNamespaces.Draw + "style-name")!;
        other.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", styleName);
        string aBefore = a.Element.ToString(), bBefore = b.Element.ToString(), otherBefore = other.ToXml().ToString();
        string styleBefore = document.Styles.Find(OdfStyleFamily.Graphic, styleName)!.Element.ToString();
        connector.ConnectorKind = OdgConnectorKind.Standard;
        connector.SetConnectorRoute(Route(new OfficePoint(70, 100), new OfficePoint(70, 164), new OfficePoint(340, 164), new OfficePoint(340, 100)));
        connector.UseNativeThreeSegmentRouting();
        Assert.Equal(aBefore, a.Element.ToString()); Assert.Equal(bBefore, b.Element.ToString());
        Assert.Equal(otherBefore, other.ToXml().ToString());
        Assert.Equal(styleBefore, document.Styles.Find(OdfStyleFamily.Graphic, styleName)!.Element.ToString());
        Assert.NotEqual(styleName, (string?)connector.Element.Attribute(OdfNamespaces.Draw + "style-name"));
        foreach (var read in RoundTrips(document)) {
            var saved = read.Pages[0].Shapes.First(shape => shape.IsConnector);
            Assert.Equal(OdgConnectorKind.Lines, saved.ConnectorKind); Assert.Equal(4, saved.ConnectorRouteCommands.Count);
            string[] offsets = ((string)saved.Element.Attribute(OdfNamespaces.Draw + "line-skew")!).Split(' ');
            Assert.Equal(2, offsets.Length);
            Assert.All(offsets, offset => Assert.InRange(Math.Abs(OdfLength.Parse(offset).ToPoints() - 44), 0, 0.1));
            Assert.Equal(OdgGluePointEscapeDirection.Down, read.Pages[0].Shapes[0].GluePoints.Last().EscapeDirection);
            Assert.Equal(OdgGluePointEscapeDirection.Down, read.Pages[0].Shapes[1].GluePoints.Last().EscapeDirection);
            Assert.False(read.Pages[0].ToDrawing().Report.HasSkippedOrUnsupported);
        }
        connector.UseNativeThreeSegmentRouting();
        Assert.Equal(2, start.GluePoints.Count); Assert.Equal(2, end.GluePoints.Count);
        foreach (double lane in new[] { 40D, 164D, 40D, 164D }) {
            connector.SetConnectorRoute(Route(new OfficePoint(70, 100), new OfficePoint(70, lane), new OfficePoint(340, lane), new OfficePoint(340, 100)));
            connector.UseNativeThreeSegmentRouting();
        }
        Assert.Equal(3, start.GluePoints.Count); Assert.Equal(3, end.GluePoints.Count);
    }

    [Theory]
    [InlineData("translate(20pt 30pt)", OdgConnectorRouteOrientation.VerticalFirst)]
    [InlineData("translate(20.17pt 30.23pt)", OdgConnectorRouteOrientation.VerticalFirst)]
    [InlineData("translate(20pt 30.005pt)", OdgConnectorRouteOrientation.VerticalFirst)]
    [InlineData("scale(2 2)", OdgConnectorRouteOrientation.VerticalFirst)]
    [InlineData("rotate(1.5707963267948966)", OdgConnectorRouteOrientation.HorizontalFirst)]
    public void ConvertsAutomaticAttachmentsAndCanRepeatAfterBakingConnectorTransform(string transform, OdgConnectorRouteOrientation orientation) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var start = page.Shapes.AddRectangle(Rect(30, 80, 40, 40)); var end = page.Shapes.AddRectangle(Rect(340, 80, 40, 40));
        var connector = page.Shapes.AddConnector(new OfficePoint(50, 70), new OfficePoint(320, 70));
        connector.Transform = transform;
        connector.AttachStartToShape(start); connector.AttachEndToShape(end);
        connector.RouteOrthogonal(orientation, 64);
        connector.UseNativeThreeSegmentRouting();
        foreach (var read in RoundTrips(document)) {
            var saved = read.Pages[0].Shapes[2]; Assert.Null(saved.Transform);
            Assert.Equal(4, saved.ConnectorRouteCommands.Count);
            AssertClose(new OfficePoint(70, 100), saved.ConnectorRouteCommands[0].Point);
            AssertClose(new OfficePoint(340, 100), saved.ConnectorRouteCommands[3].Point);
            Assert.Equal("right", (string?)read.Pages[0].Shapes[0].GluePoints.Single().Element.Attribute(OdfNamespaces.Draw + "align"));
            Assert.Equal("left", (string?)read.Pages[0].Shapes[1].GluePoints.Single().Element.Attribute(OdfNamespaces.Draw + "align"));
            string before = saved.ToXml().ToString();
            saved.UseNativeThreeSegmentRouting(); saved.UseNativeThreeSegmentRouting();
            Assert.Equal(before, saved.ToXml().ToString());
        }
    }

    [Fact]
    public void ReservesDistinctPointIdsForBothEndsOnTheSameShape() {
        var document = OdgDocument.Create(); var page = document.AddPage(); var shape = page.Shapes.AddRectangle(Rect(30, 80, 40, 40));
        var connector = page.Shapes.AddConnector(shape.AddGluePoint(OdgGluePointAlignment.Right), shape.AddGluePoint(OdgGluePointAlignment.Left));
        connector.ConnectorKind = OdgConnectorKind.Standard;
        connector.SetConnectorRoute(Route(new OfficePoint(70, 100), new OfficePoint(70, 160), new OfficePoint(30, 160), new OfficePoint(30, 100)));
        connector.UseNativeThreeSegmentRouting();
        Assert.Equal(4, shape.GluePoints.Count); Assert.Equal(4, shape.GluePoints.Select(point => point.Id).Distinct().Count());
        Assert.NotEqual((string?)connector.Element.Attribute(OdfNamespaces.Draw + "start-glue-point"), (string?)connector.Element.Attribute(OdfNamespaces.Draw + "end-glue-point"));
        connector.UseNativeThreeSegmentRouting(); Assert.Equal(4, shape.GluePoints.Count);
    }

    [Fact]
    public void RetainsExistingProducerGlueAndSiblingConnector() {
        var document = OdgDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", "libreoffice-routed-glue.odg"));
        var page = document.Pages[0]; var connectors = page.Shapes.Where(shape => shape.IsConnector).ToArray();
        var oldPoints = page.Shapes.Where(shape => !shape.IsConnector).SelectMany(shape => shape.GluePoints).Select(point => point.Element.ToString()).ToArray();
        var other = connectors[1].ToXml().ToString();
        connectors[0].UseNativeThreeSegmentRouting();
        Assert.Equal(other, connectors[1].ToXml().ToString());
        var newPoints = page.Shapes.Where(shape => !shape.IsConnector).SelectMany(shape => shape.GluePoints).Select(point => point.Element.ToString()).ToArray();
        Assert.All(oldPoints, point => Assert.Contains(point, newPoints));
        Assert.Equal(4, OdgDocument.Load(new MemoryStream(document.ToBytes())).Pages[0].Shapes.First(shape => shape.IsConnector).ConnectorRouteCommands.Count);
    }

    [Theory]
    [InlineData("free-end")]
    [InlineData("curve")]
    [InlineData("diagonal-departure")]
    [InlineData("zero-departure")]
    [InlineData("transformed-end")]
    [InlineData("transformed-text")]
    [InlineData("offset-overflow")]
    [InlineData("radial-circle")]
    [InlineData("grid-zero-departure")]
    public void RejectedNativeEncodingPreservesAllPartsAndAttachments(string failure) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var start = page.Shapes.AddRectangle(Rect(30, 80, 40, 40)); var end = page.Shapes.AddRectangle(Rect(340, 80, 40, 40));
        var connector = page.Shapes.AddConnector(start.AddGluePoint(OdgGluePointAlignment.Right), end.AddGluePoint(OdgGluePointAlignment.Left));
        connector.ConnectorKind = OdgConnectorKind.Standard;
        connector.SetConnectorRoute(Route(new OfficePoint(70, 100), new OfficePoint(70, 164), new OfficePoint(340, 164), new OfficePoint(340, 100)));
        switch (failure) {
            case "free-end": connector.AttachEnd(null); break;
            case "curve": connector.ConnectorKind = OdgConnectorKind.Curve;
                connector.SetConnectorRoute(new[] { OfficePathCommand.MoveTo(70, 100), OfficePathCommand.CubicBezierTo(70, 164, 340, 164, 340, 100) }); break;
            case "diagonal-departure": connector.SetConnectorRoute(Route(new OfficePoint(70, 100), new OfficePoint(80, 164), new OfficePoint(340, 164), new OfficePoint(340, 100))); break;
            case "zero-departure": connector.SetConnectorRoute(Route(new OfficePoint(70, 100), new OfficePoint(70, 100), new OfficePoint(340, 164), new OfficePoint(340, 100))); break;
            case "transformed-end": end.Transform = "translate(10pt 20pt)";
                connector.SetConnectorRoute(Route(new OfficePoint(70, 100), new OfficePoint(70, 164), new OfficePoint(350, 164), new OfficePoint(350, 120))); break;
            case "transformed-text": connector.Transform = "translate(20pt 30pt)"; connector.Text = "Connector label"; break;
            case "offset-overflow": start.Bounds = Rect(70, -1e9, 1e9, 2e9); connector.AttachStart(start.AddGluePoint(OdgGluePointAlignment.Left));
                connector.SetConnectorRoute(Route(new OfficePoint(70, 0), new OfficePoint(70, 164), new OfficePoint(340, 164), new OfficePoint(340, 100))); break;
            case "radial-circle": start.Element.Name = OdfNamespaces.Draw + "circle";
                start.Element.Attribute(OdfNamespaces.Svg + "width")!.Remove(); start.Element.Attribute(OdfNamespaces.Svg + "height")!.Remove();
                start.Element.SetAttributeValue(OdfNamespaces.Svg + "r", "20pt");
                connector.SetConnectorRoute(Route(new OfficePoint(30, 80), new OfficePoint(30, 164), new OfficePoint(340, 164), new OfficePoint(340, 100))); break;
            case "grid-zero-departure": connector.Transform = "scale(0.001 0.001)";
                connector.RouteOrthogonal(OdgConnectorRouteOrientation.VerticalFirst, offset: 0.05); break;
        }
        var before = Parts(document);
        if (failure == "offset-overflow") Assert.Throws<ArgumentOutOfRangeException>(() => connector.UseNativeThreeSegmentRouting());
        else Assert.Throws<NotSupportedException>(() => connector.UseNativeThreeSegmentRouting());
        Assert.Equal(before, Parts(document));
    }

    [Theory]
    [InlineData(OdgGluePointEscapeDirection.Auto)]
    [InlineData(OdgGluePointEscapeDirection.Left)]
    [InlineData(OdgGluePointEscapeDirection.Right)]
    [InlineData(OdgGluePointEscapeDirection.Up)]
    [InlineData(OdgGluePointEscapeDirection.Down)]
    [InlineData(OdgGluePointEscapeDirection.Horizontal)]
    [InlineData(OdgGluePointEscapeDirection.Vertical)]
    public void EscapeConstraintsRoundTripAndRejectInvalidValuesWithoutMutation(OdgGluePointEscapeDirection direction) {
        var document = OdgDocument.Create(); var point = document.AddPage().Shapes.AddRectangle(Rect(0, 0, 40, 40)).AddGluePoint();
        point.EscapeDirection = direction;
        foreach (var read in RoundTrips(document)) Assert.Equal(direction, read.Pages[0].Shapes[0].GluePoints.Single().EscapeDirection);
        string before = point.Element.ToString();
        Assert.Throws<ArgumentOutOfRangeException>(() => point.EscapeDirection = (OdgGluePointEscapeDirection)100);
        Assert.Equal(before, point.Element.ToString());
    }

    private static string[] Parts(OdgDocument document) => new[] { document.GetXml("content.xml").ToString(), document.GetXml("styles.xml").ToString(), document.GetXml("settings.xml").ToString() };
    private static IEnumerable<OdgDocument> RoundTrips(OdgDocument document) {
        yield return document; yield return OdgDocument.Load(new MemoryStream(document.ToBytes()));
        using var stream = new MemoryStream(); document.SaveFlatXml(stream); stream.Position = 0; yield return OdgDocument.LoadFlatXml(stream);
    }
    private static OfficePathCommand[] Route(params OfficePoint[] points) => points.Select((point, index) => index == 0 ? OfficePathCommand.MoveTo(point) : OfficePathCommand.LineTo(point)).ToArray();
    private static OdfRect Rect(double x, double y, double w, double h) => new(OdfLength.Points(x), OdfLength.Points(y), OdfLength.Points(w), OdfLength.Points(h));
    private static void AssertClose(OfficePoint expected, OfficePoint actual) { Assert.InRange(Math.Abs(expected.X - actual.X), 0, 0.1); Assert.InRange(Math.Abs(expected.Y - actual.Y), 0, 0.1); }
}
