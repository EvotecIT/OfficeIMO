using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public class OpenDocumentOdgNativePartialLineTests {
    [Theory]
    [InlineData("libreoffice-partial-line-master-start.odg", true, true)]
    [InlineData("libreoffice-partial-line-master-end.odg", true, false)]
    [InlineData("libreoffice-partial-line-page-start.odg", false, true)]
    [InlineData("libreoffice-partial-line-page-end.odg", false, false)]
    public void NativePartialLineCachesRetainAttachmentsAndRequireRefreshAfterMovingTheirTarget(string fixture, bool master, bool start) {
        var document = OdgDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", fixture));
        Assert.Equal(2, document.Pages.Count);
        foreach (var page in document.Pages)
            Assert.False(page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Report.HasSkippedOrUnsupported);

        var shapes = Shapes(document.Pages[0], master);
        var connector = shapes.Single(shape => shape.IsConnector);
        var target = shapes.Single(shape => !shape.IsConnector);
        var oldAttachment = AttachedPosition(connector, start);
        var untouchedClone = AttachedPosition(Shapes(document.Pages[1], master).Single(shape => shape.IsConnector), start);
        Assert.Equal(OdgConnectorKind.Line, connector.ConnectorKind);
        Assert.Equal(target.XmlId, start ? connector.StartShapeId : connector.EndShapeId);
        Assert.Null(start ? connector.EndShapeId : connector.StartShapeId);

        var bounds = target.Bounds;
        target.Bounds = new OdfRect(OdfLength.Points(bounds.X.ToPoints() + 3), OdfLength.Points(bounds.Y.ToPoints() + 4), bounds.Width, bounds.Height);
        Assert.Throws<NotSupportedException>(() => connector.ConnectorRouteCommands);
        Assert.Throws<OdfConversionLossException>(() => document.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        connector.SetConnectorRoute(new[] {
            OfficePathCommand.MoveTo(connector.X1.ToPoints(), connector.Y1.ToPoints()),
            OfficePathCommand.LineTo(connector.X2.ToPoints(), connector.Y2.ToPoints())
        });

        using var flat = new MemoryStream(); document.SaveFlatXml(flat); flat.Position = 0;
        foreach (var reopened in new[] { OdgDocument.Load(new MemoryStream(document.ToBytes())), OdgDocument.LoadFlatXml(flat) }) {
            var saved = Shapes(reopened.Pages[0], master).Single(shape => shape.IsConnector);
            var savedTarget = Shapes(reopened.Pages[0], master).Single(shape => !shape.IsConnector);
            Assert.Equal(savedTarget.XmlId, start ? saved.StartShapeId : saved.EndShapeId);
            Assert.Null(start ? saved.EndShapeId : saved.StartShapeId);
            var route = saved.ConnectorRouteCommands;
            Assert.Equal(2, route.Count);
            AssertClose(new OfficePoint(oldAttachment.X + 3, oldAttachment.Y + 4), route[start ? 0 : 1].Point);
            AssertClose(untouchedClone, AttachedPosition(Shapes(reopened.Pages[1], master).Single(shape => shape.IsConnector), start));
            foreach (var page in reopened.Pages)
                Assert.False(page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Report.HasSkippedOrUnsupported);
        }
    }

    private static OdgShapes Shapes(OdgPage page, bool master) => master ? page.MasterShapes : page.Shapes;
    private static OfficePoint AttachedPosition(OdgShape connector, bool start) => start
        ? new OfficePoint(connector.X1.ToPoints(), connector.Y1.ToPoints())
        : new OfficePoint(connector.X2.ToPoints(), connector.Y2.ToPoints());
    private static void AssertClose(OfficePoint expected, OfficePoint actual) {
        Assert.InRange(Math.Abs(expected.X - actual.X), 0, 0.1);
        Assert.InRange(Math.Abs(expected.Y - actual.Y), 0, 0.1);
    }
}
