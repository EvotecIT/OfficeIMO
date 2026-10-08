using System;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentOdgSizedAttachmentTests {
    [Theory]
    [InlineData("explicit")]
    [InlineData("built-in")]
    [InlineData("automatic")]
    [InlineData("relative")]
    [InlineData("grown")]
    public void ConnectorBeforeItsSizedTargetUsesProjectedBoundsAndLeavesEditingCoordinatesUnchanged(string profile) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var connector = page.Shapes.AddConnector(new OfficePoint(140, 54), new OfficePoint(300, 42));
        var target = page.Shapes.AddTextBox(Rect(20, 30, 120, 48), profile == "grown" ?
            string.Join(" ", Enumerable.Repeat("Alpha beta gamma delta", 12)) : "", "Sized target");
        target.FillColor = OdfColor.Parse("#e0f0ff"); target.StrokeColor = OdfColor.Parse("#204060");
        target.AutoGrowWidth = false; target.AutoGrowHeight = profile == "grown";
        if (profile != "grown") {
            target.TextBoxMinimumWidth = OdfLength.Points(60); target.TextBoxMinimumHeight = OdfLength.Points(24);
        } else {
            target.Paragraphs[0].FontSize = OdfLength.Points(12); target.WrapText = true;
            target.TextVerticalAlignment = OdfTextAreaVerticalAlignment.Top;
        }
        var other = page.Shapes.AddRectangle(Rect(300, 30, 40, 24), "Other");
        connector.AttachEnd(other.AddGluePoint(OdgGluePointAlignment.Left));
        if (profile is "automatic" or "built-in") {
            connector.AttachStartToShape(target);
            if (profile == "built-in") {
                connector.Element.SetAttributeValue(OdfNamespaces.Draw + "start-glue-point", "1"); document.MarkPartDirty("content.xml");
            }
        } else connector.AttachStart(target.AddGluePoint(profile == "relative" ? OdgGluePointAlignment.Bottom : OdgGluePointAlignment.Right));
        var savedStart = new OfficePoint(connector.X1.ToPoints(), connector.Y1.ToPoints());
        string[] before = XmlState(document);
        var result = page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
        var frame = Assert.Single(result.Value.Elements.OfType<OfficeDrawingShape>(), s => s.Shape.Kind == OfficeShapeKind.Rectangle && s.Shape.Width != 40);
        var route = Assert.Single(result.Value.Elements.OfType<OfficeDrawingShape>(), s => s.Shape.Kind == OfficeShapeKind.Line);
        double x = frame.X + frame.Shape.Width * (profile == "relative" ? .5 : 1);
        double y = frame.Y + frame.Shape.Height * (profile == "relative" ? 1 : .5);
        Assert.Equal(x, route.X + route.Shape.Points[0].X, 8); Assert.Equal(y, route.Y + route.Shape.Points[0].Y, 8);
        Assert.Equal(new OfficePoint(connector.X1.ToPoints(), connector.Y1.ToPoints()), savedStart);
        Assert.Equal(before, XmlState(document));
    }

    [Fact]
    public void ConnectorOutsideTransformedGroupUsesTheSizedChildInPageCoordinates() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var connector = page.Shapes.AddConnector(new OfficePoint(140, 54), new OfficePoint(300, 42));
        var group = page.Shapes.AddGroup();
        var target = group.Children.AddTextBox(Rect(20, 30, 120, 48), "", "Child");
        target.Transform = "translate(1cm 0cm)";
        target.AutoGrowHeight = false; target.AutoGrowWidth = false;
        target.TextBoxMinimumWidth = OdfLength.Points(60); target.TextBoxMinimumHeight = OdfLength.Points(24);
        connector.AttachStart(target.AddGluePoint(OdgGluePointAlignment.Right));
        string[] before = XmlState(document); var result = page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
        var route = Assert.Single(result.Value.Elements.OfType<OfficeDrawingShape>(), s => s.Shape.Kind == OfficeShapeKind.Line);
        // Connector coordinates use the existing three-decimal point lexical
        // codec; the leaf transform retains its full unit-conversion precision.
        Assert.InRange(Math.Abs(80 + OdfLength.Centimeters(1).ToPoints() - route.X - route.Shape.Points[0].X), 0, .001);
        Assert.Equal(42, route.Y + route.Shape.Points[0].Y, 8);
        Assert.Equal(before, XmlState(document));
    }

    [Fact]
    public void SavedRouteThatNoLongerMatchesSizedAttachmentsReportsLossInsteadOfSilentlyDetaching() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var target = page.Shapes.AddTextBox(Rect(20, 30, 120, 48), "", "Target");
        target.AutoGrowHeight = false; target.AutoGrowWidth = false;
        target.TextBoxMinimumWidth = OdfLength.Points(60); target.TextBoxMinimumHeight = OdfLength.Points(24);
        var connector = page.Shapes.AddConnector(new OfficePoint(140, 54), new OfficePoint(300, 42));
        connector.AttachStart(target.AddGluePoint(OdgGluePointAlignment.Right));
        connector.SetConnectorRoute(new[] { OfficePathCommand.MoveTo(140, 54), OfficePathCommand.LineTo(300, 42) });
        string[] before = XmlState(document); var result = page.ToDrawing();
        Assert.Contains(result.Report.Mappings, m => m.Status == OdfConversionMappingStatus.Skipped &&
            m.Feature == "shape:" + connector.Name);
        Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        Assert.Equal(before, XmlState(document));
    }

    private static OdfRect Rect(double x, double y, double width, double height) => new OdfRect(
        OdfLength.Points(x), OdfLength.Points(y), OdfLength.Points(width), OdfLength.Points(height));
    private static string[] XmlState(OdfDocument document) => new[] { document.GetXml("content.xml").ToString(), document.GetXml("styles.xml").ToString() };
}
