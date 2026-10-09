using System.Globalization;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioConnectorTextFrameTests {
    private static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";
    private static readonly XNamespace Svg = "http://www.w3.org/2000/svg";

    [Theory]
    [InlineData("vdx", 1, 0)]
    [InlineData("vsdx", 1, 0)]
    [InlineData("vdx", 1, 90)]
    [InlineData("vsdx", 1, 90)]
    [InlineData("vdx", 1.5, -30)]
    [InlineData("vsdx", 1.5, -30)]
    public void ProducerCalloutFrameFollowsItsNativeConnectorAndSurvivesFurtherEdits(string format, double scale, double degrees) {
        var document = Producer(); var connector = Callout(document);
        var before = ExportFrame(document); var start = connector.StartPoint; var end = connector.EndPoint;
        VisioTextStyle style = connector.TextStyle!;
        XElement text = new(CalloutElement(document).Element(Legacy + "Text")!);
        double angle = degrees * Math.PI / 180;
        OfficePoint nextStart = new(start.X + .5, start.Y - .75);
        OfficePoint Transform(OfficePoint point) {
            double x = (point.X - start.X) * scale, y = (point.Y - start.Y) * scale;
            return new OfficePoint(nextStart.X + x * Math.Cos(angle) - y * Math.Sin(angle), nextStart.Y + x * Math.Sin(angle) + y * Math.Cos(angle));
        }
        OfficePoint expected = Transform(new OfficePoint(before.X, before.Y));
        connector.StartPoint = nextStart; connector.EndPoint = Transform(end);
        Assert.Same(style, connector.TextStyle);
        foreach (var candidate in new[] { document, Reopen(document, format) }) {
            var actual = ExportFrame(candidate);
            Assert.Equal(expected.X, actual.X, 8); Assert.Equal(expected.Y, actual.Y, 8);
            Assert.Equal(before.Width * scale, actual.Width, 8); Assert.Equal(before.Height * scale, actual.Height, 8);
            Assert.Equal(before.Angle + angle, actual.Angle, 8);
            Assert.Equal(style.Size, Callout(candidate).TextStyle!.Size);
            Assert.True(XNode.DeepEquals(text, CalloutElement(candidate).Element(Legacy + "Text")));
        }
        document = Reopen(document, format); connector = Callout(document);
        before = ExportFrame(document); connector.StartPoint = new(connector.StartPoint.X - .25, connector.StartPoint.Y + .125);
        connector.EndPoint = new(connector.EndPoint.X - .25, connector.EndPoint.Y + .125);
        var again = ExportFrame(Reopen(document, format));
        Assert.Equal(before.X - .25, again.X, 8); Assert.Equal(before.Y + .125, again.Y, 8);
    }

    [Theory]
    [InlineData("vdx")]
    [InlineData("vsdx")]
    public void ProducerTranslationMovesPaintedTextWithoutChangingRuns(string format) {
        var document = Producer(); var connector = Callout(document);
        XElement[] before = PaintedText(document);
        connector.StartPoint = new(connector.StartPoint.X + .5, connector.StartPoint.Y - .75);
        connector.EndPoint = new(connector.EndPoint.X + .5, connector.EndPoint.Y - .75);
        foreach (var candidate in new[] { document, Reopen(document, format) }) {
            XElement[] after = PaintedText(candidate); Assert.Equal(before.Length, after.Length);
            for (int i = 0; i < before.Length; i++) {
                Assert.Equal(before[i].Value, after[i].Value);
                Assert.Equal(Number(before[i], "x") + 36, Number(after[i], "x"), 3);
                Assert.Equal(Number(before[i], "y") + 54, Number(after[i], "y"), 3);
                Assert.Equal((string?)before[i].Attribute("font-weight"), (string?)after[i].Attribute("font-weight"));
            }
        }
    }

    [Theory]
    [InlineData("vdx", false)]
    [InlineData("vsdx", false)]
    [InlineData("vdx", true)]
    [InlineData("vsdx", true)]
    public void ExplicitPagePinsOverrideNativeBindingEvenAfterReopening(string format, bool setters) {
        var document = Producer(); var connector = Callout(document);
        if (setters) { connector.LabelPlacement!.PinX = 5; connector.LabelPlacement.PinY = 8; }
        else connector.PlaceLabelAt(5, 8, 2, .5);
        connector.TextStyle!.TextAngle = .2;
        document = Reopen(document, format); connector = Callout(document); var before = ExportFrame(document);
        connector.StartPoint = new(connector.StartPoint.X + 1, connector.StartPoint.Y + 1);
        connector.EndPoint = new(connector.EndPoint.X + 1, connector.EndPoint.Y + 1);
        var after = ExportFrame(Reopen(document, format));
        Assert.Equal(before.X, after.X, 8); Assert.Equal(before.Y, after.Y, 8); Assert.Equal(.2, after.Angle, 8);
    }

    [Theory]
    [InlineData("vdx")]
    [InlineData("vsdx")]
    public void PathPositionAndPageOffsetsRemainEditableAfterReopening(string format) {
        var document = VisioDocument.Create(); var page = document.AddPage("Path", 8, 8);
        var connector = page.AddConnector("c", new OfficePoint(1, 2), new OfficePoint(5, 2));
        connector.Label = "Review"; connector.PlaceLabel(.25, .1, .2, 1, .3);
        document = Reopen(document, format); connector = document.Pages[0].Connectors[0];
        Assert.Equal(.25, connector.LabelPlacement!.Position); Assert.Equal(.1, connector.LabelPlacement.OffsetX); Assert.Equal(.2, connector.LabelPlacement.OffsetY);
        connector.EndPoint = new OfficePoint(5, 6);
        foreach (var candidate in new[] { document, Reopen(document, format) }) {
            var saved = candidate.Pages[0].Connectors[0];
            var point = VisioConnectorLabelFrame.ResolvePlacement(saved)!;
            Assert.Null(point.PinX); Assert.Null(point.PinY);
            var snapshot = candidate.CreateInspectionSnapshot().Pages[0].Connectors[0];
            Assert.Equal(2.1, snapshot.LabelResolvedPinX!.Value, 8); Assert.Equal(3.2, snapshot.LabelResolvedPinY!.Value, 8);
            Assert.DoesNotContain("OfficeIMOConnectorLabelPlacement", saved.Data.Keys);
            Assert.DoesNotContain(saved.ShapeData, row => row.Name == "OfficeIMOConnectorLabelPlacement");
        }
    }

    [Fact]
    public void UnknownReservedRowRemainsDataAndCannotBeOverwrittenByPlacementMetadata() {
        var document = Producer(); var connector = Callout(document);
        connector.Data["OfficeIMOConnectorLabelPlacement"] = "future-version";
        Assert.Equal("future-version", Callout(Reopen(document, "vdx")).Data["OfficeIMOConnectorLabelPlacement"]);
        connector.PlaceLabelAt(5, 8);
        Assert.Throws<NotSupportedException>(() => document.ToLegacyXmlResult());
        Assert.Equal("future-version", connector.Data["OfficeIMOConnectorLabelPlacement"]);
    }

    [Theory]
    [InlineData("vdx")]
    [InlineData("vsdx")]
    public void ExplicitAngleEditsRetainTheirPageAngleWithoutFreezingTheNativePin(string format) {
        var document = Producer(); var connector = Callout(document);
        connector.TextStyle!.TextAngle = .2;
        document = Reopen(document, format); connector = Callout(document);
        var before = ExportFrame(document); var start = connector.StartPoint; var end = connector.EndPoint;
        connector.EndPoint = new(start.X - (end.Y - start.Y), start.Y + (end.X - start.X));
        var after = ExportFrame(Reopen(document, format));
        Assert.Equal(.2, after.Angle, 8);
        Assert.Equal(start.X - (before.Y - start.Y), after.X, 8);
        Assert.Equal(start.Y + (before.X - start.X), after.Y, 8);
    }

    [Theory]
    [InlineData("vdx")]
    [InlineData("vsdx")]
    public void CopiedEditedNativeFramesRemainIndependentAndFollowTheirOwnEndpoints(string format) {
        var document = Producer(); var source = document.Pages.Single(p => p.Connectors.Any(c => c.Label?.StartsWith("This implication") == true));
        var original = Callout(document);
        original.StartPoint = new(original.StartPoint.X + .25, original.StartPoint.Y - .5);
        original.EndPoint = new(original.EndPoint.X + .25, original.EndPoint.Y - .5);
        var before = ExportFrame(document, source.Name);
        var copy = document.DuplicatePage(source, "Independent callout");
        var connector = copy.Connectors.Single(c => c.Label?.StartsWith("This implication") == true);
        Assert.NotSame(original.LabelPlacement, connector.LabelPlacement); Assert.NotSame(original.TextStyle, connector.TextStyle);
        connector.StartPoint = new(connector.StartPoint.X + .5, connector.StartPoint.Y + .125);
        connector.EndPoint = new(connector.EndPoint.X + .5, connector.EndPoint.Y + .125);
        foreach (var candidate in new[] { document, Reopen(document, format) }) {
            var retained = ExportFrame(candidate, source.Name);
            Assert.Equal(before.X, retained.X, 8); Assert.Equal(before.Y, retained.Y, 8);
            Assert.Equal(before.Width, retained.Width, 8); Assert.Equal(before.Height, retained.Height, 8);
            Assert.Equal(before.Angle, retained.Angle, 8);
            var moved = ExportFrame(candidate, copy.Name);
            Assert.Equal(before.X + .5, moved.X, 8); Assert.Equal(before.Y + .125, moved.Y, 8);
        }
    }

    [Theory]
    [InlineData("vdx")]
    [InlineData("vsdx")]
    public void ClearingPagePinsRestoresEditablePathIntent(string format) {
        var document = VisioDocument.Create(); var page = document.AddPage("Path", 8, 8);
        var connector = page.AddConnector("c", new OfficePoint(1, 2), new OfficePoint(5, 2));
        connector.Label = "Review"; connector.PlaceLabelAt(3, 4);
        connector.LabelPlacement!.PinX = null; connector.LabelPlacement.PinY = null;
        connector.LabelPlacement.Position = .25; connector.LabelPlacement.OffsetX = .1; connector.LabelPlacement.OffsetY = .2;
        document = Reopen(document, format); connector = document.Pages[0].Connectors[0];
        connector.EndPoint = new OfficePoint(5, 6);
        var snapshot = Reopen(document, format).CreateInspectionSnapshot().Pages[0].Connectors[0];
        Assert.Equal(2.1, snapshot.LabelResolvedPinX!.Value, 8); Assert.Equal(3.2, snapshot.LabelResolvedPinY!.Value, 8);
    }

    [Theory]
    [InlineData("vdx")]
    [InlineData("vsdx")]
    public void FittingAnAlreadyScaledNativeFrameUsesCurrentUnitsAndRebasesFurtherScaling(string format) {
        var document = Producer(); var connector = Callout(document);
        var start = connector.StartPoint; var end = connector.EndPoint;
        connector.EndPoint = new(start.X + 2 * (end.X - start.X), start.Y + 2 * (end.Y - start.Y));
        connector.ResizeLabelToText(maximumWidth: 1.4);
        double width = connector.LabelPlacement!.Width, height = connector.LabelPlacement.Height;
        Assert.InRange(width, .45, 1.4);
        foreach (var candidate in new[] { document, Reopen(document, format) }) {
            var frame = ExportFrame(candidate);
            Assert.Equal(width, frame.Width, 8); Assert.Equal(height, frame.Height, 8);
            var current = Callout(candidate); var currentStart = current.StartPoint; var currentEnd = current.EndPoint;
            current.EndPoint = new(currentStart.X + 1.5 * (currentEnd.X - currentStart.X), currentStart.Y + 1.5 * (currentEnd.Y - currentStart.Y));
            var scaled = ExportFrame(Reopen(candidate, format));
            Assert.Equal(width * 1.5, scaled.Width, 8); Assert.Equal(height * 1.5, scaled.Height, 8);
            Assert.Equal(currentStart.X + 1.5 * (frame.X - currentStart.X), scaled.X, 8);
            Assert.Equal(currentStart.Y + 1.5 * (frame.Y - currentStart.Y), scaled.Y, 8);
        }
    }

    [Fact]
    public void FittingACollapsedNativeFrameRequiresExplicitPlacementWithoutCorruptingDimensions() {
        var document = Producer(); var connector = Callout(document);
        connector.EndPoint = connector.StartPoint;
        double width = connector.LabelPlacement!.Width, height = connector.LabelPlacement.Height;
        Assert.Throws<NotSupportedException>(() => connector.ResizeLabelToText(maximumWidth: 1.4));
        Assert.Equal(width, connector.LabelPlacement.Width); Assert.Equal(height, connector.LabelPlacement.Height);
        connector.PlaceLabelAt(3, 4); connector.ResizeLabelToText(maximumWidth: 1.4);
        Assert.InRange(connector.LabelPlacement!.Width, .45, 1.4);
    }

    [Theory]
    [InlineData("vdx", false)]
    [InlineData("vsdx", false)]
    [InlineData("vdx", true)]
    [InlineData("vsdx", true)]
    public void UnsuccessfulCleanupRetainsNativeBindingAndResolvedAngle(string format, bool globalCleanup) {
        var document = Producer(); var connector = Callout(document);
        var page = document.Pages.Single(p => p.Connectors.Contains(connector));
        var start = connector.StartPoint; var end = connector.EndPoint;
        connector.EndPoint = new(start.X - (end.Y - start.Y), start.Y + (end.X - start.X));
        AddLabelObstacle(page, connector);
        var before = ExportFrame(document); var placement = connector.LabelPlacement;
        page.ResolveConnectorLabelOverlaps(maxAttempts: 0, maxPositionShifts: 0, avoidLabels: globalCleanup, avoidConnectorPaths: false);
        Assert.Same(placement, connector.LabelPlacement);
        Assert.Equal(before.Angle, ExportFrame(document).Angle, 8);
        connector.StartPoint = new(connector.StartPoint.X + .5, connector.StartPoint.Y - .75);
        connector.EndPoint = new(connector.EndPoint.X + .5, connector.EndPoint.Y - .75);
        var after = ExportFrame(Reopen(document, format));
        Assert.Equal(before.X + .5, after.X, 8); Assert.Equal(before.Y - .75, after.Y, 8);
        Assert.Equal(before.Angle, after.Angle, 8);
    }

    [Theory]
    [InlineData("vdx")]
    [InlineData("vsdx")]
    public void SuccessfulCleanupRetainsNativeBindingAndResolvedAngle(string format) {
        var document = Producer(); var connector = Callout(document);
        var page = document.Pages.Single(p => p.Connectors.Contains(connector));
        var start = connector.StartPoint; var end = connector.EndPoint;
        connector.EndPoint = new(start.X - (end.Y - start.Y), start.Y + (end.X - start.X));
        AddLabelObstacle(page, connector);
        var before = ExportFrame(document);
        page.ResolveConnectorLabelOverlaps(avoidLabels: false, avoidConnectorPaths: false);
        var moved = ExportFrame(document);
        Assert.True(Math.Abs(moved.X - before.X) > 1e-6 || Math.Abs(moved.Y - before.Y) > 1e-6);
        Assert.Equal(before.Angle, moved.Angle, 8);
        connector.StartPoint = new(connector.StartPoint.X + .5, connector.StartPoint.Y - .75);
        connector.EndPoint = new(connector.EndPoint.X + .5, connector.EndPoint.Y - .75);
        var after = ExportFrame(Reopen(document, format));
        Assert.Equal(moved.X + .5, after.X, 8); Assert.Equal(moved.Y - .75, after.Y, 8);
        Assert.Equal(moved.Angle, after.Angle, 8);
    }

    private static void AddLabelObstacle(VisioPage page, VisioConnector connector) {
        var bounds = connector.GetLabelBounds();
        page.AddRectangle((bounds.Left + bounds.Right) / 2, (bounds.Bottom + bounds.Top) / 2, .5, .5, "Obstacle", VisioMeasurementUnit.Inches);
    }

    private static VisioDocument Producer() => VisioDocument.LoadLegacyXml(Path.Combine(AppContext.BaseDirectory, "Fixtures", "LegacyXml", "nxbre-chocolatebox.vdx")).Value;
    private static VisioConnector Callout(VisioDocument document) => document.Pages.SelectMany(page => page.Connectors).Single(c => c.Label?.StartsWith("This implication") == true);
    private static VisioDocument Reopen(VisioDocument document, string format) => format == "vdx"
        ? VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value
        : VisioDocument.Load(new MemoryStream(document.ToBytes()));
    private static bool IsCallout(XElement shape) => shape.Element(Legacy + "Text")?.Value.StartsWith("This implication") == true;
    private static XElement CalloutElement(VisioDocument document, string? pageName = null) => XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value))
        .Descendants(Legacy + "Shape").Single(shape => IsCallout(shape) && (pageName == null || (string?)shape.Ancestors(Legacy + "Page").Single().Attribute("Name") == pageName));
    private static (double X, double Y, double Width, double Height, double Angle) ExportFrame(VisioDocument document, string? pageName = null) {
        XElement shape = CalloutElement(document, pageName), frame = shape.Element(Legacy + "XForm")!, text = shape.Element(Legacy + "TextXForm")!;
        double Get(XElement group, string name) => double.Parse(group.Element(Legacy + name)!.Value, CultureInfo.InvariantCulture);
        double angle = Get(frame, "Angle"), x = Get(text, "TxtPinX") - Get(frame, "LocPinX"), y = Get(text, "TxtPinY") - Get(frame, "LocPinY");
        return (Get(frame, "PinX") + x * Math.Cos(angle) - y * Math.Sin(angle), Get(frame, "PinY") + x * Math.Sin(angle) + y * Math.Cos(angle),
            Get(text, "TxtWidth"), Get(text, "TxtHeight"), angle + Get(text, "TxtAngle"));
    }
    private static XElement[] PaintedText(VisioDocument document) {
        var page = document.Pages.Single(p => p.Connectors.Any(c => c.Label?.StartsWith("This implication") == true));
        return XDocument.Parse(page.ToSvg(new VisioSvgSaveOptions { PixelsPerInch = 72, ResolveConnectorLabelOverlaps = false, RenderStencilArtwork = false }))
            .Descendants(Svg + "g").Single(g => (string?)g.Attribute("data-visio-connector-id") == "18").Descendants(Svg + "text").ToArray();
    }
    private static double Number(XElement element, string attribute) => double.Parse(element.Attribute(attribute)!.Value, CultureInfo.InvariantCulture);
}
