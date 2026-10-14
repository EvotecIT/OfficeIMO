using System;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed partial class OpenDocumentOdgLineLabelTests {
    [Theory]
    [InlineData("cubic", 9000, 12500)]
    [InlineData("quadratic", 9000, 12500)]
    [InlineData("reversed", 9000, 12500)]
    [InlineData("upper", 9000, 11500)]
    [InlineData("mixed", 9000, 13875)]
    [InlineData("asymmetric", 9000, 12231.639838222847)]
    [InlineData("sideways", 12500, 9000)]
    public void CurvedCaptionsUseTheTrueRouteBoundsAcrossPackageAndFlatReopening(string profile, double nativeX, double nativeY) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var connector = page.Shapes.AddConnector(new OfficePoint(0, 0), new OfficePoint(1, 1), OdgConnectorKind.Curve);
        connector.SetConnectorRoute(Curve(profile)); AddLabel(document, connector);
        string before = document.GetXml("content.xml").ToString();
        foreach (var read in RoundTrips(document)) {
            var result = read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
            var frame = Assert.Single(Frames(result.Value));
            Assert.Equal(nativeX * NativePoint, Center(frame).X, 6);
            Assert.Equal(nativeY * NativePoint, Center(frame).Y, 6);
            Assert.Equal(1, frame.Transform.M11, 6); Assert.Equal(0, frame.Transform.M12, 6);
            Assert.Equal("Primary emphasized\nSecond line", frame.Text.PlainText);
            var run = frame.Text.Paragraphs[0].Runs.Single(r => r.Text == "emphasized");
            Assert.True(run.Bold); Assert.Equal(OfficeColor.Parse("#BA1234"), run.Color);
            Assert.Contains("Second line", OfficeDrawingSvgExporter.ToSvg(result.Value));
        }
        Assert.Equal(before, document.GetXml("content.xml").ToString());
    }

    [Theory]
    [InlineData("shape")]
    [InlineData("bake")]
    [InlineData("group-bake")]
    public void CurvedCaptionTranslationsUseTheSameAnchorBeforeAndAfterBaking(string operation) {
        var document = LoadProducer(); var page = document.Pages[0];
        var original = page.Shapes.Single(s => s.XmlId == "id13");
        original.AttachStartToShape(null); original.ConnectorKind = OdgConnectorKind.Curve;
        original.SetConnectorRoute(Curve("cubic"));
        for (int i = page.Shapes.Count - 1; i >= 0; i--) if (page.Shapes[i].XmlId != "id13") page.Shapes.RemoveAt(i);
        Graphic(document, original).SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "wrap-option", "no-wrap");
        if (operation.StartsWith("group", StringComparison.Ordinal)) {
            var group = page.Shapes.AddGroup();
            // Retain the producer's paragraphs and their named style declarations.
            group.Element.Add(new System.Xml.Linq.XElement(original.Element));
            page.Shapes.RemoveAt(0);
            string styles = document.GetXml("styles.xml").ToString();
            group.TransformChildren("translate(20pt 30pt)");
            Assert.Equal(styles, document.GetXml("styles.xml").ToString());
        } else {
            string styles = document.GetXml("styles.xml").ToString();
            original.Transform = "translate(20pt 30pt)";
            if (operation == "bake") original.BakeConnectorTransform();
            Assert.Equal(styles, document.GetXml("styles.xml").ToString());
        }
        foreach (var read in RoundTrips(document)) {
            var result = read.Pages[0].ToDrawing();
            var unsupported = result.Report.Mappings.Where(m => m.Status == OdfConversionMappingStatus.Unsupported).ToArray();
            Assert.NotEmpty(unsupported);
            Assert.All(unsupported, m => Assert.EndsWith(":text-window-color", m.Feature));
            Assert.DoesNotContain(result.Report.Mappings, m => m.Status == OdfConversionMappingStatus.Skipped);
            Assert.Throws<OdfConversionLossException>(() => read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
            var frame = Assert.Single(Frames(result.Value));
            Assert.InRange(Center(frame).X, 9000 * NativePoint + 19.98, 9000 * NativePoint + 20.02);
            Assert.InRange(Center(frame).Y, 12500 * NativePoint + 29.98, 12500 * NativePoint + 30.02);
            Assert.Equal("95.217.96.40/29\n95.216.28.178", frame.Text.PlainText);
            Assert.All(frame.Text.Paragraphs.SelectMany(p => p.Runs), r => Assert.Equal(10, r.FontSize));
        }
    }

    private const double NativePoint = 72D / 2540;

    private static OfficePathCommand[] Curve(string profile) => (profile switch {
        "quadratic" => new[] { OfficePathCommand.MoveTo(3000, 8000), OfficePathCommand.QuadraticBezierTo(9000, 26000, 15000, 8000) },
        "reversed" => new[] { OfficePathCommand.MoveTo(15000, 8000), OfficePathCommand.CubicBezierTo(15000, 20000, 3000, 20000, 3000, 8000) },
        "upper" => new[] { OfficePathCommand.MoveTo(3000, 16000), OfficePathCommand.CubicBezierTo(3000, 4000, 15000, 4000, 15000, 16000) },
        "mixed" => new[] { OfficePathCommand.MoveTo(3000, 8000), OfficePathCommand.LineTo(3000, 13000),
            OfficePathCommand.CubicBezierTo(3000, 22000, 15000, 22000, 15000, 13000), OfficePathCommand.LineTo(15000, 8000) },
        "asymmetric" => new[] { OfficePathCommand.MoveTo(3000, 8000), OfficePathCommand.CubicBezierTo(3000, 26000, 15000, 10000, 15000, 8000) },
        "sideways" => new[] { OfficePathCommand.MoveTo(8000, 3000), OfficePathCommand.CubicBezierTo(20000, 3000, 20000, 15000, 8000, 15000) },
        _ => new[] { OfficePathCommand.MoveTo(3000, 8000), OfficePathCommand.CubicBezierTo(3000, 20000, 15000, 20000, 15000, 8000) }
    }).Select(command => command.Kind switch {
        OfficePathCommandKind.MoveTo => OfficePathCommand.MoveTo(Native(command.Point)),
        OfficePathCommandKind.LineTo => OfficePathCommand.LineTo(Native(command.Point)),
        OfficePathCommandKind.QuadraticBezierTo => OfficePathCommand.QuadraticBezierTo(Native(command.ControlPoint1), Native(command.Point)),
        _ => OfficePathCommand.CubicBezierTo(Native(command.ControlPoint1), Native(command.ControlPoint2), Native(command.Point))
    }).ToArray();

    private static OfficePoint Native(OfficePoint point) => new(point.X * NativePoint, point.Y * NativePoint);
}
