using System;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed partial class OpenDocumentOdgLineLabelTests {
    [Theory]
    [InlineData("horizontal", false)]
    [InlineData("horizontal", true)]
    [InlineData("vertical", false)]
    [InlineData("vertical", true)]
    [InlineData("diagonal", false)]
    [InlineData("diagonal", true)]
    [InlineData("curve", false)]
    [InlineData("bent", false)]
    public void DeclaredWrappingRetainsTheCompleteUnwrappedCaptionAndAnExplicitLoss(string route, bool line) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        OfficePoint start = route == "vertical" ? new(200, 40) : new(100, 100);
        OfficePoint end = route switch { "vertical" => new(200, 180), "diagonal" => new(180, 200), _ => new(180, 100) };
        OdgShape shape = line ? page.Shapes.AddLine(OdfLength.Points(start.X), OdfLength.Points(start.Y),
            OdfLength.Points(end.X), OdfLength.Points(end.Y)) : page.Shapes.AddConnector(start, end,
                route == "curve" ? OdgConnectorKind.Curve : route == "bent" ? OdgConnectorKind.Lines : OdgConnectorKind.Line);
        if (route == "curve") shape.SetConnectorRoute(new[] { OfficePathCommand.MoveTo(start),
            OfficePathCommand.CubicBezierTo(100, 260, 180, 260, end.X, end.Y) });
        if (route == "bent") shape.SetConnectorRoute(new[] { OfficePathCommand.MoveTo(start), OfficePathCommand.LineTo(100, 240),
            OfficePathCommand.LineTo(180, 240), OfficePathCommand.LineTo(end) });
        shape.Name = "Caption"; AddLabel(document, shape);
        shape.Paragraphs[0].Text = "Alpha Beta Gamma Delta Epsilon Zeta Eta Theta Iota Kappa Lambda Mu";
        Graphic(document, shape).SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "wrap-option", "wrap");
        string before = document.GetXml("content.xml").ToString();
        foreach (var read in RoundTrips(document)) {
            var result = read.Pages[0].ToDrawing();
            var frame = Assert.Single(Frames(result.Value));
            Assert.Equal("Alpha Beta Gamma Delta Epsilon Zeta Eta Theta Iota Kappa Lambda Mu\nSecond line", frame.Text.PlainText);
            Assert.False(frame.Text.WrapText);
            Assert.Contains(result.Report.Mappings, m => m.Feature == "shape:Caption:text:label-wrapping" &&
                m.Status == OdfConversionMappingStatus.Unsupported && m.Message!.Contains("unwrapped"));
            Assert.DoesNotContain(result.Report.Mappings, m => m.Feature == "shape:Caption:text" && m.Status == OdfConversionMappingStatus.Skipped);
            Assert.Throws<OdfConversionLossException>(() => read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
            Assert.Equal((start.X + end.X) / 2, Center(frame).X, 2);
            Assert.InRange(Center(frame).Y, (route switch { "curve" => 160D, "bent" => 170D, _ => (start.Y + end.Y) / 2 }) - .02,
                (route switch { "curve" => 160D, "bent" => 170D, _ => (start.Y + end.Y) / 2 }) + .02);
        }
        Assert.Equal(before, document.GetXml("content.xml").ToString());
    }

    [Fact]
    public void InheritedWrappingDeclarationsRetainProducerSpansAndHardBreaksWithoutMutatingStyles() {
        var document = LoadProducer(); var connector = document.Pages[0].Shapes.Single(s => s.XmlId == "id13");
        var style = Graphic(document, connector); style.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "wrap-option", null);
        var parent = document.Styles.CreateNamed("WrappedCaption", OdfStyleFamily.Graphic);
        parent.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "wrap-option", "wrap");
        style.ParentStyleName = parent.Name;
        connector.Paragraphs[0].AddRun("\n"); connector.Paragraphs[0].AddRun("Cached annotation").Bold = true;
        string before = document.GetXml("content.xml").ToString(), styles = document.GetXml("styles.xml").ToString();
        foreach (var read in RoundTrips(document)) {
            var result = read.Pages[0].ToDrawing();
            var frame = Assert.Single(Frames(result.Value), f => f.Text.PlainText.StartsWith("95.217.96.40/29", StringComparison.Ordinal));
            Assert.Equal("95.217.96.40/29\nCached annotation\n95.216.28.178", frame.Text.PlainText);
            Assert.Equal("Cached annotation", string.Concat(frame.Text.Paragraphs[0].Runs.Where(r => r.Bold).Select(r => r.Text)).Trim());
            Assert.Contains(result.Report.Mappings, m => m.Feature.EndsWith(":label-wrapping", StringComparison.Ordinal));
        }
        Assert.Equal(before, document.GetXml("content.xml").ToString()); Assert.Equal(styles, document.GetXml("styles.xml").ToString());
    }
}
