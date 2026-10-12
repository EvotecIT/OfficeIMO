using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;
using OfficeIMO.Visio.Stencils;
using Xunit;

namespace OfficeIMO.Tests;

public class VisioDiagramPageBackgroundTests {
    [Theory]
    [InlineData("Technical")]
    [InlineData("Modern")]
    [InlineData("Office")]
    [InlineData("Fluent")]
    [InlineData("Dark")]
    public void StencilCaptionUsesPersistedPageTextColorInsteadOfNodeFillText(string preset) {
        var theme = preset switch {
            "Modern" => VisioStyleTheme.Modern(),
            "Office" => VisioStyleTheme.Office(),
            "Fluent" => VisioStyleTheme.Fluent(),
            _ => VisioStyleTheme.Technical()
        };
        if (preset == "Dark") {
            theme.PageBackgroundColor = OfficeColor.FromRgb(17, 19, 23);
            theme.LegendText.Color = OfficeColor.FromRgb(230, 235, 240);
        }
        var expectedCaptionColor = theme.LegendText.Color;
        var expectedPageColor = theme.PageBackgroundColor ?? OfficeColor.White;
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
        try {
            var document = VisioDocument.Create(path).GraphDiagram("Service", graph => graph
                .Theme(theme).StencilNode("server", "Service host", VisioStencils.Network.Get("server")));
            document.Save();
            Assert.Empty(VisioValidator.Validate(path));
            var page = Assert.Single(VisioDocument.Load(path).Pages);
            var caption = page.Shapes.Single(shape => shape.Id == "server-label");
            Assert.Equal("Service host", caption.Text);
            Assert.Equal(expectedCaptionColor, caption.TextStyle!.Color);
            Assert.Equal(expectedPageColor, caption.TextStyle.BackgroundColor);
            Assert.NotEqual(expectedPageColor, caption.TextStyle.Color);
            Assert.Equal(OfficeColor.White, theme.Primary.TextStyle!.Color);
            XNamespace ns = "http://www.w3.org/2000/svg";
            var svg = XDocument.Parse(page.ToSvg());
            var text = svg.Descendants(ns + "text").Single(element => element.Value == "Service host");
            Assert.Equal(expectedCaptionColor!.Value.ToString(), (string?)text.Attribute("fill"), ignoreCase: true);
        } finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void SequenceThemeFillAndExplicitMessageBackgroundSurviveSave() {
        var background = OfficeColor.FromRgb(17, 19, 23);
        var explicitLabel = OfficeColor.FromRgb(60, 70, 80);
        var theme = VisioStyleTheme.Minimal();
        theme.PageBackgroundColor = background;
        theme.Connector.TextStyle!.Color = OfficeColor.White;
        theme.ControlConnector.TextStyle!.BackgroundColor = explicitLabel;
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
        try {
            var document = VisioDocument.Create(path).SequenceDiagram("Requests", sequence => sequence
                .Theme(theme).Title().Participant("client", "Client").Participant("service", "Service")
                .Call("client", "service", "Request", "request")
                .Return("service", "client", "Accepted", "accepted"));
            document.Save();
            Assert.Empty(VisioValidator.Validate(path));
            var page = Assert.Single(VisioDocument.Load(path).Pages);
            Assert.Equal(background, page.Shapes[0].FillColor);
            Assert.Equal(page.Width, page.Shapes[0].Width, 6);
            Assert.Equal(page.Height, page.Shapes[0].Height, 6);
            Assert.True(page.Shapes[0].IsBackgroundSurface);
            Assert.Contains(page.Shapes, shape => shape.Id == "client" && shape.Text == "Client");
            Assert.Contains(page.Shapes, shape => shape.Id == "service" && shape.Text == "Service");
            Assert.Equal(background, page.Connectors.Single(item => item.Id == "request").TextStyle!.BackgroundColor);
            Assert.Equal(explicitLabel, page.Connectors.Single(item => item.Id == "accepted").TextStyle!.BackgroundColor);
        } finally { if (File.Exists(path)) File.Delete(path); }
    }
}
