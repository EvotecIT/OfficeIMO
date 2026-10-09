using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.OpenDocument;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class OpenDocumentOdgTextChainProjectionTests {
    [Theory]
    [InlineData("ordinary")]
    [InlineData("nested")]
    [InlineData("master")]
    [InlineData("other-page")]
    public void LinkedSourcesAndEmptyIncomingTerminalsReportFlowWithoutGrowth(string placement) {
        var document = OdgDocument.Create();
        var page = document.AddPage("Terminal page");
        var terminal = Box(page.Shapes, "Terminal", string.Empty);
        OdgShapes sources = placement switch {
            "nested" => page.Shapes.AddGroup("Source group").Children,
            "master" => page.MasterShapes,
            "other-page" => document.AddPage("Source page").Shapes,
            _ => page.Shapes
        };
        var source = Box(sources, "Source", LongText());
        source.TextRoot.SetAttributeValue(OdfNamespaces.Draw + "chain-next-name", terminal.Name);
        document.MarkPartDirty(placement == "master" ? "styles.xml" : "content.xml");

        foreach (OdgDocument read in RoundTrips(document)) {
            string[] before = XmlState(read);
            var result = read.Pages[0].ToDrawing();
            Assert.Contains(result.Report.Mappings, m => m.Feature == "shape:Terminal:text:text-chain-flow" &&
                m.Status == OdfConversionMappingStatus.Unsupported);
            Assert.Throws<OdfConversionLossException>(() => read.Pages[0].ToDrawing(
                OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
            if (placement != "other-page") {
                Assert.Contains(result.Report.Mappings, m => m.Feature == "shape:Source:text:text-chain-flow" &&
                    m.Status == OdfConversionMappingStatus.Unsupported);
                Assert.Equal(48D, Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>()).Height);
                Assert.DoesNotContain(result.Report.Mappings, m => m.Feature == "shape:Source:text:text-auto-size" &&
                    m.Status == OdfConversionMappingStatus.Approximated);
            }
            Assert.Equal(before, XmlState(read));
            Assert.Equal("Terminal", read.GetXml(placement == "master" ? "styles.xml" : "content.xml")
                .Descendants(OdfNamespaces.Draw + "text-box").Select(e => (string?)e.Attribute(
                    OdfNamespaces.Draw + "chain-next-name")).Single(value => value != null));
        }
    }

    [Fact]
    public void UnresolvedChainBlocksOnlyItsSourceAndPreservesTheReference() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var linked = Box(page.Shapes, "Linked", LongText());
        Box(page.Shapes, "Independent", LongText());
        linked.TextRoot.SetAttributeValue(OdfNamespaces.Draw + "chain-next-name", "Missing target");
        document.MarkPartDirty("content.xml");
        var result = page.ToDrawing();
        Assert.Contains(result.Report.Mappings, m => m.Feature == "shape:Linked:text:text-chain-flow" &&
            m.Status == OdfConversionMappingStatus.Unsupported);
        Assert.DoesNotContain(result.Report.Mappings, m => m.Feature == "shape:Independent:text:text-chain-flow");
        double[] heights = result.Value.Elements.OfType<OfficeDrawingRichText>().Select(t => t.Height).ToArray();
        Assert.Equal(48D, heights[0]); Assert.True(heights[1] > 48D);
        Assert.Equal("Missing target", (string?)linked.TextRoot.Attribute(OdfNamespaces.Draw + "chain-next-name"));
    }

    private static OdgShape Box(OdgShapes shapes, string name, string text) {
        var shape = shapes.AddTextBox(new OdfRect(OdfLength.Points(20), OdfLength.Points(20),
            OdfLength.Points(120), OdfLength.Points(48)), text, name);
        shape.AutoGrowHeight = true; shape.AutoGrowWidth = false; shape.WrapText = true;
        shape.TextVerticalAlignment = OdfTextAreaVerticalAlignment.Top;
        shape.TextFitMode = OdfTextFitMode.None;
        if (shape.Paragraphs.Count > 0) shape.Paragraphs[0].FontSize = OdfLength.Points(12);
        return shape;
    }

    private static string LongText() => string.Join(" ", Enumerable.Repeat("Alpha beta gamma delta", 12));
    private static string[] XmlState(OdgDocument document) => new[] {
        document.GetXml("content.xml").ToString(), document.GetXml("styles.xml").ToString()
    };
    private static IEnumerable<OdgDocument> RoundTrips(OdgDocument document) {
        using var package = new MemoryStream(); document.Save(package); package.Position = 0;
        yield return OdgDocument.Load(package);
        using var flat = new MemoryStream(); document.SaveFlatXml(flat); flat.Position = 0;
        yield return OdgDocument.LoadFlatXml(flat);
    }
}
