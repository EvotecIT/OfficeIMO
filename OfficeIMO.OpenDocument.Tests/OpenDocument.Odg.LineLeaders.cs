using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentOdgLineLeaderTests {
    [Theory]
    [InlineData("solid", OfficeTextTabLineLeaderStyle.Solid)]
    [InlineData("dotted", OfficeTextTabLineLeaderStyle.Dotted)]
    [InlineData("dash", OfficeTextTabLineLeaderStyle.Dash)]
    [InlineData("long-dash", OfficeTextTabLineLeaderStyle.LongDash)]
    [InlineData("dot-dash", OfficeTextTabLineLeaderStyle.DotDash)]
    [InlineData("dot-dot-dash", OfficeTextTabLineLeaderStyle.DotDotDash)]
    [InlineData("wave", OfficeTextTabLineLeaderStyle.Wave)]
    public void TypedDeclarationsSurviveContainersCloneImportAndReadOnlyProjection(string pattern, OfficeTextTabLineLeaderStyle expected) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var paragraph = page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 12, 8), "Line").AddParagraph("A\tB");
        var stop = new OdfTabStop(OdfLength.Points(100));
        paragraph.SetTabStops(new[] { stop.WithLineLeader(new OdfTabLineLeader(pattern, "double", "2pt", OdfColor.Parse("#0000ff"))) });
        Assert.Null(stop.LineLeader);
        var imported = OdgDocument.Create(); imported.ImportPage(document, 0);
        document.ClonePage(0, "Cloned");
        foreach (var owner in new[] { document, imported }) {
            using var package = new MemoryStream(); owner.Save(package); package.Position = 0;
            using var flat = new MemoryStream(); owner.SaveFlatXml(flat); flat.Position = 0;
            foreach (var reopened in new[] { owner, OdgDocument.Load(package), OdgDocument.LoadFlatXml(flat) }) {
                string before = reopened.GetXml("content.xml").ToString();
                foreach (var reopenedPage in reopened.Pages) {
                    var result = reopenedPage.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
                    var text = Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>());
                    var leader = text.Paragraphs[0].TabStops!.Stops[0].LineLeader!;
                    Assert.Equal(expected, leader.Style); Assert.True(leader.DoubleLine);
                    Assert.Equal(2, leader.WidthPoints); Assert.Equal(OfficeColor.Blue, leader.Color);
                    Assert.Equal("A\tB", text.PlainText);
                    Assert.Contains(result.Report.Mappings, m => m.Feature.EndsWith(":tab-line-leaders") && m.Status == OdfConversionMappingStatus.Approximated);
                }
                Assert.Equal(before, reopened.GetXml("content.xml").ToString());
                Assert.All(reopened.GetXml("content.xml").Descendants(OdfNamespaces.Style + "tab-stop"), element => {
                    Assert.Equal(pattern, (string?)element.Attribute(OdfNamespaces.Style + "leader-style"));
                    Assert.Equal("double", (string?)element.Attribute(OdfNamespaces.Style + "leader-type"));
                    Assert.Equal("2pt", (string?)element.Attribute(OdfNamespaces.Style + "leader-width"));
                    Assert.Equal("#0000FF", (string?)element.Attribute(OdfNamespaces.Style + "leader-color"));
                });
            }
        }
    }

    [Theory]
    [InlineData("auto", .05)]
    [InlineData("normal", .05)]
    [InlineData("thin", .025)]
    [InlineData("medium", .05)]
    [InlineData("bold", .1)]
    [InlineData("thick", .1)]
    [InlineData("2", .1)]
    [InlineData("+2", .1)]
    [InlineData("50%", .025)]
    [InlineData("200%", .1)]
    public void RelativeWidthsUseTheDeclaredApproximationProfile(string width, double fraction) {
        var leader = new OdfTabLineLeader(width: width).ToDrawingLeader();
        Assert.Null(leader.WidthPoints); Assert.Equal(fraction, leader.WidthFontFraction, 8); Assert.Null(leader.Color);
    }

    [Fact]
    public void TextStyleWithoutTextDoesNotSuppressLinesAndTypeNoneSuppressesPaint() {
        var document = OdgDocument.Create(); var paragraph = document.AddPage().Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 12, 8)).AddParagraph("A\tB");
        paragraph.SetTabStops(new[] { new OdfTabStop(OdfLength.Points(100)).WithLineLeader(new OdfTabLineLeader("wave", "none")) });
        var element = paragraph.EnsureStyle().Element.Descendants(OdfNamespaces.Style + "tab-stop").Single();
        element.SetAttributeValue(OdfNamespaces.Style + "leader-text-style", "UnusedTextStyle");
        var result = document.Pages[0].ToDrawing();
        Assert.False(result.Report.HasSkippedOrUnsupported);
        var text = Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>());
        Assert.Equal(OfficeTextTabLineLeaderStyle.None, text.Paragraphs[0].TabStops!.Stops[0].LineLeader!.Style);
        element.SetAttributeValue(OdfNamespaces.Style + "leader-type", "double");
        result = document.Pages[0].ToDrawing(); Assert.False(result.Report.HasSkippedOrUnsupported);
        Assert.Contains(result.Report.Mappings, m => m.Feature.EndsWith(":tab-line-leaders"));
    }

    [Theory]
    [InlineData("0pt")]
    [InlineData("NaNpt")]
    [InlineData("-10%")]
    [InlineData("2.5")]
    [InlineData("unknown")]
    public void UnsupportedWidthsPreserveSourceAndTabSpacingAndStrictProjectionRejects(string width) {
        Assert.ThrowsAny<ArgumentException>(() => new OdfTabLineLeader(width: width));
        var document = OdgDocument.Create(); var paragraph = document.AddPage().Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 12, 8)).AddParagraph("A\tB");
        paragraph.SetTabStops(new[] { new OdfTabStop(OdfLength.Points(100)).WithLineLeader(new OdfTabLineLeader()) });
        var element = paragraph.EnsureStyle().Element.Descendants(OdfNamespaces.Style + "tab-stop").Single();
        element.SetAttributeValue(OdfNamespaces.Style + "leader-width", width);
        string before = document.GetXml("content.xml").ToString();
        var result = document.Pages[0].ToDrawing();
        Assert.True(result.Report.HasSkippedOrUnsupported);
        var text = Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>());
        Assert.Equal(100, text.Paragraphs[0].TabStops!.Stops[0].Position); Assert.Null(text.Paragraphs[0].TabStops!.Stops[0].LineLeader);
        Assert.Equal(before, document.GetXml("content.xml").ToString());
        Assert.Throws<OdfConversionLossException>(() => document.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
    }
}
