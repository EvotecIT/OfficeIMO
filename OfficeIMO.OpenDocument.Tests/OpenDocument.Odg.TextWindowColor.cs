using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentOdgTextWindowColorTests {
    [Theory]
    [InlineData(null, "true")]
    [InlineData("100%", "true")]
    [InlineData("25%", "true")]
    [InlineData("100%", "1")]
    [InlineData("100%", "invalid")]
    public void UnsupportedWindowPolicyLossIsIndependentOfBodyOpacityAndRetainsFallbackPaint(string? opacity, string policy) {
        var document = OdgDocument.Create(); var paragraph = Shape(document).AddParagraph("Window");
        paragraph.Color = OdfColor.Parse("#FF0000");
        var style = paragraph.EnsureStyle(); SetWindowPolicy(style, policy); SetOpacity(style, opacity);
        foreach (var read in RoundTrips(document)) {
            string[] before = Parts(read);
            var result = read.Pages[0].ToDrawing();
            Assert.Equal(OfficeColor.Red, Text(result.Value).Paragraphs.Single().Runs.Single().Color);
            AssertLoss(result.Report, "text-window-color", true);
            AssertLoss(result.Report, "text-opacity", opacity == "25%");
            var paint = Assert.Single(TextPaint(result.Value));
            Assert.Equal("Window", paint.Value); Assert.Equal("#FF0000", (string?)paint.Attribute("fill"));
            Assert.Null(paint.Attribute("fill-opacity"));
            Assert.Throws<OdfConversionLossException>(() => read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
            var actual = read.Pages[0].Shapes.Single().Paragraphs.Single().Styles.First().TextProperties!;
            Assert.Equal(policy, (string?)actual.Attribute(OdfNamespaces.Style + "use-window-font-color"));
            Assert.Equal(opacity, (string?)actual.Attribute(OdfNamespaces.LoExt + "opacity"));
            Assert.Equal(before, Parts(read));
        }
    }

    [Theory]
    [InlineData(null, null, true)]
    [InlineData("false", null, false)]
    [InlineData("0", null, false)]
    [InlineData(" false ", null, false)]
    [InlineData(" 0 ", null, false)]
    [InlineData(null, "#0000FF", false)]
    [InlineData("1", "#0000FF", true)]
    public void NearestForegroundOverridesRetainTheExistingNestedStyleCascade(string? policy, string? color, bool windowColor) {
        var document = OdgDocument.Create(); var shape = Shape(document);
        var parent = document.Styles.CreateNamed("WindowParent", OdfStyleFamily.Paragraph);
        parent.Color = OdfColor.Parse("#FF0000"); parent.TextOpacity = .5D; SetWindowPolicy(parent, "true");
        var paragraph = shape.AddParagraph(); paragraph.StyleName = parent.Name;
        var outer = paragraph.AddRun(); outer.Bold = true;
        var nested = outer.AddRun("Nested");
        if (policy != null) SetWindowPolicy(nested.EnsureStyle(), policy);
        if (color != null) nested.Color = OdfColor.Parse(color);
        foreach (var read in RoundTrips(document)) {
            string[] before = Parts(read);
            var result = read.Pages[0].ToDrawing();
            var run = Text(result.Value).Paragraphs.Single().Runs.Single();
            Assert.Equal("Nested", run.Text); Assert.True(run.Bold);
            byte alpha = windowColor ? (byte)255 : (byte)128;
            Assert.Equal(color == null ? OfficeColor.FromRgba(255, 0, 0, alpha) : OfficeColor.FromRgba(0, 0, 255, alpha), run.Color);
            AssertLoss(result.Report, "text-window-color", windowColor);
            AssertLoss(result.Report, "text-opacity", windowColor);
            if (windowColor) Assert.Throws<OdfConversionLossException>(() => read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
            else read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
            Assert.Equal(before, Parts(read));
        }
    }

    [Theory]
    [InlineData("label", null, "true")]
    [InlineData("label", "100%", "true")]
    [InlineData("label", "25%", "true")]
    [InlineData("label", "25%", "false")]
    [InlineData("label", "25%", "0")]
    [InlineData("leader", null, "1")]
    [InlineData("leader", "100%", "1")]
    [InlineData("leader", "25%", "1")]
    public void LabelAndInheritedLeaderPoliciesReportWindowLossWithoutChangingTheirAlphaProfile(string target, string? opacity, string policy) {
        var document = OdgDocument.Create(); var shape = Shape(document);
        bool label = target == "label", windowColor = policy is "true" or "1";
        if (label) {
            shape.AddList(true).AddItem("Body"); shape.Paragraphs.Single().Color = OdfColor.Parse("#FF0000");
            var level = document.GetXml("content.xml").Descendants(OdfNamespaces.Text + "list-style").Single().Elements().Single();
            level.Add(new XElement(OdfNamespaces.Style + "text-properties",
                new XAttribute(OdfNamespaces.Fo + "color", "#0000FF"),
                new XAttribute(OdfNamespaces.Style + "use-window-font-color", policy),
                opacity == null ? null : new XAttribute(OdfNamespaces.LoExt + "opacity", opacity)));
        } else {
            var paragraph = shape.AddParagraph("A\tB"); paragraph.Color = OdfColor.Parse("#FF0000");
            var parent = document.Styles.CreateNamed("WindowLeaderParent", OdfStyleFamily.Text);
            parent.Color = OdfColor.Parse("#0000FF"); SetWindowPolicy(parent, policy); SetOpacity(parent, opacity);
            var style = document.Styles.CreateNamed("WindowLeader", OdfStyleFamily.Text, parent.Name); style.Bold = true;
            paragraph.SetTabStops(new[] { new OdfTabStop(OdfLength.Points(100)).WithLeader(".").WithLeaderTextStyle(style) });
        }
        foreach (var read in RoundTrips(document)) {
            string[] before = Parts(read);
            var result = read.Pages[0].ToDrawing(); var paragraph = Text(result.Value).Paragraphs.Single();
            AssertLoss(result.Report, "text-window-color", windowColor);
            AssertLoss(result.Report, "text-opacity", label && windowColor && opacity == "25%");
            var paint = TextPaint(result.Value);
            if (label) {
                Assert.Equal(OfficeColor.Blue, paragraph.Label!.Run.Color);
                var marker = Assert.Single(paint, e => e.Value == "1.");
                Assert.Equal("#0000FF", (string?)marker.Attribute("fill")); Assert.Null(marker.Attribute("fill-opacity"));
            } else {
                var leader = paragraph.TabStops!.Stops.Single().LeaderStyle!;
                Assert.Equal(OfficeColor.Blue, leader.Color); Assert.True(leader.Bold);
                var glyph = Assert.Single(paint, e => e.Value.Length > 0 && e.Value.All(c => c == '.'));
                Assert.Equal("#0000FF", (string?)glyph.Attribute("fill"));
                Assert.Equal(opacity == "25%" ? OfficeSvgFormatting.FormatNumber(64 / 255D) : "1", (string?)glyph.Attribute("fill-opacity") ?? "1");
            }
            if (windowColor) Assert.Throws<OdfConversionLossException>(() => read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
            else read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
            Assert.Equal(before, Parts(read));
        }
    }

    private static void SetWindowPolicy(OdfStyle style, string value) =>
        style.SetProperty(OdfNamespaces.Style + "text-properties", OdfNamespaces.Style + "use-window-font-color", value);
    private static void SetOpacity(OdfStyle style, string? value) {
        if (value != null) style.SetProperty(OdfNamespaces.Style + "text-properties", OdfNamespaces.LoExt + "opacity", value);
    }
    private static void AssertLoss(OdfConversionReport report, string feature, bool expected) =>
        Assert.Equal(expected, report.Mappings.Any(m => m.Feature.EndsWith(":" + feature, StringComparison.Ordinal) && m.Status == OdfConversionMappingStatus.Unsupported));
    private static OdgShape Shape(OdgDocument document) {
        var shape = document.AddPage().Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 12, 8));
        shape.FontFamily = "Arial"; shape.FontSize = OdfLength.Points(14);
        shape.EnsureGraphicStyle().SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "wrap-option", "no-wrap");
        return shape;
    }
    private static OfficeDrawingRichText Text(OfficeDrawing drawing) => Assert.Single(drawing.Elements.OfType<OfficeDrawingRichText>());
    private static XElement[] TextPaint(OfficeDrawing drawing) => XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(drawing))
        .Descendants(XName.Get("text", "http://www.w3.org/2000/svg")).ToArray();
    private static string[] Parts(OdgDocument document) => new[] { "content.xml", "styles.xml" }.Select(part => document.GetXml(part).ToString()).ToArray();
    private static OdgDocument[] RoundTrips(OdgDocument document) {
        using var flat = new MemoryStream(); document.SaveFlatXml(flat); flat.Position = 0;
        return new[] { document, OdgDocument.Load(new MemoryStream(document.ToBytes(new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.PreserveSource }))), OdgDocument.LoadFlatXml(flat) };
    }
}
