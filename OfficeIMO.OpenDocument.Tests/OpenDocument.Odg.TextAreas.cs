using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentOdgTextAreaTests {
    private static OdgDocument Document(string area, bool rectangle) {
        var doc = OdgDocument.Create(); var page = doc.AddPage();
        var bounds = OdfRect.FromCentimeters(1, 1, 15, 5);
        var shape = rectangle ? page.Shapes.AddRectangle(bounds, "Area") : page.Shapes.AddTextBox(bounds, "", "Area");
        var p = rectangle ? shape.AddParagraph("") : shape.Paragraphs[0];
        p.Text = "Before\nA\t123,45\nAfter"; p.TextAlign = "right";
        p.FontFamily = "Liberation Sans"; p.FontSize = OdfLength.Points(12);
        p.SetTabStops(new[] { new OdfTabStop(OdfLength.Points(220)) });
        var style = doc.Styles.Find(OdfStyleFamily.Graphic, (string)shape.Element.Attribute(OdfNamespaces.Draw + "style-name")!)!;
        style.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "textarea-horizontal-align", area);
        style.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "wrap-option", "wrap");
        return doc;
    }

    [Theory]
    [InlineData("left", true, OfficeTextAreaAlignment.Left)]
    [InlineData("center", true, OfficeTextAreaAlignment.Center)]
    [InlineData("right", true, OfficeTextAreaAlignment.Right)]
    [InlineData("left", false, OfficeTextAreaAlignment.Left)]
    [InlineData("center", false, OfficeTextAreaAlignment.Center)]
    [InlineData("right", false, OfficeTextAreaAlignment.Right)]
    public void IntrinsicAreasRetainParagraphAlignmentAcrossBothContainers(string area, bool rectangle, OfficeTextAreaAlignment expected) {
        var doc = Document(area, rectangle);
        using var flat = new MemoryStream(); doc.SaveFlatXml(flat); flat.Position = 0;
        foreach (var loaded in new[] { doc, OdgDocument.Load(new MemoryStream(doc.ToBytes())), OdgDocument.LoadFlatXml(flat) }) {
            string before = loaded.GetXml("content.xml").ToString();
            var result = loaded.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
            var text = Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>());
            Assert.Equal(expected, text.TextAreaAlignment); Assert.Equal(!rectangle, text.WrapText);
            Assert.Equal("Before\nA\t123,45\nAfter", text.PlainText);
            Assert.Equal(OfficeTextAlignment.Right, Assert.Single(text.Paragraphs).Alignment);
            Assert.Contains(result.Report.Mappings, m => m.Feature.EndsWith(":text-area-alignment", StringComparison.Ordinal) && m.Status == OdfConversionMappingStatus.Approximated);
            Assert.Equal(rectangle, result.Report.Mappings.Any(m => m.Feature.EndsWith(":text-area-wrapping", StringComparison.Ordinal)));
            Assert.Equal(before, loaded.GetXml("content.xml").ToString());
        }
    }

    [Theory]
    [InlineData("libreoffice-intrinsic-rectangle-areas.fodg", false)]
    [InlineData("libreoffice-intrinsic-textbox-areas.fodg", true)]
    public void NativeResavesRetainTheDistinctRectangleAndTextBoxWrapProfiles(string fixture, bool wrapped) {
        var doc = OdgDocument.LoadFlatXml(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", fixture));
        string before = doc.GetXml("content.xml").ToString(); var result = doc.Pages[0].ToDrawing();
        var texts = result.Value.Elements.OfType<OfficeDrawingRichText>().ToArray(); Assert.Equal(36, texts.Length);
        Assert.All(texts, text => { Assert.NotEqual(OfficeTextAreaAlignment.FullWidth, text.TextAreaAlignment); Assert.Equal(wrapped, text.WrapText); });
        Assert.Equal(36, result.Report.Mappings.Count(m => m.Feature.EndsWith(":text-area-alignment", StringComparison.Ordinal) && m.Status == OdfConversionMappingStatus.Approximated));
        Assert.DoesNotContain(result.Report.Mappings, m => m.Feature.EndsWith(":text-area-alignment", StringComparison.Ordinal) && m.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Equal(before, doc.GetXml("content.xml").ToString());
    }

    [Fact]
    public void NonLeftIndentedTabAreasRemainOutsideTheStrictNativeProfile() {
        foreach (bool rectangle in new[] { true, false }) {
            var doc = Document("center", rectangle);
            var p = doc.Pages[0].Shapes[0].Paragraphs[0]; p.Text = "A\t123,45";
            p.EnsureStyle().TextIndent = OdfLength.Points(25);
            string before = doc.GetXml("content.xml").ToString(); var result = doc.Pages[0].ToDrawing();
            Assert.Contains(result.Report.Mappings, m => m.Feature.EndsWith(":text-area-tab-indent", StringComparison.Ordinal) && m.Status == OdfConversionMappingStatus.Unsupported);
            Assert.Throws<OdfConversionLossException>(() => doc.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
            Assert.Equal("A\t123,45", Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>()).PlainText);
            Assert.Equal(before, doc.GetXml("content.xml").ToString());
        }
    }

    [Fact]
    public void UnknownAreaValuesRemainExplicitlyUnsupported() {
        var doc = Document("diagonal", rectangle: false); string before = doc.GetXml("content.xml").ToString();
        var result = doc.Pages[0].ToDrawing();
        Assert.Equal(OfficeTextAreaAlignment.FullWidth, Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>()).TextAreaAlignment);
        Assert.Contains(result.Report.Mappings, m => m.Feature.EndsWith(":text-area-alignment", StringComparison.Ordinal) && m.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => doc.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        Assert.Equal(before, doc.GetXml("content.xml").ToString());
    }
}
