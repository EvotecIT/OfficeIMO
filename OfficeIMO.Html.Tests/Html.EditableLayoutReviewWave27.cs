using A = DocumentFormat.OpenXml.Drawing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlEditableLayoutReviewWave27Tests {
    [Theory]
    [InlineData("position:relative;left:80px")]
    [InlineData("position:relative;right:12px")]
    [InlineData("position:relative;top:16px")]
    [InlineData("position:relative;bottom:4px")]
    [InlineData("position:sticky;top:8px")]
    public void OffsetRelativeAndStickyDescendantsStayInSemanticFlow(string declaration) {
        string html = "<div style='position:absolute;width:220px;height:80px'>"
            + "<span style='" + declaration + "'>Offset content</span></div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
            && diagnostic.Detail == "nestedLayoutPlacement=true; semanticFlow=true");
    }

    [Fact]
    public void LoadedCustomGeometryTextBoxInsertsFillBeforeOutlineAndEffects() {
        using var stream = new MemoryStream();
        using (WordDocument source = WordDocument.Create()) {
            source.AddTextBox("Custom geometry");
            source.Save(stream);
        }

        using WordDocument loaded = WordDocument.Load(new MemoryStream(stream.ToArray()));
        WordTextBox textBox = Assert.Single(loaded.TextBoxes);
        Wps.ShapeProperties properties = Assert.IsType<Wps.ShapeProperties>(textBox.DrawingShapeProperties);
        properties.GetFirstChild<A.PresetGeometry>()?.Remove();
        properties.Append(new A.CustomGeometry());
        properties.Append(new A.Outline());
        properties.Append(new A.EffectList());

        textBox.FillColorHex = "ABCDEF";

        List<DocumentFormat.OpenXml.OpenXmlElement> children = properties.ChildElements.ToList();
        int fillIndex = children.IndexOf(Assert.Single(properties.Elements<A.SolidFill>()));
        Assert.True(fillIndex > children.IndexOf(Assert.Single(properties.Elements<A.CustomGeometry>())));
        Assert.True(fillIndex < children.IndexOf(Assert.Single(properties.Elements<A.Outline>())));
        Assert.True(fillIndex < children.IndexOf(Assert.Single(properties.Elements<A.EffectList>())));
    }

    [Fact]
    public void CustomGeometryWordShapeInsertsFillBeforeOutlineAndEffects() {
        using WordDocument document = WordDocument.Create();
        WordShape shape = document.AddShapeDrawing(WordShapeType.Rectangle, 80, 40);
        Wps.ShapeProperties properties = Assert.IsType<Wps.ShapeProperties>(
            shape._wpsShape?.GetFirstChild<Wps.ShapeProperties>());
        properties.GetFirstChild<A.PresetGeometry>()?.Remove();
        properties.Append(new A.CustomGeometry());
        properties.Append(new A.Outline());
        properties.Append(new A.EffectList());

        shape.FillColorHex = "ABCDEF";

        List<DocumentFormat.OpenXml.OpenXmlElement> children = properties.ChildElements.ToList();
        int fillIndex = children.IndexOf(Assert.Single(properties.Elements<A.SolidFill>()));
        Assert.True(fillIndex > children.IndexOf(Assert.Single(properties.Elements<A.CustomGeometry>())));
        Assert.True(fillIndex < children.IndexOf(Assert.Single(properties.Elements<A.Outline>())));
        Assert.True(fillIndex < children.IndexOf(Assert.Single(properties.Elements<A.EffectList>())));
    }

    [Fact]
    public void ProjectedFixedHeightWordRegionDisablesShapeAutoFitAfterReopen() {
        const string html = "<div style='position:absolute;width:120px;height:18px'>"
            + "Text that is intentionally longer than the fixed rendered box height.</div>";
        HtmlToWordResult result = HtmlConversionDocument.Parse(html).ToWordDocumentResult();
        using var stream = new MemoryStream();
        result.Value.Save(stream);
        result.Value.Dispose();

        using WordDocument reopened = WordDocument.Load(new MemoryStream(stream.ToArray()));
        WordTextBox textBox = Assert.Single(reopened.TextBoxes);

        Assert.Equal(WordTextBoxAutoFitType.NoAutoFit, textBox.AutoFit);
        Assert.Equal(171450L, textBox.Height);
    }
}