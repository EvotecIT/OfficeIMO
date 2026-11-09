using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void ShapeLengths_VmlAndLineRetainSubPointGeometryAfterSave() {
        using var output = new MemoryStream();
        using (var document = WordDocument.Create()) {
            var shape = document.AddParagraph().AddShape(OfficeLength.Emus(2), OfficeLength.Millimeters(2));
            shape.Left = OfficeLength.Emus(-2);
            shape.Top = OfficeLength.Emus(12701);
            shape.StrokeWeight = OfficeLength.Emus(2);
            var line = document.AddParagraph().AddLine(OfficeLength.Emus(-2), OfficeLength.Zero,
                OfficeLength.Inches(1), OfficeLength.Millimeters(2), strokeWeight: OfficeLength.Emus(2));
            line.StrokeWeight = OfficeLength.Emus(3);
            document.Save(output);
        }
        output.Position = 0;
        using (var package = WordprocessingDocument.Open(output, false)) {
            var rectangle = Assert.Single(package.MainDocumentPart!.Document.Descendants<V.Rectangle>());
            Assert.Equal(2L, VmlStyleEmus(rectangle.Style!.Value!, "width"));
            Assert.Equal(-2L, VmlStyleEmus(rectangle.Style.Value!, "margin-left"));
            Assert.Equal(12701L, VmlStyleEmus(rectangle.Style.Value!, "margin-top"));
            Assert.Equal(2L, VmlEmus(rectangle.StrokeWeight!.Value!));
            var line = Assert.Single(package.MainDocumentPart.Document.Descendants<V.Line>());
            Assert.Equal(-2L, VmlEmus(line.From!.Value!.Split(',')[0]));
            Assert.Equal(914400L, VmlEmus(line.To!.Value!.Split(',')[0]));
            Assert.Equal(72000L, VmlEmus(line.To.Value!.Split(',')[1]));
            Assert.Equal(3L, VmlEmus(line.StrokeWeight!.Value!));
        }
        output.Position = 0;
        using var loaded = WordDocument.Load(output);
        var loadedShape = loaded.Paragraphs[0].Shape!;
        Assert.Equal(OfficeLength.Emus(2), loadedShape.Width);
        Assert.Equal(OfficeLength.Millimeters(2), loadedShape.Height);
        Assert.Equal(OfficeLength.Emus(-2), loadedShape.Left);
        Assert.Equal(OfficeLength.Emus(2), loadedShape.StrokeWeight);
        Assert.Equal(OfficeLength.Emus(3), loaded.Paragraphs[1].Line!.StrokeWeight);
    }

    [Theory]
    [InlineData("inline")]
    [InlineData("anchored")]
    [InlineData("simple")]
    public void ShapeLengths_DrawingMlEditsPersistWithoutChangingGraphicMetadata(string placement) {
        using var output = new MemoryStream();
        using (var document = WordDocument.Create()) {
            var paragraph = document.AddParagraph("Geometry marker");
            bool anchored = placement != "inline";
            var shape = anchored
                ? paragraph.AddShapeDrawing(WordShapeType.Rectangle, OfficeLength.Emus(9526), OfficeLength.Emus(19051),
                    OfficeLength.Emus(3), OfficeLength.Emus(5))
                : paragraph.AddShapeDrawing(WordShapeType.Rectangle, OfficeLength.Emus(9526), OfficeLength.Emus(19051));
            shape.FillColor = OfficeColor.Blue;
            var drawing = shape._drawing!;
            var withoutOutline = drawing.OuterXml;
            shape.StrokeWeight = null;
            Assert.Equal(withoutOutline, drawing.OuterXml);
            shape.StrokeWeight = OfficeLength.Emus(7);
            shape.StrokeColor = OfficeColor.Parse("1F4E79");
            var outline = Assert.Single(drawing.Descendants<A.Outline>());
            outline.CapType = A.LineCapValues.Round;
            var outlineWithoutWidth = (A.Outline)outline.CloneNode(true);
            outlineWithoutWidth.RemoveAttribute("w", string.Empty);
            if (placement == "simple") {
                drawing.Anchor!.SimplePos = true;
                drawing.Anchor.SimplePosition!.X = 11L;
                drawing.Anchor.SimplePosition.Y = 13L;
                Assert.Equal(OfficeLength.Emus(11), shape.Left);
                Assert.Equal(OfficeLength.Emus(13), shape.Top);
                shape.Left = OfficeLength.Emus(23);
                Assert.True(drawing.Anchor.SimplePos!.Value);
                Assert.Equal(OfficeLength.Emus(23), shape.Left);
                Assert.Equal(OfficeLength.Emus(13), shape.Top);
            }
            var properties = drawing.Inline?.DocProperties ?? drawing.Anchor!.GetFirstChild<DW.DocProperties>()!;
            properties.Name = "Geometry shape";
            properties.Description = "Shape description";
            var nonVisualMetadata = properties.OuterXml;
            var graphic = drawing.Descendants<A.Graphic>().Single();
            var bodyProperties = graphic.Descendants<DocumentFormat.OpenXml.Office2010.Word.DrawingShape.TextBodyProperties>().Single();
            bodyProperties.Rotation = 120000;
            var bodyMetadata = bodyProperties.OuterXml;
            Assert.Equal(OfficeLength.Emus(9526), shape.Width);
            Assert.Equal(OfficeLength.Emus(19051), shape.Height);
            shape.Left = null;
            Assert.Null(shape.Left);
            if (placement == "simple") Assert.Equal(OfficeLength.Emus(13), shape.Top);
            shape.Top = null;
            Assert.Null(shape.Top);
            Assert.Equal(!anchored, drawing.Inline != null);
            var beforeInvalid = drawing.OuterXml;
            Assert.Throws<ArgumentOutOfRangeException>(() => shape.Width = OfficeLength.Emus(-1));
            Assert.Throws<ArgumentOutOfRangeException>(() => shape.Height = OfficeLength.Zero);
            Assert.Throws<OverflowException>(() => shape.Left = OfficeLength.Emus((long)int.MaxValue + 1));
            Assert.Throws<OverflowException>(() => shape.Top = OfficeLength.Emus((long)int.MinValue - 1));
            Assert.Throws<OverflowException>(() => shape.StrokeWeight = OfficeLength.Emus((long)int.MaxValue + 1));
            Assert.Equal(beforeInvalid, drawing.OuterXml);
            shape.Width = OfficeLength.Inches(1);
            shape.Height = OfficeLength.Inches(0.25);
            Assert.Equal(!anchored, drawing.Inline != null);
            var initialFrame = drawing.Inline?.Extent ?? drawing.Anchor!.Extent;
            Assert.Equal(914400L, initialFrame!.Cx!.Value);
            Assert.Equal(228600L, initialFrame.Cy!.Value);
            var initialTransform = graphic.Descendants<A.Transform2D>().Single();
            Assert.Equal(914400L, initialTransform.Extents!.Cx!.Value);
            Assert.Equal(228600L, initialTransform.Extents.Cy!.Value);
            var metadata = graphic.OuterXml;
            shape.Left = OfficeLength.Pixels(32);
            shape.Top = OfficeLength.Points(12);
            Assert.Equal(metadata, graphic.OuterXml);
            Assert.Equal(bodyMetadata, bodyProperties.OuterXml);
            Assert.Equal(nonVisualMetadata, drawing.Anchor!.GetFirstChild<DW.DocProperties>()!.OuterXml);
            Assert.Same(graphic, drawing.Descendants<A.Graphic>().Single());
            Assert.Null(drawing.Inline);
            // Resize after conversion so stale references to a cloned graphic cannot pass.
            shape.Width = OfficeLength.Inches(2);
            shape.Height = OfficeLength.Inches(0.5);
            shape.StrokeWeight = null;
            Assert.Null(shape.StrokeWeight);
            Assert.Same(outline, Assert.Single(drawing.Descendants<A.Outline>()));
            Assert.Equal(outlineWithoutWidth.OuterXml, outline.OuterXml);
            Assert.Equal(OfficeColor.Parse("1F4E79"), shape.StrokeColor);
            Assert.Equal(OfficeLength.Inches(2), shape.Width);
            Assert.Equal(OfficeLength.Inches(0.5), shape.Height);
            Assert.Equal(OfficeLength.Pixels(32), shape.Left);
            Assert.Equal(OfficeLength.Points(12), shape.Top);
            Assert.Equal(bodyMetadata, bodyProperties.OuterXml);
            Assert.Equal(nonVisualMetadata, drawing.Anchor!.GetFirstChild<DW.DocProperties>()!.OuterXml);
            document.Save(output);
        }
        output.Position = 0;
        using (var package = WordprocessingDocument.Open(output, false)) {
            var anchor = Assert.Single(package.MainDocumentPart!.Document.Descendants<DW.Anchor>());
            Assert.Equal(1828800L, anchor.Extent!.Cx!.Value);
            Assert.Equal(457200L, anchor.Extent.Cy!.Value);
            var transform = Assert.Single(anchor.Descendants<A.Transform2D>());
            Assert.Equal(1828800L, transform.Extents!.Cx!.Value);
            Assert.Equal(457200L, transform.Extents.Cy!.Value);
            Assert.Equal(0L, transform.Offset!.X!.Value);
            Assert.Equal(0L, transform.Offset.Y!.Value);
            Assert.Equal("304800", anchor.HorizontalPosition!.PositionOffset!.Text);
            Assert.Equal("152400", anchor.VerticalPosition!.PositionOffset!.Text);
            var outline = Assert.Single(anchor.Descendants<A.Outline>());
            Assert.Null(outline.Width);
            Assert.Equal(A.LineCapValues.Round, outline.CapType!.Value);
            Assert.Equal("1F4E79", Assert.Single(outline.Descendants<A.RgbColorModelHex>()).Val!.Value);
            Assert.Equal(120000, Assert.Single(anchor.Descendants<DocumentFormat.OpenXml.Office2010.Word.DrawingShape.TextBodyProperties>()).Rotation!.Value);
            Assert.Equal("Geometry shape", anchor.GetFirstChild<DW.DocProperties>()!.Name!.Value);
            Assert.Equal("Shape description", anchor.GetFirstChild<DW.DocProperties>()!.Description!.Value);
        }
        output.Position = 0;
        using var loaded = WordDocument.Load(output);
        var saved = Assert.Single(loaded.Shapes);
        Assert.Equal(OfficeLength.Inches(2), saved.Width);
        Assert.Equal(OfficeLength.Inches(0.5), saved.Height);
        Assert.Equal(OfficeLength.Pixels(32), saved.Left);
        Assert.Equal(OfficeLength.Points(12), saved.Top);
        Assert.Equal(OfficeColor.Blue, saved.FillColor);
        Assert.Null(saved.StrokeWeight);
        Assert.Equal(OfficeColor.Parse("1F4E79"), saved.StrokeColor);
        Assert.Equal("Geometry marker", loaded.Paragraphs[0].Text);
        saved.Left = OfficeLength.Emus(-2);
        saved.Top = null;
        Assert.Equal(OfficeLength.Emus(-2), saved.Left);
        Assert.Null(saved.Top);
        using var cleared = new MemoryStream();
        loaded.Save(cleared);
        cleared.Position = 0;
        using (var package = WordprocessingDocument.Open(cleared, false)) {
            var anchor = Assert.Single(package.MainDocumentPart!.Document.Descendants<DW.Anchor>());
            Assert.Equal("-2", anchor.HorizontalPosition!.PositionOffset!.Text);
            Assert.Null(anchor.VerticalPosition!.PositionOffset);
            Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2010).Validate(anchor));
        }
        cleared.Position = 0;
        using var reopened = WordDocument.Load(cleared);
        Assert.Equal(OfficeLength.Emus(-2), Assert.Single(reopened.Shapes).Left);
        Assert.Null(Assert.Single(reopened.Shapes).Top);
        Assert.Equal(OfficeLength.Inches(2), Assert.Single(reopened.Shapes).Width);
        Assert.Null(Assert.Single(reopened.Shapes).StrokeWeight);
        Assert.Equal(OfficeColor.Parse("1F4E79"), Assert.Single(reopened.Shapes).StrokeColor);
    }

    [Fact]
    public void ShapeLengths_NullVmlGeometryRemovesOnlyTheExplicitAttributes() {
        using var document = WordDocument.Create();
        var shape = document.AddParagraph().AddShape(WordShapeType.RoundedRectangle,
            OfficeLength.Centimeters(1), OfficeLength.Points(12));
        Assert.Equal(OfficeColor.White, shape.FillColor);
        Assert.Equal(OfficeColor.Black, shape.StrokeColor);
        Assert.Equal(OfficeLength.Points(1), shape.StrokeWeight);
        shape.Left = OfficeLength.Emus(-2);
        shape.Top = OfficeLength.Emus(3);
        shape.Left = null;
        shape.Top = null;
        shape.StrokeWeight = null;
        using var output = new MemoryStream();
        document.Save(output);
        output.Position = 0;
        using var loaded = WordDocument.Load(output);
        var saved = loaded.Paragraphs[0].Shape!;
        Assert.Null(saved.Left);
        Assert.Null(saved.Top);
        Assert.Null(saved.StrokeWeight);
        Assert.Equal(OfficeLength.Centimeters(1), saved.Width);
        Assert.Equal(OfficeLength.Points(12), saved.Height);
        Assert.Equal(0.25, saved.ArcSize);
    }

    [Fact]
    public void ShapeLengths_VmlFactoryResultsMutateAndRemoveOnlyTheirCreatedSibling() {
        using var output = new MemoryStream();
        using (var document = WordDocument.Create()) {
            var paragraph = document.AddParagraph();
            var rectangle = paragraph.AddShape(OfficeLength.Points(10), OfficeLength.Points(11));
            var firstEllipse = WordShape.AddEllipse(paragraph, OfficeLength.Points(20), OfficeLength.Points(21));
            var ellipse = paragraph.AddShape(WordShapeType.Ellipse, OfficeLength.Points(30), OfficeLength.Points(31),
                strokeColor: OfficeColor.Red, strokeWeight: OfficeLength.Points(3));
            var rounded = WordShape.AddRoundedRectangle(paragraph, OfficeLength.Points(40), OfficeLength.Points(41));
            var line = WordShape.AddLine(paragraph, OfficeLength.Zero, OfficeLength.Zero,
                OfficeLength.Points(50), OfficeLength.Points(51));
            var polygon = WordShape.AddPolygon(paragraph, "0,0 20,0 10,20");
            Assert.Equal(6, document.Shapes.Count);
            Assert.Equal(OfficeLength.Points(30), ellipse.Width);
            Assert.Equal(OfficeLength.Points(40), rounded.Width);
            ellipse.Width = OfficeLength.Points(32);
            line.StrokeWeight = OfficeLength.Points(5);
            polygon.FillColor = OfficeColor.Blue;
            Assert.Equal(OfficeLength.Points(10), rectangle.Width);
            Assert.Equal(OfficeColor.White, rectangle.FillColor);
            Assert.Null(rectangle.StrokeWeight);
            Assert.Equal(OfficeLength.Points(20), firstEllipse.Width);
            Assert.Null(firstEllipse.StrokeWeight);
            Assert.Equal(OfficeColor.White, rounded.FillColor);
            ellipse.Remove();
            rounded.Remove();
            line.Remove();
            polygon.Remove();
            document.Save(output);
        }
        output.Position = 0;
        using (var package = WordprocessingDocument.Open(output, false)) {
            var xml = package.MainDocumentPart!.Document;
            Assert.Equal(127000L, VmlStyleEmus(Assert.Single(xml.Descendants<V.Rectangle>()).Style!.Value!, "width"));
            Assert.Equal(254000L, VmlStyleEmus(Assert.Single(xml.Descendants<V.Oval>()).Style!.Value!, "width"));
            Assert.Empty(xml.Descendants<V.RoundRectangle>());
            Assert.Empty(xml.Descendants<V.Line>());
            Assert.Empty(xml.Descendants<V.PolyLine>());
        }
        output.Position = 0;
        using var loaded = WordDocument.Load(output);
        Assert.Equal(new[] { OfficeLength.Points(10), OfficeLength.Points(20) }, loaded.Shapes.Select(shape => shape.Width));
    }

    private static long VmlStyleEmus(string style, string name) =>
        VmlEmus(style.Split(';').Single(value => value.StartsWith(name + ":", StringComparison.Ordinal)).Substring(name.Length + 1));

    private static long VmlEmus(string text) =>
        checked((long)decimal.Round(decimal.Parse(text.Substring(0, text.Length - 2), CultureInfo.InvariantCulture) * 12700M));
}
