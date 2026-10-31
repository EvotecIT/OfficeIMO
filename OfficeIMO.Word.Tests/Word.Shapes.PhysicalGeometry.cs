using System;
using System.Globalization;
using System.IO;
using System.Linq;
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

    [Fact]
    public void ShapeLengths_DrawingMlKeepsExactEmusAndRejectsOutlineOverflowBeforeMutation() {
        using var output = new MemoryStream();
        using (var document = WordDocument.Create()) {
            var shape = document.AddParagraph().AddShapeDrawing(WordShapeType.Rectangle,
                OfficeLength.Emus(9526), OfficeLength.Emus(19051), OfficeLength.Emus(3), OfficeLength.Emus(5));
            shape.StrokeWeight = OfficeLength.Emus(7);
            Assert.Throws<OverflowException>(() => shape.StrokeWeight = OfficeLength.Emus((long)int.MaxValue + 1));
            Assert.Equal(OfficeLength.Emus(7), shape.StrokeWeight);
            shape.StrokeWeight = null;
            Assert.Equal(OfficeLength.Emus(7), shape.StrokeWeight);
            document.Save(output);
        }
        output.Position = 0;
        using (var package = WordprocessingDocument.Open(output, false)) {
            var anchor = Assert.Single(package.MainDocumentPart!.Document.Descendants<DW.Anchor>());
            Assert.Equal(9526L, anchor.Extent!.Cx!.Value);
            Assert.Equal(19051L, anchor.Extent.Cy!.Value);
            Assert.Equal("3", anchor.HorizontalPosition!.GetFirstChild<DW.PositionOffset>()!.Text);
            Assert.Equal("5", anchor.VerticalPosition!.GetFirstChild<DW.PositionOffset>()!.Text);
            Assert.Equal(7, Assert.Single(anchor.Descendants<A.Outline>()).Width!.Value);
        }
        output.Position = 0;
        using var loaded = WordDocument.Load(output);
        Assert.True(loaded.Paragraphs[0].Shape!.TryGetLayoutSnapshot(out var layout));
        Assert.Equal(9526L, OfficeLength.Points(layout.WidthPoints).ToEmus());
        Assert.Equal(19051L, OfficeLength.Points(layout.HeightPoints).ToEmus());
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
