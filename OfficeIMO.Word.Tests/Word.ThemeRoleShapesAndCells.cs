using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Word;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;

namespace OfficeIMO.Tests;

public class WordThemeRoleShapesAndCells {
    [Theory]
    [InlineData("drawing")]
    [InlineData("rectangle")]
    [InlineData("ellipse")]
    [InlineData("rounded")]
    [InlineData("polygon")]
    [InlineData("line")]
    [InlineData("group")]
    public void ShapeRolesUseNativeSlotsOrDocumentThemeLiterals(string route) {
        var theme = OfficeTheme.Modern.WithPalette(OfficeTheme.Modern.Palette.With(
            accent1: OfficeColor.ParseHex("123456"), accent2: OfficeColor.ParseHex("654321")));
        using var document = WordDocument.Create(new WordCreateOptions { Theme = theme });
        var paragraph = document.AddParagraph();
        if (route == "group") {
            paragraph.AddShapeGroup(new[] {
                new WordShapeGroupItem(WordShapeType.Rectangle, OfficeLength.Zero, OfficeLength.Zero,
                    OfficeLength.Points(20), OfficeLength.Points(10)) {
                    FillColor = OfficeColor.Accent1, StrokeColor = OfficeColor.Accent2
                },
                new WordShapeGroupItem(WordShapeType.Ellipse, OfficeLength.Points(20), OfficeLength.Zero,
                    OfficeLength.Points(20), OfficeLength.Points(10)) { FillColor = OfficeColor.Success }
            });
        } else if (route == "line") {
            var line = paragraph.AddLine(OfficeLength.Zero, OfficeLength.Zero, OfficeLength.Points(20),
                OfficeLength.Points(10), OfficeColor.Accent2);
            line.Color = OfficeColor.Accent1;
        } else {
            var width = OfficeLength.Points(20);
            var height = OfficeLength.Points(10);
            var shape = route switch {
                "drawing" => paragraph.AddShapeDrawing(WordShapeType.Rectangle, width, height),
                "rectangle" => paragraph.AddShape(WordShapeType.Rectangle, width, height, OfficeColor.Accent1, OfficeColor.Accent2),
                "ellipse" => WordShape.AddEllipse(paragraph, width, height, OfficeColor.Accent1),
                "rounded" => WordShape.AddRoundedRectangle(paragraph, width, height, OfficeColor.Accent1),
                _ => WordShape.AddPolygon(paragraph, "0,0 20,0 20,10", OfficeColor.Accent1, OfficeColor.Accent2)
            };
            shape.FillColor = OfficeColor.Accent1;
            shape.StrokeColor = OfficeColor.Accent2;
        }

        using var stream = new MemoryStream(document.ToBytes());
        using var package = WordprocessingDocument.Open(stream, false);
        var body = package.MainDocumentPart!.Document.Body!;
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(package));
        if (route == "drawing" || route == "group") {
            var schemes = body.Descendants<A.SchemeColor>().Select(c => c.GetAttribute("val", "").Value).ToArray();
            Assert.Contains("accent1", schemes);
            Assert.Contains("accent2", schemes);
            if (route == "group") Assert.Contains(body.Descendants<A.RgbColorModelHex>(),
                c => c.Val?.Value == theme.Palette.Success.ToRgbHex());
        } else if (route == "line") {
            Assert.Equal("#123456", Assert.Single(body.Descendants<V.Line>()).StrokeColor!.Value);
        } else {
            var vml = body.Descendants().Single(e => e.NamespaceUri == "urn:schemas-microsoft-com:vml"
                && new[] { "rect", "oval", "roundrect", "polyline" }.Contains(e.LocalName));
            Assert.Equal("#123456", vml.GetAttribute("fillcolor", "").Value);
            Assert.Equal("#654321", vml.GetAttribute("strokecolor", "").Value);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task LoadedShapeAndCellRolesFollowEmbeddedThemeAndLiteralEditsClearOverrides(bool disposeAsync) {
        var theme = OfficeTheme.Modern.WithPalette(OfficeTheme.Modern.Palette.With(
            accent1: OfficeColor.ParseHex("123456")));
        byte[] original;
        using (var document = WordDocument.Create(new WordCreateOptions { Theme = theme })) {
            document.AddParagraph().AddShapeDrawing(WordShapeType.Rectangle, OfficeLength.Points(20), OfficeLength.Points(10));
            document.AddParagraph().AddShape(OfficeLength.Points(20), OfficeLength.Points(10));
            document.AddTable(1, 2);
            original = document.ToBytes();
        }
        using var input = new MemoryStream(original);
        using var loaded = WordDocument.Load(input);
        var shapes = loaded.Shapes.ToArray();
        var first = loaded.Tables[0].Rows[0].Cells[0];
        var second = loaded.Tables[0].Rows[0].Cells[1];
        first.ShadingFillColor = OfficeColor.Accent1;
        first.Borders.TopStyle = WordBorderStyle.Single;
        first.Borders.TopColor = OfficeColor.ParseHex("456789");
        shapes[0].FillColor = OfficeColor.Accent1;
        shapes[0].StrokeColor = OfficeColor.TextMuted;
        shapes[1].FillColor = OfficeColor.Accent1;
        Assert.Equal(OfficeColor.ParseHex("123456"), first.ShadingFillColor);
        Assert.Equal(OfficeColor.ParseHex("123456"), shapes[0].FillColor);
        Assert.Equal(theme.Palette.TextMuted, shapes[0].StrokeColor);
        second.ShadingFillColor = OfficeColor.Success;
        using (var stream = new MemoryStream(loaded.ToBytes()))
        using (var package = WordprocessingDocument.Open(stream, false)) {
            var body = package.MainDocumentPart!.Document.Body!;
            var shading = body.Descendants<W.Shading>().ToArray();
            Assert.Equal("accent1", shading[0].ThemeFill!.InnerText);
            Assert.Equal("123456", shading[0].Fill!.Value);
            Assert.Null(shading[1].ThemeFill);
            Assert.Equal(OfficeTheme.Default.Palette.Success.ToRgbHex(), shading[1].Fill!.Value);
            Assert.Equal("#123456", Assert.Single(body.Descendants<V.Rectangle>()).FillColor!.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(package));
        }
        WordShape[] detachedShapes;
        WordTableCell detachedCell;
        WordParagraph[] detachedParagraphs;
        WordSection[] detachedSections;
        using (var inspectionBytes = new MemoryStream(loaded.ToBytes()))
        using (var inspection = WordDocument.Load(inspectionBytes)) {
            detachedShapes = inspection.Shapes.ToArray();
            detachedCell = inspection.Tables[0].Rows[0].Cells[0];
            detachedParagraphs = inspection.Paragraphs.ToArray();
            detachedSections = inspection.Sections.ToArray();
            if (disposeAsync) await inspection.DisposeAsync();
            else inspection.Dispose();
        }
        Assert.Equal("123456", detachedShapes[0].FillColor.ToRgbHex());
        Assert.Equal(theme.Palette.TextMuted.ToRgbHex(), detachedShapes[0].StrokeColor.ToRgbHex());
        Assert.Equal("123456", detachedCell.ShadingFillColor!.Value.ToRgbHex());
        Assert.Equal("456789", detachedCell.Borders.TopColor!.Value.ToRgbHex());
        var lazyParagraphShapes = detachedParagraphs.Select(p => p.Shape).Where(s => s != null).Cast<WordShape>().ToArray();
        Assert.Equal(2, lazyParagraphShapes.Length);
        Assert.Equal("123456", lazyParagraphShapes[0].FillColor.ToRgbHex());
        Assert.Equal(theme.Palette.TextMuted.ToRgbHex(), lazyParagraphShapes[0].StrokeColor.ToRgbHex());
        Assert.Equal("123456", lazyParagraphShapes[1].FillColor.ToRgbHex());
        var lazySectionShapes = detachedSections.SelectMany(s => s.Shapes).ToArray();
        Assert.Equal(2, lazySectionShapes.Length);
        Assert.Equal("123456", lazySectionShapes[0].FillColor.ToRgbHex());
        Assert.Equal("123456", lazySectionShapes[1].FillColor.ToRgbHex());
        first.ShadingFillColor = OfficeColor.Red;
        shapes[0].FillColor = OfficeColor.Blue;
        shapes[0].StrokeColor = OfficeColor.Warning;
        second.ShadingFillColor = null;
        using var final = new MemoryStream(loaded.ToBytes());
        using var saved = WordprocessingDocument.Open(final, false);
        var savedBody = saved.MainDocumentPart!.Document.Body!;
        var savedShading = Assert.Single(savedBody.Descendants<W.Shading>());
        Assert.Null(savedShading.ThemeFill);
        Assert.Equal("FF0000", savedShading.Fill!.Value);
        Assert.Empty(savedBody.Descendants<A.SchemeColor>());
        Assert.Contains(savedBody.Descendants<A.RgbColorModelHex>(), c => c.Val?.Value == "0000FF");
        Assert.Contains(savedBody.Descendants<A.RgbColorModelHex>(), c => c.Val?.Value == OfficeTheme.Default.Palette.Warning.ToRgbHex());
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(saved));
    }
}
