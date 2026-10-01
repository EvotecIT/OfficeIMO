using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Wpg = DocumentFormat.OpenXml.Office2010.Word.DrawingGroup;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public sealed class PdfShapeGroupTests {
    // Public reporter fixture: https://github.com/EvotecIT/OfficeIMO/issues/2675
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ReporterGroupedDrawingConvertsAsOnePageWithoutChangingSource(bool legacyOnly) {
        using WordDocument word = WordDocument.Load(Path.Combine(AppContext.BaseDirectory, "Documents", "Drawing", "WordGroupedWarning.docx"));
        if (legacyOnly) {
            var alternate = word.Paragraphs[0]._paragraph!.Descendants<AlternateContent>().Single();
            alternate.GetFirstChild<AlternateContentChoice>()!.Requires = "unknownDrawing";
            alternate.AddNamespaceDeclaration("unknownDrawing", "urn:unsupported-drawing");
        }
        var paragraph = Assert.Single(word.Paragraphs);
        Assert.False(paragraph.IsShape);
        if (!legacyOnly) {
            Assert.True(paragraph.IsShapeGroup);
            Assert.Equal(6, paragraph.ShapeGroup!.ChildCount);
            Assert.True(paragraph.ShapeGroup.TryGetLayoutSnapshot(out var layout));
            Assert.Equal(217D, layout!.WidthPoints, 6);
        }
        string before = word._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using var output = new MemoryStream();
        var result = word.SaveAsPdfResult(output);
        Assert.True(result.Succeeded, string.Join(" ", result.Diagnostics));
        Assert.Equal(before, word._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        var pdf = PdfCore.PdfDocument.Load(output.ToArray());
        Assert.Single(pdf.Reader.Pages());
        var bitmap = Render(pdf, 1);
        AssertPixel(bitmap, 90, 80, 255, 0, 0); // panel, in page points
        AssertPixel(bitmap, 95, 102, 255, 255, 255); // white warning triangle
        AssertPixel(bitmap, 103, 100, 0, 0, 0); // exclamation mark
        AssertPixel(bitmap, 300, 80, 255, 255, 255); // outside the 217-point group
        Assert.DoesNotContain(result.FidelityDiagnostics, item => item.Code == "NativeShapeGroupUnsupported");
        if (!legacyOnly) Assert.Contains(result.FidelityDiagnostics, item => item.Code == "NativeShapeGroupVmlFallback");
        SaveEvidence(pdf, legacyOnly ? "reporter-vml" : "reporter-drawingml");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeGroupScalesNestedCoordinatesAndKeepsAllChildren(bool alternateContent) {
        using WordDocument word = WordDocument.Create();
        var paragraph = word.AddParagraph();
        paragraph.AddShapeGroup(new[] {
            new WordShapeGroupItem(WordShapeType.Rectangle, 0, 0, 36, 18) { FillColorHex = "FF0000" },
            new WordShapeGroupItem(WordShapeType.Rectangle, 54, 18, 36, 18) { FillColorHex = "0000FF" }
        });
        var drawing = paragraph._run!.GetFirstChild<W.Drawing>()!;
        var group = drawing.Descendants<Wpg.WordprocessingGroup>().Single();
        var transform = group.GetFirstChild<Wpg.GroupShapeProperties>()!.GetFirstChild<A.TransformGroup>()!;
        transform.ChildOffset = new A.ChildOffset { X = 100, Y = 200 };
        foreach (var child in group.Elements<Wps.WordprocessingShape>()) {
            var t = child.GetFirstChild<Wps.ShapeProperties>()!.GetFirstChild<A.Transform2D>()!;
            t.Offset!.X = t.Offset.X!.Value + 100;
            t.Offset.Y = t.Offset.Y!.Value + 200;
        }
        var nestedTransform = (A.TransformGroup)transform.CloneNode(true);
        nestedTransform.Offset = new A.Offset { X = 100, Y = 200 };
        var first = group.Elements<Wps.WordprocessingShape>().First();
        first.Remove();
        group.Append(new Wpg.GroupShape(new Wpg.GroupShapeProperties(nestedTransform), first));
        drawing.Inline!.Extent!.Cx = 180 * 12700L;
        drawing.Inline.Extent.Cy = 72 * 12700L;
        if (alternateContent) WrapInAlternateContent(paragraph._run!, drawing, "wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
        Assert.True(paragraph.IsShapeGroup);
        Assert.False(paragraph.IsShape);
        var result = word.ToPdfDocumentResult();
        var pdf = PdfCore.PdfDocument.Load(result.Value.ToBytes());
        Assert.Single(pdf.Reader.Pages());
        Assert.DoesNotContain(result.Warnings, warning => warning.Code.StartsWith("NativeShapeGroup"));
        var bitmap = Render(pdf, 1);
        AssertPixel(bitmap, 80, 80, 255, 0, 0);
        AssertPixel(bitmap, 190, 115, 0, 0, 255);
        SaveEvidence(pdf, "native-scaled-" + alternateContent);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void ParagraphRelativeGroupFollowsAnchorAfterPaginationAndCoexistsWithImage(bool columns, bool heading) {
        using WordDocument word = WordDocument.Create();
        word.Sections[0].PageSettings.Width = 6000;
        word.Sections[0].PageSettings.Height = 6000;
        if (columns) { word.Sections[0].ColumnCount = 2; word.Sections[0].ColumnsSpace = 120; }
        for (int i = 0; i < (columns ? 40 : 20); i++) word.AddParagraph("Filler " + i).FontSize = 12;
        var paragraph = word.AddParagraph();
        paragraph.AddShapeGroup(new[] {
            new WordShapeGroupItem(WordShapeType.Rectangle, 0, 0, 24, 12) { FillColorHex = "FF0000" },
            new WordShapeGroupItem(WordShapeType.Ellipse, 24, 0, 12, 12) { FillColorHex = "00FF00" }
        }, 10, 0);
        var anchor = paragraph._run!.Descendants<DW.Anchor>().Single();
        anchor.BehindDoc = true;
        anchor.RemoveAllChildren<DW.WrapSquare>();
        anchor.Append(new DW.WrapNone());
        anchor.VerticalPosition!.RelativeFrom = DW.VerticalRelativePositionValues.Paragraph;
        paragraph.AddText("ANCHOR");
        if (heading) paragraph.SetStyle(WordParagraphStyles.Heading1);
        paragraph._paragraph!.ParagraphProperties ??= new W.ParagraphProperties();
        paragraph._paragraph.ParagraphProperties.KeepLines = new W.KeepLines();
        using var imageStream = new MemoryStream(OfficeRasterImageEncoder.Encode(new OfficeRasterImage(4, 4, OfficeColor.Blue), OfficeImageExportFormat.Png));
        var image = paragraph.AddText(string.Empty).InsertImage(imageStream, "fixed.png", 8, 8, WordImageTextWrapping.InFrontOfText);
        image.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
        image.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Page;
        image.HorizontalPositionOffset = 40 * 12700;
        image.VerticalPositionOffset = 10 * 12700;
        var result = word.ToPdfDocumentResult(new WordToPdfOptions { Margins = PdfCore.PageMargins.Uniform(30) });
        var pdf = PdfCore.PdfDocument.Load(result.Value.ToBytes());
        var page = Assert.Single(pdf.Reader.Pages(), item => pdf.Reader.Text(PdfCore.PdfPageSelection.From(item.PageNumber)).Contains("ANCHOR"));
        Assert.True(page.PageNumber > 1);
        Assert.Equal(page.PageNumber, Assert.Single(pdf.Images.Placements()).PageNumber);
        var bitmap = Render(pdf, page.PageNumber);
        // The group starts at the final paragraph top, rather than at the page top.
        AssertPixel(bitmap, 15, 5, 255, 255, 255);
        byte[] pixels = bitmap.GetPixels();
        int redRow = Enumerable.Range(0, bitmap.Height).First(row => pixels[(row * bitmap.Width + 15) * 4] > 240 && pixels[(row * bitmap.Width + 15) * 4 + 1] < 20);
        Assert.True(redRow >= 30);
        var span = Assert.Single(PdfCore.PdfReadDocument.Open(pdf.ToBytes()).Pages[page.PageNumber - 1].GetTextSpans(), item => item.Text.Contains("ANCHOR"));
        Assert.InRange(300D - span.Y - redRow, 0D, span.FontSize + 5D);
        Assert.DoesNotContain(result.Warnings, warning => warning.Code.StartsWith("NativeShapeGroup"));
        SaveEvidence(pdf, "paragraph-group-" + columns + "-heading-" + heading);
    }

    [Fact]
    public void NativeShapeUsesSelectedDrawingInsteadOfInactiveVmlFallback() {
        using WordDocument word = WordDocument.Create();
        var paragraph = word.AddParagraph();
        paragraph.AddShapeDrawing(WordShapeType.Rectangle, 36, 18).FillColorHex = "FF0000";
        var drawing = paragraph._run!.GetFirstChild<W.Drawing>()!;
        var alternate = WrapInAlternateContent(paragraph._run!, drawing, "wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
        alternate.GetFirstChild<AlternateContentFallback>()!.Append(new W.Picture(new V.Rectangle {
            Style = "width:4320pt;height:900pt", FillColor = "blue"
        }));
        Assert.Equal("FF0000", paragraph.Shape!.FillColorHex);
        var pdf = PdfCore.PdfDocument.Load(word.ToPdfDocument().ToBytes());
        AssertPixel(Render(pdf, 1), 80, 80, 255, 0, 0);
    }

    [Fact]
    public void UnsupportedGroupGeometryReportsLossInsteadOfRenderingAnInventedRectangle() {
        using WordDocument word = WordDocument.Create();
        var paragraph = word.AddParagraph("Visible text");
        paragraph.AddShapeGroup(new[] {
            new WordShapeGroupItem(WordShapeType.Rectangle, 0, 0, 36, 18),
            new WordShapeGroupItem(WordShapeType.Ellipse, 36, 0, 18, 18)
        });
        var properties = paragraph._run!.Descendants<Wps.ShapeProperties>().First();
        properties.RemoveAllChildren<A.PresetGeometry>();
        properties.Append(new A.CustomGeometry());
        var result = word.ToPdfDocumentResult();
        var pdf = PdfCore.PdfDocument.Load(result.Value.ToBytes());
        Assert.Contains("Visible text", pdf.Reader.Text());
        Assert.Contains(result.Warnings, item => item.Code == "NativeShapeGroupUnsupported");
    }

    [Fact]
    public void SquareWrappedGroupRetainsItsChildrenAndReportsFlowApproximation() {
        using WordDocument word = WordDocument.Create();
        word.AddParagraph().AddShapeGroup(new[] {
            new WordShapeGroupItem(WordShapeType.Rectangle, 0, 0, 36, 18) { FillColorHex = "FF0000" },
            new WordShapeGroupItem(WordShapeType.Rectangle, 54, 0, 36, 18) { FillColorHex = "0000FF" }
        }, 10, 10);
        word.AddParagraph("Following text");
        var result = word.ToPdfDocumentResult();
        var pdf = PdfCore.PdfDocument.Load(result.Value.ToBytes());
        AssertPixel(Render(pdf, 1), 80, 80, 255, 0, 0);
        AssertPixel(Render(pdf, 1), 135, 80, 0, 0, 255);
        Assert.Contains("Following text", pdf.Reader.Text());
        Assert.Contains(result.Warnings, warning => warning.Code == "NativeShapeGroupFlowed");
    }

    [Fact]
    public void ExplicitNoFillKeepsAnOverlappingGroupChildTransparent() {
        using WordDocument word = WordDocument.Create();
        var paragraph = word.AddParagraph();
        paragraph.AddShapeGroup(new[] {
            new WordShapeGroupItem(WordShapeType.Rectangle, 0, 0, 36, 18) { FillColorHex = "FF0000" },
            new WordShapeGroupItem(WordShapeType.Rectangle, 0, 0, 36, 18) { StrokeColorHex = "0000FF" }
        });
        var properties = paragraph._run!.Descendants<Wps.ShapeProperties>().Last();
        properties.RemoveAllChildren<A.SolidFill>();
        properties.InsertAt(new A.NoFill(), 2);
        var pdf = PdfCore.PdfDocument.Load(word.ToPdfDocument().ToBytes());
        AssertPixel(Render(pdf, 1), 80, 80, 255, 0, 0);
    }

    [Fact]
    public void BehindTextGroupDoesNotCoverEarlierParagraphsAndRespectsGroupZOrder() {
        using WordDocument word = WordDocument.Create();
        word.AddParagraph("FIRST TEXT").FontSize = 18;
        foreach (var item in new[] { (Color: "00FF00", Z: 20U), (Color: "FF0000", Z: 10U) }) {
            var paragraph = word.AddParagraph("Anchor");
            paragraph.AddShapeGroup(new[] {
                new WordShapeGroupItem(WordShapeType.Rectangle, 0, 0, 140, 36) { FillColorHex = item.Color },
                new WordShapeGroupItem(WordShapeType.Rectangle, 140, 0, 10, 36) { FillColorHex = item.Color }
            }, 60, 60);
            var anchor = paragraph._run!.Descendants<DW.Anchor>().Single();
            anchor.BehindDoc = true;
            anchor.RelativeHeight = item.Z;
            anchor.RemoveAllChildren<DW.WrapSquare>();
            anchor.Append(new DW.WrapNone());
        }
        var pdf = PdfCore.PdfDocument.Load(word.ToPdfDocument().ToBytes());
        var bitmap = Render(pdf, 1);
        AssertPixel(bitmap, 65, 65, 0, 255, 0);
        byte[] pixels = bitmap.GetPixels();
        int blackPixels = 0;
        for (int y = 72; y < 95; y++) for (int x = 72; x < 190; x++) {
            int index = (y * bitmap.Width + x) * 4;
            if (pixels[index] < 30 && pixels[index + 1] < 30 && pixels[index + 2] < 30) blackPixels++;
        }
        Assert.True(blackPixels > 30, "Earlier text must remain visible above the later behind-text drawing.");
        SaveEvidence(pdf, "behind-text-groups");
    }

    [Fact]
    public void LegacyGroupOwnsItsImagePlacementWithoutDuplicateFlowImages() {
        string imagePath = Path.Combine(Path.GetTempPath(), "officeimo-group-image-" + Guid.NewGuid().ToString("N") + ".png");
        try {
            File.WriteAllBytes(imagePath, OfficeRasterImageEncoder.Encode(new OfficeRasterImage(4, 4, OfficeColor.Red), OfficeImageExportFormat.Png));
            using WordDocument word = WordDocument.Create();
            var paragraph = word.AddParagraph().AddImageVml(imagePath, 20, 20);
            var shape = paragraph._run!.Descendants<V.Shape>().Single();
            var picture = paragraph._run.GetFirstChild<W.Picture>()!;
            shape.Remove();
            shape.Style = "position:absolute;left:0;top:0;width:20;height:20";
            var group = new V.Group {
                Style = "position:absolute;margin-left:72pt;margin-top:0pt;width:40pt;height:20pt;z-index:-1;mso-position-horizontal-relative:page",
                CoordinateSize = "40,20"
            };
            group.Append(shape, new V.Rectangle { Style = "position:absolute;left:20;top:0;width:20;height:20", FillColor = "blue", Stroked = false });
            picture.Append(group);
            var result = word.ToPdfDocumentResult();
            var pdf = PdfCore.PdfDocument.Load(result.Value.ToBytes());
            var image = Assert.Single(pdf.Images.Placements());
            Assert.Equal(72D, image.X, 3);
            Assert.Equal(20D, image.Width, 3);
            Assert.Single(pdf.Reader.Pages());
            AssertPixel(Render(pdf, 1), 80, 80, 255, 0, 0);
            AssertPixel(Render(pdf, 1), 100, 80, 0, 0, 255);
            Assert.DoesNotContain(result.Warnings, warning => warning.Code == "NativeShapeGroupUnsupported");
        } finally { File.Delete(imagePath); }
    }

    [Fact]
    public void GroupLabelStaysInsideItsShapeAndDoesNotReplaceOrDuplicateBodyText() {
        using WordDocument word = WordDocument.Create();
        var paragraph = word.AddParagraph("Body text");
        var group = new V.Group {
            Style = "position:absolute;margin-left:72pt;margin-top:24pt;width:160pt;height:30pt;z-index:-1;mso-position-horizontal-relative:page",
            CoordinateSize = "160,30"
        };
        var rectangle = new V.Rectangle {
            Style = "position:absolute;left:0;top:0;width:160;height:30", FillColor = "red", Stroked = false
        };
        rectangle.Append(new V.TextBox(new W.TextBoxContent(new W.Paragraph(new W.Run(new W.Text("GROUP LABEL"))))));
        group.Append(rectangle);
        paragraph._run!.Append(new W.Picture(group));
        Assert.Null(paragraph.TextBox);
        Assert.Equal("Body text", paragraph.Text);
        var result = word.ToPdfDocumentResult();
        var pdf = PdfCore.PdfDocument.Load(result.Value.ToBytes());
        string text = pdf.Reader.Text();
        Assert.Contains("Body text", text);
        Assert.Equal(1, text.Split(new[] { "GROUP LABEL" }, StringSplitOptions.None).Length - 1);
        AssertPixel(Render(pdf, 1), 220, 120, 255, 0, 0);
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "NativeShapeGroupUnsupported");
        SaveEvidence(pdf, "group-label");
    }

    private static AlternateContent WrapInAlternateContent(W.Run run, W.Drawing drawing, string prefix, string ns) {
        drawing.Remove();
        var choice = new AlternateContentChoice { Requires = prefix };
        choice.AddNamespaceDeclaration(prefix, ns);
        choice.Append(drawing);
        var alternate = new AlternateContent(choice, new AlternateContentFallback());
        run.Append(alternate);
        return alternate;
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyGroupsPreserveDeclaredStackOrderRatherThanParagraphOrder(bool whitespaceOnly) {
        using WordDocument word = WordDocument.Create();
        foreach (var item in new[] { (Z: -1, Color: "green"), (Z: -2, Color: "red") }) {
            var group = new V.Group {
                Style = $"position:absolute;margin-left:72pt;margin-top:72pt;width:40pt;height:20pt;z-index:{item.Z};mso-position-horizontal-relative:page;mso-position-vertical-relative:page",
                CoordinateSize = "40,20"
            };
            group.Append(new V.Rectangle { Style = "position:absolute;left:0;top:0;width:40;height:20", FillColor = item.Color, Stroked = false });
            word.AddParagraph(whitespaceOnly ? " " : "X")._run!.Append(new W.Picture(group));
        }
        var result = word.ToPdfDocumentResult();
        var pdf = PdfCore.PdfDocument.Load(result.Value.ToBytes());
        SaveEvidence(pdf, "legacy-stack-order-" + whitespaceOnly);
        AssertPixel(Render(pdf, 1), 100, 80, 0, 128, 0);
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "NativeShapeGroupUnsupported");
    }

    [Theory]
    [InlineData(0, 500)]
    [InlineData(1400, 200)]
    public void ParagraphGroupFollowsTextClearedBelowFloatingTable(int tableOffset, int pageHeight) {
        using WordDocument word = WordDocument.Create();
        var table = word.AddTable(1, 1);
        table.LayoutMode = WordTableLayoutMode.Fixed;
        table.WidthType = WordTableWidthUnit.Dxa;
        table.Width = 6400;
        table.Rows[0].Height = 900;
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Floating";
        table._tableProperties!.TablePositionProperties = new W.TablePositionProperties {
            HorizontalAnchor = W.HorizontalAnchorValues.Margin, VerticalAnchor = W.VerticalAnchorValues.Text,
            TablePositionY = tableOffset, BottomFromText = 180
        };
        var paragraph = word.AddParagraph();
        paragraph.AddShapeGroup(new[] {
            new WordShapeGroupItem(WordShapeType.Rectangle, 0, 0, 24, 12) { FillColorHex = "FF0000" },
            new WordShapeGroupItem(WordShapeType.Ellipse, 24, 0, 12, 12) { FillColorHex = "00FF00" }
        }, 10, 0);
        var anchor = paragraph._run!.Descendants<DW.Anchor>().Single();
        anchor.BehindDoc = true;
        anchor.RemoveAllChildren<DW.WrapSquare>();
        anchor.Append(new DW.WrapNone());
        anchor.VerticalPosition!.RelativeFrom = DW.VerticalRelativePositionValues.Paragraph;
        paragraph.AddText("ANCHOR " + string.Join(" ", Enumerable.Range(0, 40).Select(index => "word" + index)));
        var result = word.ToPdfDocumentResult(new WordToPdfOptions {
            IncludePageNumbers = false, PageSize = new PdfCore.PageSize(400, pageHeight), Margins = PdfCore.PageMargins.Uniform(40)
        });
        var pdf = PdfCore.PdfDocument.Load(result.Value.ToBytes());
        var page = Assert.Single(pdf.Reader.Pages(), item => pdf.Reader.Text(PdfCore.PdfPageSelection.From(item.PageNumber)).Contains("ANCHOR"));
        var bitmap = Render(pdf, page.PageNumber);
        byte[] pixels = bitmap.GetPixels();
        int redRow = Enumerable.Range(0, bitmap.Height).First(row => pixels[(row * bitmap.Width + 15) * 4] > 250 && pixels[(row * bitmap.Width + 15) * 4 + 1] < 5);
        var span = Assert.Single(PdfCore.PdfReadDocument.Open(pdf.ToBytes()).Pages[page.PageNumber - 1].GetTextSpans(), item => item.Text.Contains("ANCHOR"));
        Assert.InRange(pageHeight - span.Y - redRow, 0D, span.FontSize + 5D);
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "NativeShapeGroupUnsupported");
        SaveEvidence(pdf, "floating-table-group-" + tableOffset);
    }

    [Theory]
    [InlineData(WordListStyle.Bulleted)]
    [InlineData(WordListStyle.Numbered)]
    public void ListParagraphKeepsItsGroupedDrawingAndMarker(WordListStyle listStyle) {
        using WordDocument word = WordDocument.Create();
        var paragraph = word.AddList(listStyle).AddItem("List anchor");
        paragraph.AddShapeGroup(new[] {
            new WordShapeGroupItem(WordShapeType.Rectangle, 0, 0, 24, 12) { FillColorHex = "FF0000" },
            new WordShapeGroupItem(WordShapeType.Ellipse, 24, 0, 12, 12) { FillColorHex = "00FF00" }
        }, 10, 0);
        var anchor = paragraph._run!.Descendants<DW.Anchor>().Single();
        anchor.BehindDoc = true;
        anchor.RemoveAllChildren<DW.WrapSquare>();
        anchor.Append(new DW.WrapNone());
        anchor.VerticalPosition!.RelativeFrom = DW.VerticalRelativePositionValues.Paragraph;
        var result = word.ToPdfDocumentResult();
        var pdf = PdfCore.PdfDocument.Load(result.Value.ToBytes());
        AssertPixel(Render(pdf, 1), 15, 75, 255, 0, 0);
        Assert.Contains("List anchor", pdf.Reader.Text());
        if (listStyle == WordListStyle.Numbered) Assert.Contains("1.", pdf.Reader.Text());
        else Assert.Contains("•", pdf.Reader.Text());
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "NativeShapeGroupUnsupported");
        SaveEvidence(pdf, "list-group-" + listStyle);
    }

    private static OfficeRasterImage Render(PdfCore.PdfDocument pdf, int page) {
        byte[] bytes = pdf.Render.Pages(PdfCore.PdfPageSelection.From(page), new PdfCore.PdfPageRenderOptions { Dpi = 72 })[0].Bytes!;
        Assert.True(OfficePngReader.TryDecode(bytes, out var bitmap));
        return bitmap!;
    }

    private static void AssertPixel(OfficeRasterImage image, int x, int y, byte red, byte green, byte blue) {
        byte[] pixels = image.GetPixels();
        int offset = (y * image.Width + x) * 4;
        Assert.InRange(pixels[offset], Math.Max(0, red - 5), Math.Min(255, red + 5));
        Assert.InRange(pixels[offset + 1], Math.Max(0, green - 5), Math.Min(255, green + 5));
        Assert.InRange(pixels[offset + 2], Math.Max(0, blue - 5), Math.Min(255, blue + 5));
    }

    private static void SaveEvidence(PdfCore.PdfDocument pdf, string name) {
        if (Environment.GetEnvironmentVariable("OFFICEIMO_PDF_VISUAL_OUTPUT") is not { Length: > 0 } output) return;
        Directory.CreateDirectory(output);
        pdf.Save(Path.Combine(output, name + ".pdf"));
        foreach (var page in pdf.Render.Pages(options: new PdfCore.PdfPageRenderOptions { Dpi = 96 }))
            File.WriteAllBytes(Path.Combine(output, name + "-" + page.PageNumber + ".png"), page.Bytes!);
    }
}
