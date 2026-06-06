using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfDocumentCanvasTests {
    [Fact]
    public void CanvasText_RendersAtFixedTopLeftCoordinatesWithoutMovingFlowContent() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 160,
                MarginLeft = 24,
                MarginRight = 24,
                MarginTop = 72,
                MarginBottom = 24,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Text("CanvasTitle", 24, 20, 120, 24, fontSize: 12, color: PdfColor.FromRgb(20, 90, 160)))
            .Paragraph(paragraph => paragraph.Text("FlowAfterCanvas"))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double canvasY = FindWordStartY(page, "CanvasTitle");
        double flowY = FindWordStartY(page, "FlowAfterCanvas");

        Assert.InRange(FindWordStartX(page, "CanvasTitle"), 23D, 26D);
        Assert.True(canvasY > flowY, "Canvas text should render above the flow paragraph when placed near the page top.");
        Assert.InRange(flowY, 77D, 91D);
    }

    [Fact]
    public void CanvasShape_RendersRectangleAtFixedTopLeftCoordinates() {
        var shape = OfficeShape.Rectangle(60, 20);
        shape.FillColor = PdfColor.FromRgb(230, 245, 255).ToOfficeColor();
        shape.StrokeColor = PdfColor.FromRgb(15, 98, 160).ToOfficeColor();
        shape.StrokeWidth = 1.25D;

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 160,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Shape(shape, 30, 40))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("30 100 60 20 re", content, StringComparison.Ordinal);
        Assert.Contains("1.25 w", content, StringComparison.Ordinal);
        Assert.Contains(" B", content, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasShape_WithRotation_RendersUsingSharedShapeTransform() {
        var shape = OfficeShape.Rectangle(40, 20);
        shape.FillColor = PdfColor.FromRgb(230, 245, 255).ToOfficeColor();
        shape.StrokeColor = PdfColor.FromRgb(15, 98, 160).ToOfficeColor();
        shape.StrokeWidth = 1D;

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 160,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Shape(shape, 30, 40, rotationAngle: 90D))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("0 -1 -1 0 60 130 cm", content, StringComparison.Ordinal);
        Assert.Contains("0 0 40 20 re", content, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasDrawing_RendersSharedVectorSceneInsideFixedFrame() {
        var drawing = new OfficeDrawing(50, 20);
        var shape = OfficeShape.Rectangle(20, 10);
        shape.FillColor = PdfColor.FromRgb(230, 245, 255).ToOfficeColor();
        shape.StrokeColor = PdfColor.FromRgb(15, 98, 160).ToOfficeColor();
        shape.StrokeWidth = 1D;
        drawing.AddShape(shape, 5, 5);
        drawing.AddText(
            "SceneText",
            8,
            4,
            36,
            10,
            new OfficeFontInfo("Aptos", 6D, OfficeFontStyle.Bold),
            PdfColor.FromRgb(31, 78, 121).ToOfficeColor(),
            OfficeTextAlignment.Center);

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 140,
                MarginLeft = 0,
                MarginRight = 0,
                MarginTop = 0,
                MarginBottom = 0,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Drawing(drawing, 20, 30, 100, 40))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("2 0 0 2 20 70 cm", content, StringComparison.Ordinal);
        Assert.Contains("5 5 20 10 re", content, StringComparison.Ordinal);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("", pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("SceneText", text, StringComparison.Ordinal);
    }

    [Fact]
    public void FlowDrawing_RendersSharedVectorSceneText() {
        var drawing = new OfficeDrawing(120, 36)
            .AddText(
                "FlowSceneText",
                8,
                8,
                104,
                16,
                new OfficeFontInfo("Aptos", 10D, OfficeFontStyle.Bold),
                PdfColor.FromRgb(31, 78, 121).ToOfficeColor(),
                OfficeTextAlignment.Center);

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 160,
                MarginLeft = 24,
                MarginRight = 24,
                MarginTop = 24,
                MarginBottom = 24,
                CompressContentStreams = false
            })
            .Drawing(drawing, PdfAlign.Left)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("", pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("FlowSceneText", text, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasImage_RendersImageAtFixedTopLeftCoordinatesWithAltText() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 160,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Image(CreateMinimalRgbPng(), 30, 40, 60, 30, alternativeText: "Canvas logo"))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("60 0 0 30 30 90 cm", content, StringComparison.Ordinal);
        Assert.Contains("/Im1 Do", content, StringComparison.Ordinal);
        Assert.Contains("/Figure << /Alt <43616E766173206C6F676F> >> BDC", content, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasImage_WithSourceCrop_ClipsAndOffsetsImageInsideDeclaredFrame() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 160,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Image(CreateMinimalRgbPng(), 40, 50, 60, 30, new PdfImageStyle {
                SourceCrop = new PdfImageSourceCrop(left: 0.5D, top: 0D, right: 0D, bottom: 0D)
            }, linkUri: "https://evotec.xyz/cropped"))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("120 0 0 30 -20 80 cm", content, StringComparison.Ordinal);
        Assert.Contains("0.5 0 0.5 1 re", content, StringComparison.Ordinal);
        Assert.Contains("/Im1 Do", content, StringComparison.Ordinal);
        PdfLinkAnnotation link = Assert.Single(PdfInspector.Inspect(bytes).LinkAnnotations);
        AssertClose(40D, link.X1);
        AssertClose(80D, link.Y1);
        AssertClose(100D, link.X2);
        AssertClose(110D, link.Y2);
    }

    [Fact]
    public void CanvasImage_WithRotation_RendersImageAroundDeclaredFrameCenter() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 160,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Image(CreateMinimalRgbPng(), 30, 40, 60, 30, rotationAngle: 90D))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("0 60 -30 0 75 75 cm", content, StringComparison.Ordinal);
        Assert.Contains("/Im1 Do", content, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasImage_RendersBeforeFollowingShapeInCanvasOrder() {
        var shape = OfficeShape.Rectangle(70, 35);
        shape.FillColor = PdfColor.FromRgb(255, 255, 255).ToOfficeColor();
        shape.StrokeColor = PdfColor.FromRgb(15, 98, 160).ToOfficeColor();
        shape.StrokeWidth = 1D;

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 160,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas
                .Image(CreateMinimalRgbPng(), 30, 40, 60, 30)
                .Shape(shape, 25, 35))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int imageDraw = content.IndexOf("/Im1 Do", StringComparison.Ordinal);
        int shapeDraw = content.IndexOf("25 90 70 35 re", StringComparison.Ordinal);

        Assert.True(imageDraw >= 0, "Expected the canvas image draw command to be present.");
        Assert.True(shapeDraw >= 0, "Expected the following canvas shape draw command to be present.");
        Assert.True(imageDraw < shapeDraw, "Canvas images should be painted in declared order instead of being appended after later canvas items.");
    }

    [Fact]
    public void CanvasTextBox_RendersStyledBoxAndClippedTextAtFixedCoordinates() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 260,
                PageHeight = 180,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.TextBox("Premium text box", 30, 40, 140, 50, new PdfCanvasTextBoxStyle {
                Background = PdfColor.FromRgb(245, 250, 255),
                BorderColor = PdfColor.FromRgb(24, 96, 160),
                BorderWidth = 1.5D,
                PaddingX = 8D,
                PaddingY = 6D,
                FontSize = 10D,
                TextColor = PdfColor.FromRgb(20, 40, 70),
                Align = PdfAlign.Center
            }))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("30 90 140 50 re", content, StringComparison.Ordinal);
        Assert.Contains("1.5 w", content, StringComparison.Ordinal);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        Assert.InRange(FindWordStartX(page, "Premium"), 61D, 91D);
        Assert.InRange(FindWordStartY(page, "Premium"), 120D, 135D);
    }

    [Fact]
    public void CanvasDiagnosticOverloads_PreservePreviousClrSignatures() {
        Assert.NotNull(typeof(PdfPageCanvas).GetMethod(nameof(PdfPageCanvas.TextBox), new[] {
            typeof(string),
            typeof(double),
            typeof(double),
            typeof(double),
            typeof(double),
            typeof(PdfCanvasTextBoxStyle),
            typeof(double)
        }));
        Assert.NotNull(typeof(PdfPageCanvas).GetMethod(nameof(PdfPageCanvas.TextBox), new[] {
            typeof(IEnumerable<TextRun>),
            typeof(double),
            typeof(double),
            typeof(double),
            typeof(double),
            typeof(PdfCanvasTextBoxStyle),
            typeof(double)
        }));
        Assert.NotNull(typeof(PdfPageCanvas).GetMethod(nameof(PdfPageCanvas.Table), new[] {
            typeof(IEnumerable<string[]>),
            typeof(double),
            typeof(double),
            typeof(double),
            typeof(double),
            typeof(PdfTableStyle),
            typeof(double)
        }));
        Assert.NotNull(typeof(PdfPageCanvas).GetMethod(nameof(PdfPageCanvas.Table), new[] {
            typeof(IEnumerable<PdfTableCell[]>),
            typeof(double),
            typeof(double),
            typeof(double),
            typeof(double),
            typeof(PdfTableStyle),
            typeof(double)
        }));
    }

    [Fact]
    public void CanvasTextBox_ReportsClippedContentDuringRender() {
        PdfLayoutDiagnostic? diagnostic = null;
        PdfDocument document = PdfDocument.Create(new PdfOptions {
                PageWidth = 180,
                PageHeight = 120,
                MarginLeft = 0,
                MarginRight = 0,
                MarginTop = 0,
                MarginBottom = 0,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.TextBox(
                "One two three four five six seven eight nine ten eleven twelve",
                20,
                20,
                80,
                18,
                new PdfCanvasTextBoxStyle {
                    Background = null,
                    BorderColor = null,
                    FontSize = 12D,
                    PaddingX = 0D,
                    PaddingY = 0D
                },
                rotationAngle: 0D,
                diagnosticHandler: item => diagnostic = item));

        Assert.Null(diagnostic);

        document.ToBytes();

        Assert.NotNull(diagnostic);
        Assert.Equal(PdfLayoutDiagnosticKind.ClippedContent, diagnostic!.Kind);
        Assert.Equal("PdfCanvasTextBox", diagnostic.Source);
        Assert.True(diagnostic.HasBounds);
        Assert.Equal(20D, diagnostic.X);
        Assert.Equal(20D, diagnostic.Y);
        Assert.Equal(80D, diagnostic.Width);
        Assert.Equal(18D, diagnostic.Height);
    }

    [Fact]
    public void CanvasTextBox_WithAsymmetricPadding_UsesIndividualEdges() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 260,
                PageHeight = 180,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.TextBox("Asymmetric", 30, 40, 140, 50, new PdfCanvasTextBoxStyle {
                Background = null,
                BorderColor = null,
                PaddingLeft = 20D,
                PaddingRight = 4D,
                PaddingTop = 6D,
                PaddingBottom = 2D,
                FontSize = 10D
            }))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("50 92 116 42 re", content, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasTextBox_UsesConfiguredVerticalAlignmentInsideFixedFrame() {
        static PdfCanvasTextBoxStyle Style(PdfVerticalAlign verticalAlign) =>
            new PdfCanvasTextBoxStyle {
                Background = null,
                BorderColor = null,
                PaddingX = 0D,
                PaddingY = 0D,
                FontSize = 10D,
                LineHeight = 12D,
                VerticalAlign = verticalAlign
            };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 360,
                PageHeight = 200,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas
                .TextBox("TopAlign", 20, 30, 90, 90, Style(PdfVerticalAlign.Top))
                .TextBox("MiddleAlign", 130, 30, 90, 90, Style(PdfVerticalAlign.Middle))
                .TextBox("BottomAlign", 240, 30, 90, 90, Style(PdfVerticalAlign.Bottom)))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double topY = FindWordStartY(page, "TopAlign");
        double middleY = FindWordStartY(page, "MiddleAlign");
        double bottomY = FindWordStartY(page, "BottomAlign");

        Assert.True(topY > middleY + 30D, $"Expected middle-aligned text to render lower than top-aligned text. Top: {topY:0.##}, middle: {middleY:0.##}.");
        Assert.True(middleY > bottomY + 30D, $"Expected bottom-aligned text to render lower than middle-aligned text. Middle: {middleY:0.##}, bottom: {bottomY:0.##}.");
    }

    [Fact]
    public void CanvasTextBox_RejectsInvalidVerticalAlignment() {
        ArgumentException ex = Assert.Throws<ArgumentException>(() =>
            new PdfCanvasTextBoxStyle {
                VerticalAlign = (PdfVerticalAlign)99
            });

        Assert.Contains("Canvas text box vertical alignment must be Top, Middle, or Bottom.", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasTextBox_RendersBackgroundBeforeTextAndFollowingShape() {
        var shape = OfficeShape.Rectangle(25, 20);
        shape.FillColor = PdfColor.FromRgb(255, 255, 255).ToOfficeColor();
        shape.StrokeColor = PdfColor.FromRgb(30, 64, 175).ToOfficeColor();
        shape.StrokeWidth = 1D;

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 260,
                PageHeight = 180,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas
                .TextBox("Layered", 30, 40, 120, 42, new PdfCanvasTextBoxStyle {
                    Background = PdfColor.FromRgb(250, 250, 250),
                    BorderColor = PdfColor.FromRgb(75, 85, 99),
                    BorderWidth = 1D,
                    FontSize = 10D
                })
                .Shape(shape, 40, 48))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int textBoxDraw = content.IndexOf("30 98 120 42 re", StringComparison.Ordinal);
        int textStart = content.IndexOf("BT", textBoxDraw, StringComparison.Ordinal);
        int followingShapeDraw = content.IndexOf("40 112 25 20 re", StringComparison.Ordinal);

        Assert.True(textBoxDraw >= 0, "Expected the text box background rectangle to be present.");
        Assert.True(textStart > textBoxDraw, "Expected text box text to render after its own background.");
        Assert.True(followingShapeDraw > textStart, "Expected later canvas items to render after the complete text box.");
    }

    [Fact]
    public void CanvasTable_RendersFixedFrameStyledCellsAndText() {
        var style = new PdfTableStyle {
            HeaderRowCount = 1,
            RowStripeFill = null,
            ColumnWidthPoints = new System.Collections.Generic.List<double?> { 70D, 50D },
            RowMinHeights = new System.Collections.Generic.List<double?> { 24D, 36D },
            CellFills = new System.Collections.Generic.Dictionary<(int Row, int Column), PdfColor> {
                [(1, 1)] = PdfColor.FromRgb(230, 245, 255)
            },
            CellPaddings = new System.Collections.Generic.Dictionary<(int Row, int Column), PdfCellPadding> {
                [(1, 1)] = new PdfCellPadding { Left = 8D, Right = 8D, Top = 4D, Bottom = 4D }
            },
            CellAlignments = new System.Collections.Generic.Dictionary<(int Row, int Column), PdfColumnAlign> {
                [(1, 1)] = PdfColumnAlign.Center
            },
            CellVerticalAlignments = new System.Collections.Generic.Dictionary<(int Row, int Column), PdfCellVerticalAlign> {
                [(1, 1)] = PdfCellVerticalAlign.Middle
            }
        };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Table(new[] {
                new[] { "Name", "Score" },
                new[] { "OfficeIMO", "99" }
            }, 30, 30, 120, 60, style))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("30 90 120 60 re", raw, StringComparison.Ordinal);
        Assert.Contains("100 150 m", raw, StringComparison.Ordinal);
        Assert.Contains("100 90 l", raw, StringComparison.Ordinal);
        Assert.Contains("30 126 m", raw, StringComparison.Ordinal);
        Assert.Contains("150 126 l", raw, StringComparison.Ordinal);
        Assert.Contains("100 90 50 36 re", raw, StringComparison.Ordinal);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("", pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("Name", text, StringComparison.Ordinal);
        Assert.Contains("OfficeIMO", text, StringComparison.Ordinal);
        Assert.Contains("99", text, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasTable_ReportsClippedCellContentDuringRender() {
        var diagnostics = new List<PdfLayoutDiagnostic>();
        PdfDocument document = PdfDocument.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 140,
                MarginLeft = 0,
                MarginRight = 0,
                MarginTop = 0,
                MarginBottom = 0,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Table(
                new[] {
                    new[] {
                        PdfTableCell.TextCell("One two three four five six seven eight nine ten eleven twelve")
                    }
                },
                20,
                20,
                80,
                22,
                new PdfTableStyle {
                    RowMinHeights = new List<double?> { 22D },
                    ColumnWidthPoints = new List<double?> { 80D },
                    CellPaddings = new Dictionary<(int Row, int Column), PdfCellPadding> {
                        [(0, 0)] = new PdfCellPadding { Left = 2D, Right = 2D, Top = 2D, Bottom = 2D }
                    }
                },
                rotationAngle: 0D,
                diagnosticHandler: diagnostics.Add));

        Assert.Empty(diagnostics);

        document.ToBytes();

        PdfLayoutDiagnostic diagnostic = Assert.Single(diagnostics);
        Assert.Equal(PdfLayoutDiagnosticKind.ClippedContent, diagnostic.Kind);
        Assert.Equal("PdfCanvasTableCell", diagnostic.Source);
        Assert.True(diagnostic.HasBounds);
        Assert.Equal(20D, diagnostic.X);
        Assert.Equal(20D, diagnostic.Y);
        Assert.Equal(80D, diagnostic.Width);
        Assert.Equal(22D, diagnostic.Height);
    }

    [Fact]
    public void CanvasTable_WithRotation_RendersInsideRotatedFrame() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Table(new[] {
                new[] { "Name", "Score" },
                new[] { "OfficeIMO", "99" }
            }, 30, 30, 120, 60, rotationAngle: 90D))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        int transform = raw.IndexOf("0 1 -1 0 210 30 cm", StringComparison.Ordinal);
        int tableRect = raw.IndexOf("30 90 120 60 re", StringComparison.Ordinal);

        Assert.True(transform >= 0, "Expected a rotation matrix around the declared table frame center.");
        Assert.True(tableRect > transform, "Expected table geometry to render inside the rotated frame.");
    }

    [Fact]
    public void CanvasTable_RendersRichCellImagesAndFormControls() {
        var rows = new[] {
            new[] {
                PdfTableCell.WithImages(
                    "Assets",
                    new[] { new PdfTableCellImage(CreateMinimalRgbPng(), 12, 12) },
                    checkBoxes: new[] { new PdfTableCellCheckBox("Canvas.Approved", isChecked: true, size: 10) },
                    formFields: new[] { PdfTableCellFormField.TextField("Canvas.Owner", "Ada", width: 44, height: 12, fontSize: 8) })
            }
        };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Table(rows, 24, 24, 120, 86, new PdfTableStyle {
                RowMinHeights = new System.Collections.Generic.List<double?> { 86D },
                CellPaddingX = 6D,
                CellPaddingY = 6D
            }))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/Im1 Do", raw, StringComparison.Ordinal);

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        Assert.Contains(info.FormFields, field => field.Name == "Canvas.Approved" && field.IsCheckBox && field.Value == "Yes");
        Assert.Contains(info.FormFields, field => field.Name == "Canvas.Owner" && field.IsTextField && field.Value == "Ada");
    }

    [Fact]
    public void CanvasTable_WithRotation_RotatesRichCellImagesAndFormControls() {
        var rows = new[] {
            new[] {
                PdfTableCell.WithImages(
                    "Assets",
                    new[] { new PdfTableCellImage(CreateMinimalRgbPng(), 12, 12) },
                    checkBoxes: new[] { new PdfTableCellCheckBox("Canvas.Rotated", isChecked: true, size: 10) },
                    formFields: new[] { PdfTableCellFormField.TextField("Canvas.RotatedOwner", "Ada", width: 44, height: 12, fontSize: 8) })
            }
        };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Table(rows, 24, 24, 120, 86, new PdfTableStyle {
                RowMinHeights = new System.Collections.Generic.List<double?> { 86D },
                CellPaddingX = 6D,
                CellPaddingY = 6D
            }, rotationAngle: 90D))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0 12 -12 0", raw, StringComparison.Ordinal);

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        Assert.Contains(info.FormFields, field => field.Name == "Canvas.Rotated" && field.IsCheckBox && field.Value == "Yes");
        Assert.Contains(info.FormFields, field => field.Name == "Canvas.RotatedOwner" && field.IsTextField && field.Value == "Ada");
    }

    [Fact]
    public void CanvasClip_ClipsDeferredTableImagesAndFormControls() {
        var rows = new[] {
            new[] {
                PdfTableCell.WithImages(
                    string.Empty,
                    new[] { new PdfTableCellImage(CreateMinimalRgbPng(), 40, 40) },
                    formFields: new[] { PdfTableCellFormField.TextField("Canvas.ClippedOwner", "Ada", width: 44, height: 40, fontSize: 8) })
            }
        };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Clip(50, 24, 34, 86, clipped => clipped.Table(rows, 24, 24, 120, 86, new PdfTableStyle {
                RowMinHeights = new System.Collections.Generic.List<double?> { 86D },
                CellPaddingX = 6D,
                CellPaddingY = 6D
            })))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/Im1 Do", raw, StringComparison.Ordinal);
        Assert.Contains("50 90 20 40 re W", raw, StringComparison.Ordinal);

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfFormField field = Assert.Single(info.FormFields, item => item.Name == "Canvas.ClippedOwner");
        PdfFormWidget widget = Assert.Single(field.Widgets);
        AssertClose(50D, widget.X1);
        AssertClose(50D, widget.Y1);
        AssertClose(74D, widget.X2);
        AssertClose(88D, widget.Y2);
    }

    [Fact]
    public void CanvasClip_PreservesInlineImageClipPathInsideFrame() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Clip(20, 20, 100, 80, clipped => clipped.Image(CreateMinimalRgbPng(), 30, 30, 40, 40, new PdfImageStyle {
                ClipPath = OfficeClipPath.Rectangle(20, 20)
            })))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("20 60 100 80 re W", raw, StringComparison.Ordinal);
        Assert.Contains("30 110 20 20 re W", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("30 90 40 40 re W", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasTable_SkipsVerticalGridDividersInsideMergedCells() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Table(new[] {
                new[] { PdfTableCell.Span("Merged", 2) },
                new[] { PdfTableCell.TextCell("Left"), PdfTableCell.TextCell("Right") }
            }, 30, 30, 120, 60))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("30 90 120 60 re", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("90 150 m", raw, StringComparison.Ordinal);
        Assert.Contains("90 120 m", raw, StringComparison.Ordinal);
        Assert.Contains("90 90 l", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasTable_RowSpanSkipsContinuationRowStripeFill() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Table(new[] {
                new[] { PdfTableCell.Merge("Span", rowSpan: 2), PdfTableCell.TextCell("Top") },
                new[] { PdfTableCell.TextCell("Bottom") }
            }, 30, 30, 120, 60, new PdfTableStyle {
                HeaderRowCount = 0,
                RowStripeFill = PdfColor.FromRgb(220, 235, 250),
                ColumnWidthPoints = new System.Collections.Generic.List<double?> { 60D, 60D },
                RowMinHeights = new System.Collections.Generic.List<double?> { 30D, 30D }
            }))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);

        Assert.DoesNotContain("30 90 120 30 re", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("30 90 60 30 re", raw, StringComparison.Ordinal);
        Assert.Contains("90 90 60 30 re", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasTextBox_WithRotation_RendersBoxAndTextInsideRotatedGroup() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 260,
                PageHeight = 180,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.TextBox("Rotated box", 30, 40, 120, 42, new PdfCanvasTextBoxStyle {
                Background = PdfColor.FromRgb(250, 250, 250),
                BorderColor = PdfColor.FromRgb(75, 85, 99),
                BorderWidth = 1D,
                FontSize = 10D
            }, rotationAngle: 90D))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int transform = content.IndexOf("0 1 -1 0 209 29 cm", StringComparison.Ordinal);
        int rectangle = content.IndexOf("30 98 120 42 re", StringComparison.Ordinal);
        int textStart = content.IndexOf("BT", rectangle, StringComparison.Ordinal);

        Assert.True(transform >= 0, "Expected a rotation matrix around the declared text box frame center.");
        Assert.True(rectangle > transform, "Expected the text box geometry to render inside the rotated group.");
        Assert.True(textStart > rectangle, "Expected text to render after the rotated text box background.");
    }

    [Fact]
    public void CanvasTextBox_InvalidPadding_ThrowsClearDiagnostic() {
        ArgumentException ex = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(new PdfOptions {
                    PageWidth = 100,
                    PageHeight = 100
                })
                .Canvas(canvas => canvas.TextBox("Bad", 0, 0, 20, 10, new PdfCanvasTextBoxStyle {
                    PaddingY = 5D
                })));

        Assert.Contains("Canvas text box padding must leave a positive text area.", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasRotation_NonFiniteAngle_ThrowsClearDiagnostic() {
        var shape = OfficeShape.Rectangle(10, 10);
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfDocument.Create()
                .Canvas(canvas => canvas.TextBox("Bad", 0, 0, 10, 10, rotationAngle: double.NegativeInfinity)));

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfDocument.Create()
                .Canvas(canvas => canvas.Shape(shape, 0, 0, rotationAngle: double.NaN)));

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfDocument.Create()
                .Canvas(canvas => canvas.Image(CreateMinimalRgbPng(), 0, 0, 10, 10, rotationAngle: double.PositiveInfinity)));

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfDocument.Create()
                .Canvas(canvas => canvas.Table(new[] { new[] { "Bad" } }, 0, 0, 10, 10, rotationAngle: double.NaN)));
    }

    [Fact]
    public void CanvasTextBox_WithRotationAndLinkedRun_RotatesLinkAnnotationBounds() {
        PdfOptions options = CreateCanvasOptions();
        const string uri = "https://evotec.xyz/canvas-textbox";
        var style = new PdfCanvasTextBoxStyle {
            FontSize = 10D
        };

        byte[] flatBytes = PdfDocument.Create(options)
            .Canvas(canvas => canvas.TextBox(new[] {
                TextRun.Link("Linked", uri)
            }, 30, 40, 120, 42, style))
            .ToBytes();
        byte[] rotatedBytes = PdfDocument.Create(options)
            .Canvas(canvas => canvas.TextBox(new[] {
                TextRun.Link("Linked", uri)
            }, 30, 40, 120, 42, style, rotationAngle: 90D))
            .ToBytes();

        PdfLinkAnnotation flatLink = Assert.Single(PdfInspector.Inspect(flatBytes).LinkAnnotations);
        PdfLinkAnnotation rotatedLink = Assert.Single(PdfInspector.Inspect(rotatedBytes).LinkAnnotations);
        var expected = RotateRectangle(flatLink, 30, 98, 120, 42, 90D);

        Assert.Equal(uri, rotatedLink.Uri);
        AssertClose(expected.X1, rotatedLink.X1);
        AssertClose(expected.Y1, rotatedLink.Y1);
        AssertClose(expected.X2, rotatedLink.X2);
        AssertClose(expected.Y2, rotatedLink.Y2);
    }

    [Fact]
    public void CanvasImage_WithRotationAndLink_RotatesLinkAnnotationBounds() {
        PdfOptions options = CreateCanvasOptions();
        const string uri = "https://evotec.xyz/canvas-image";

        byte[] flatBytes = PdfDocument.Create(options)
            .Canvas(canvas => canvas.Image(CreateMinimalRgbPng(), 30, 40, 60, 30, linkUri: uri))
            .ToBytes();
        byte[] rotatedBytes = PdfDocument.Create(options)
            .Canvas(canvas => canvas.Image(CreateMinimalRgbPng(), 30, 40, 60, 30, linkUri: uri, rotationAngle: 90D))
            .ToBytes();

        PdfLinkAnnotation flatLink = Assert.Single(PdfInspector.Inspect(flatBytes).LinkAnnotations);
        PdfLinkAnnotation rotatedLink = Assert.Single(PdfInspector.Inspect(rotatedBytes).LinkAnnotations);
        var expected = RotateRectangle(flatLink, 30, 110, 60, 30, 90D);

        Assert.Equal(uri, rotatedLink.Uri);
        AssertClose(expected.X1, rotatedLink.X1);
        AssertClose(expected.Y1, rotatedLink.Y1);
        AssertClose(expected.X2, rotatedLink.X2);
        AssertClose(expected.Y2, rotatedLink.Y2);
    }

    [Fact]
    public void CanvasImage_AppliesFitAfterSourceCrop() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 100,
                PageHeight = 100,
                MarginLeft = 0,
                MarginRight = 0,
                MarginTop = 0,
                MarginBottom = 0,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Image(
                CreateMinimalRgbPng(),
                0,
                0,
                100,
                100,
                new PdfImageStyle {
                    Fit = OfficeImageFit.Contain,
                    SourceCrop = new PdfImageSourceCrop(0.5D, 0D, 0D, 0D)
                }))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);

        Assert.Contains("100 0 0 100 -25 0 cm", raw, StringComparison.Ordinal);
        Assert.Contains("0.5 0 0.5 1 re", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void CanvasItem_OutsidePageBounds_ThrowsClearDiagnostic() {
        var doc = PdfDocument.Create(new PdfOptions {
                PageWidth = 100,
                PageHeight = 100,
                MarginLeft = 10,
                MarginRight = 10,
                MarginTop = 10,
                MarginBottom = 10
            })
            .Canvas(canvas => canvas.Text("Out", 90, 10, 20, 20));

        ArgumentException ex = Assert.Throws<ArgumentException>(() => doc.ToBytes());
        Assert.Contains("Canvas text exceeds the current page bounds.", ex.Message, StringComparison.Ordinal);
    }

    private static double FindWordStartX(UglyToad.PdfPig.Content.Page page, string word) {
        var lines = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1));

        foreach (var line in lines) {
            var ordered = line.OrderBy(letter => letter.StartBaseLine.X).ToList();
            string text = string.Concat(ordered.Select(letter => letter.Value));
            int index = text.IndexOf(word, StringComparison.Ordinal);
            if (index >= 0) {
                return ordered[index].StartBaseLine.X;
            }
        }

        throw new InvalidOperationException("Could not find word '" + word + "' in rendered PDF text.");
    }

    private static double FindWordStartY(UglyToad.PdfPig.Content.Page page, string word) {
        var lines = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1));

        foreach (var line in lines) {
            var ordered = line.OrderBy(letter => letter.StartBaseLine.X).ToList();
            string text = string.Concat(ordered.Select(letter => letter.Value));
            int index = text.IndexOf(word, StringComparison.Ordinal);
            if (index >= 0) {
                return ordered[index].StartBaseLine.Y;
            }
        }

        throw new InvalidOperationException("Could not find word '" + word + "' in rendered PDF text.");
    }

    private static PdfOptions CreateCanvasOptions() =>
        new PdfOptions {
            PageWidth = 260,
            PageHeight = 180,
            CompressContentStreams = false
        };

    private static (double X1, double Y1, double X2, double Y2) RotateRectangle(PdfLinkAnnotation rectangle, double x, double bottomY, double width, double height, double rotationAngle) {
        double angle = rotationAngle * Math.PI / 180D;
        double cos = Math.Cos(angle);
        double sin = Math.Sin(angle);
        double centerX = x + width / 2D;
        double centerY = bottomY + height / 2D;

        RotatePoint(rectangle.X1, rectangle.Y1, centerX, centerY, cos, sin, out double x1, out double y1);
        RotatePoint(rectangle.X1, rectangle.Y2, centerX, centerY, cos, sin, out double x2, out double y2);
        RotatePoint(rectangle.X2, rectangle.Y1, centerX, centerY, cos, sin, out double x3, out double y3);
        RotatePoint(rectangle.X2, rectangle.Y2, centerX, centerY, cos, sin, out double x4, out double y4);

        return (
            Math.Min(Math.Min(x1, x2), Math.Min(x3, x4)),
            Math.Min(Math.Min(y1, y2), Math.Min(y3, y4)),
            Math.Max(Math.Max(x1, x2), Math.Max(x3, x4)),
            Math.Max(Math.Max(y1, y2), Math.Max(y3, y4)));
    }

    private static void RotatePoint(double x, double y, double centerX, double centerY, double cos, double sin, out double rotatedX, out double rotatedY) {
        double dx = x - centerX;
        double dy = y - centerY;
        rotatedX = centerX + cos * dx - sin * dy;
        rotatedY = centerY + sin * dx + cos * dy;
    }

    private static void AssertClose(double expected, double actual) =>
        Assert.InRange(Math.Abs(expected - actual), 0D, 0.01D);

    private static byte[] CreateMinimalRgbPng() {
        return new byte[] {
            137, 80, 78, 71, 13, 10, 26, 10,
            0, 0, 0, 13,
            73, 72, 68, 82,
            0, 0, 0, 1,
            0, 0, 0, 1,
            8, 2, 0, 0, 0,
            0, 0, 0, 0,
            0, 0, 0, 12,
            73, 68, 65, 84,
            0x78, 0x9C, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, 0x03, 0x01, 0x01, 0x00,
            0, 0, 0, 0,
            0, 0, 0, 0,
            73, 69, 78, 68,
            0, 0, 0, 0
        };
    }
}
