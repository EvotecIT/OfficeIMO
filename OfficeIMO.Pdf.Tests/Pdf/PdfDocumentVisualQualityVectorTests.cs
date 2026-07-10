using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Fact]
    public void VectorRoundedRectangle_RendersBezierCornersFromSharedShapeDescriptor() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20
            })
            .RoundedRectangle(
                width: 100,
                height: 36,
                cornerRadius: 8,
                strokeColor: PdfColor.FromRgb(26, 51, 77),
                strokeWidth: 2,
                fillColor: PdfColor.FromRgb(204, 179, 153),
                align: PdfAlign.Center,
                spacingBefore: 4,
                spacingAfter: 6)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("0.8 0.702 0.6 rg", content);
        Assert.Contains("0.102 0.2 0.302 RG", content);
        Assert.Contains("2 w", content);
        Assert.Contains("78 124 m", content);
        Assert.Contains("162 124 l", content);
        Assert.Contains("166.418 124 170 127.582 170 132 c", content);
        Assert.Contains("70 127.582 73.582 124 78 124 c h B", content);
    }

    [Fact]
    public void VectorLine_RendersStrokeOperatorFromSharedShapeDescriptor() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20
            })
            .Line(
                x1: 0,
                y1: 0,
                x2: 100,
                y2: 40,
                strokeColor: PdfColor.FromRgb(51, 102, 153),
                strokeWidth: 2,
                align: PdfAlign.Center,
                spacingBefore: 4,
                spacingAfter: 6,
                strokeDashStyle: OfficeStrokeDashStyle.Dash)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("0.2 0.4 0.6 RG", content);
        Assert.Contains("2 w", content);
        Assert.Contains("[6 3] 0 d", content);
        Assert.Contains("70 160 m 170 120 l S", content);
    }

    [Fact]
    public void VectorLine_RendersConfiguredStrokeCapAndJoin() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20
            })
            .Line(
                x1: 0,
                y1: 0,
                x2: 100,
                y2: 0,
                strokeColor: PdfColor.FromRgb(51, 102, 153),
                strokeWidth: 3,
                align: PdfAlign.Center,
                spacingBefore: 4,
                spacingAfter: 6,
                strokeLineCap: OfficeStrokeLineCap.Square,
                strokeLineJoin: OfficeStrokeLineJoin.Bevel)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("0.2 0.4 0.6 RG", content);
        Assert.Contains("3 w", content);
        Assert.Contains("2 J", content);
        Assert.Contains("2 j", content);
        Assert.Contains("70 160 m 170 160 l S", content);
    }

    [Fact]
    public void VectorShape_UsesSharedOfficeDrawingShapeDescriptor() {
        var shape = OfficeShape.Rectangle(90, 24);
        shape.FillColor = OfficeColor.WhiteSmoke;
        shape.StrokeColor = OfficeColor.SteelBlue;
        shape.StrokeWidth = 1.5;

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Shape(shape, align: PdfAlign.Right)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("0.961 0.961 0.961 rg", content);
        Assert.Contains("0.275 0.51 0.706 RG", content);
        Assert.Contains("1.5 w", content);
        Assert.Contains("100 106 90 24 re B", content);
    }

    [Fact]
    public void VectorShape_RendersSharedTransformAsGraphicsStateMatrix() {
        var shape = OfficeShape.Rectangle(40, 20);
        shape.FillColor = OfficeColor.WhiteSmoke;
        shape.StrokeColor = OfficeColor.SteelBlue;
        shape.StrokeWidth = 1.5;
        shape.Transform = OfficeTransform.Translate(10, 5);

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Shape(shape)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("q\n1 0 0 -1 40 125 cm", content);
        Assert.Contains("0.961 0.961 0.961 rg", content);
        Assert.Contains("0.275 0.51 0.706 RG", content);
        Assert.Contains("1.5 w", content);
        Assert.Contains("0 0 40 20 re B", content);
    }

    [Fact]
    public void VectorShape_RendersSharedOpacityAsExtGStateResource() {
        var shape = OfficeShape.Rectangle(90, 24);
        shape.FillColor = OfficeColor.WhiteSmoke;
        shape.StrokeColor = OfficeColor.SteelBlue;
        shape.StrokeWidth = 1.5;
        shape.FillOpacity = 0.35;
        shape.StrokeOpacity = 0.75;

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Shape(shape)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("<< /Type /ExtGState /ca 0.35 /CA 0.75 >>", content);
        Assert.Contains("/ExtGState << /GS1 ", content);
        Assert.Contains("q\n/GS1 gs\nq\n0.961 0.961 0.961 rg", content);
        Assert.Contains("30 106 90 24 re B", content);
    }

    [Fact]
    public void VectorShape_RendersSharedClipPathBeforePainting() {
        var shape = OfficeShape.Rectangle(90, 40);
        shape.FillColor = OfficeColor.WhiteSmoke;
        shape.StrokeColor = OfficeColor.SteelBlue;
        shape.StrokeWidth = 1.5;
        shape.ClipPath = OfficeClipPath.Rectangle(45, 20);

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Shape(shape)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("q\n30 110 45 20 re W n\nq\n0.961 0.961 0.961 rg", content);
        Assert.Contains("30 90 90 40 re B", content);
    }

    [Fact]
    public void VectorShape_RendersSharedClipPathInsideTransformGraphicsState() {
        var shape = OfficeShape.Rectangle(80, 40);
        shape.FillColor = OfficeColor.WhiteSmoke;
        shape.ClipPath = OfficeClipPath.RoundedRectangle(40, 20, 6);
        shape.Transform = OfficeTransform.Translate(10, 5);

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Shape(shape)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("q\n1 0 0 -1 40 125 cm", content);
        Assert.Contains("6 0 m 34 0 l", content);
        Assert.Contains("W n\n0.961 0.961 0.961 rg", content);
        Assert.Contains("0 0 80 40 re f", content);
    }

    [Fact]
    public void VectorShape_RendersSharedLinearGradientAsAxialShadingResource() {
        var shape = OfficeShape.Rectangle(90, 24);
        shape.FillColor = OfficeColor.Red;
        shape.FillGradient = OfficeLinearGradient.Horizontal(OfficeColor.SteelBlue, OfficeColor.WhiteSmoke);
        shape.StrokeColor = OfficeColor.Black;
        shape.StrokeWidth = 1.25;

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Shape(shape)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("/Shading << /SH1 ", content);
        Assert.Contains("/ShadingType 2 /ColorSpace /DeviceRGB /Coords [30 118 120 118]", content);
        Assert.Contains("/C0 [0.275 0.51 0.706] /C1 [0.961 0.961 0.961]", content);
        Assert.Contains("q\n30 106 90 24 re W n\n/SH1 sh\nQ", content);
        Assert.Contains("1.25 w", content);
        Assert.Contains("30 106 90 24 re S", content);
        Assert.DoesNotContain("1 0 0 rg", content, StringComparison.Ordinal);
    }

    [Fact]
    public void VectorShape_RendersEveryMultiStopGradientSegmentAsAStitchingFunction() {
        var shape = OfficeShape.Rectangle(90, 24);
        shape.FillGradient = new OfficeLinearGradient(
            0,
            0.5,
            1,
            0.5,
            new[] {
                new OfficeGradientStop(0, OfficeColor.Red),
                new OfficeGradientStop(0.5, OfficeColor.Lime),
                new OfficeGradientStop(1, OfficeColor.Blue)
            });

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Shape(shape)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("/FunctionType 3 /Domain [0 1]", content);
        Assert.Contains("/C0 [1 0 0] /C1 [0 1 0]", content);
        Assert.Contains("/C0 [0 1 0] /C1 [0 0 1]", content);
        Assert.Contains("/Bounds [0.5] /Encode [0 1 0 1]", content);
    }

    [Fact]
    public void VectorShape_RendersMultiStopRadialGradientAsScaledNativeShading() {
        var shape = OfficeShape.Rectangle(90, 40);
        shape.FillRadialGradient = new OfficeRadialGradient(
            0.5D,
            0.5D,
            0D,
            0.5D,
            0.5D,
            0.5D,
            new[] {
                new OfficeGradientStop(0D, OfficeColor.Red),
                new OfficeGradientStop(0.5D, OfficeColor.Lime),
                new OfficeGradientStop(1D, OfficeColor.Blue)
            });

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Shape(shape)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("/ShadingType 3 /ColorSpace /DeviceRGB /Coords [0.5 0.5 0 0.5 0.5 0.5]", content);
        Assert.Contains("/FunctionType 3 /Domain [0 1]", content);
        Assert.Contains("/Bounds [0.5] /Encode [0 1 0 1]", content);
        Assert.Contains("q\n30 90 90 40 re W n\n90 0 0 40 30 90 cm\n/SH1 sh\nQ", content);
    }

    [Fact]
    public void VectorShape_RendersSharedLinearGradientInsideTransformGraphicsState() {
        var shape = OfficeShape.Rectangle(40, 20);
        shape.FillGradient = OfficeLinearGradient.Vertical(OfficeColor.SteelBlue, OfficeColor.WhiteSmoke);
        shape.Transform = OfficeTransform.Translate(10, 5);

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Shape(shape)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("/ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 20 20 0]", content);
        Assert.Contains("q\n1 0 0 -1 40 125 cm", content);
        Assert.Contains("q\n0 0 40 20 re W n\n/SH1 sh\nQ", content);
    }

    [Fact]
    public void VectorShape_RendersRadialGradientInsideTransformGraphicsState() {
        var shape = OfficeShape.Rectangle(40, 20);
        shape.FillRadialGradient = OfficeRadialGradient.Centered(OfficeColor.Red, OfficeColor.Blue);
        shape.Transform = OfficeTransform.Translate(10, 5);

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Shape(shape)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("/ShadingType 3 /ColorSpace /DeviceRGB /Coords [0.5 0.5 0 0.5 0.5 0.5]", content);
        Assert.Contains("q\n1 0 0 -1 40 125 cm", content);
        Assert.Contains("q\n0 0 40 20 re W n\n40 0 0 20 0 0 cm\n/SH1 sh\nQ", content);
    }

    [Fact]
    public void VectorShape_RendersSharedShadowBehindShapeGeometry() {
        var shape = OfficeShape.RoundedRectangle(90, 24, 6);
        shape.FillColor = OfficeColor.WhiteSmoke;
        shape.StrokeColor = OfficeColor.SteelBlue;
        shape.StrokeWidth = 1;
        shape.Shadow = new OfficeShadow(OfficeColor.Black, 0.22, 3, 4);

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Shape(shape)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("/ExtGState << /GS1 ", content);
        Assert.Contains("/Type /ExtGState /ca 0.22 /CA 0.22", content);
        Assert.Contains("q\n/GS1 gs\nq\n0 0 0 rg", content);
        Assert.Contains("39 102", content);
        Assert.Contains("0.961 0.961 0.961 rg", content);
        Assert.True(content.IndexOf("/GS1 gs", StringComparison.Ordinal) < content.IndexOf("0.961 0.961 0.961 rg", StringComparison.Ordinal));
    }

    [Fact]
    public void VectorRectangle_RendersConfiguredDashStyle() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Rectangle(
                width: 90,
                height: 24,
                strokeColor: PdfColor.FromRgb(51, 102, 153),
                strokeWidth: 2,
                align: PdfAlign.Left,
                strokeDashStyle: OfficeStrokeDashStyle.DashDot)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("0.2 0.4 0.6 RG", content);
        Assert.Contains("2 w", content);
        Assert.Contains("1 J", content);
        Assert.Contains("[6 3 2 3] 0 d", content);
        Assert.Contains("30 106 90 24 re S", content);
    }

    [Fact]
    public void VectorEllipse_RendersBezierPathFromSharedShapeDescriptor() {
        var shape = OfficeShape.Ellipse(80, 40);
        shape.FillColor = OfficeColor.WhiteSmoke;
        shape.StrokeColor = OfficeColor.SteelBlue;
        shape.StrokeWidth = 2;
        shape.StrokeDashStyle = OfficeStrokeDashStyle.Dot;

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20
            })
            .Shape(shape, align: PdfAlign.Center, spacingBefore: 4, spacingAfter: 6)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("0.961 0.961 0.961 rg", content);
        Assert.Contains("0.275 0.51 0.706 RG", content);
        Assert.Contains("2 w", content);
        Assert.Contains("1 J", content);
        Assert.Contains("[2 3] 0 d", content);
        Assert.Contains("160 140 m", content);
        Assert.Contains("160 151.046 142.091 160 120 160 c", content);
        Assert.Contains("142.091 120 160 128.954 160 140 c B", content);
    }

    [Fact]
    public void VectorPolygon_RendersClosedPathFromSharedShapeDescriptor() {
        var shape = OfficeShape.Polygon(
            new OfficePoint(0, 40),
            new OfficePoint(40, 0),
            new OfficePoint(80, 40));
        shape.FillColor = OfficeColor.WhiteSmoke;
        shape.StrokeColor = OfficeColor.SteelBlue;
        shape.StrokeWidth = 1.5;

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20
            })
            .Shape(shape, align: PdfAlign.Center, spacingBefore: 4, spacingAfter: 6)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("0.961 0.961 0.961 rg", content);
        Assert.Contains("0.275 0.51 0.706 RG", content);
        Assert.Contains("1.5 w", content);
        Assert.Contains("80 120 m", content);
        Assert.Contains("120 160 l", content);
        Assert.Contains("160 120 l", content);
        Assert.Contains("h B", content);
    }

    [Fact]
    public void VectorPolygon_RendersConfiguredStrokeJoinFromSharedShapeDescriptor() {
        var shape = OfficeShape.Polygon(
            new OfficePoint(0, 40),
            new OfficePoint(40, 0),
            new OfficePoint(80, 40));
        shape.StrokeColor = OfficeColor.SteelBlue;
        shape.StrokeWidth = 2;
        shape.StrokeLineJoin = OfficeStrokeLineJoin.Round;

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20
            })
            .Shape(shape, align: PdfAlign.Center, spacingBefore: 4, spacingAfter: 6)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("0.275 0.51 0.706 RG", content);
        Assert.Contains("2 w", content);
        Assert.Contains("1 j", content);
        Assert.Contains("80 120 m", content);
        Assert.Contains("h S", content);
    }

    [Fact]
    public void VectorPath_RendersMoveCurveAndCloseOperatorsFromSharedShapeDescriptor() {
        var shape = OfficeShape.Path(
            OfficePathCommand.MoveTo(0, 40),
            OfficePathCommand.CubicBezierTo(20, 0, 60, 0, 80, 40),
            OfficePathCommand.Close());
        shape.FillColor = OfficeColor.WhiteSmoke;
        shape.StrokeColor = OfficeColor.SteelBlue;
        shape.StrokeWidth = 1.5;

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20
            })
            .Shape(shape, align: PdfAlign.Center, spacingBefore: 4, spacingAfter: 6)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("0.961 0.961 0.961 rg", content);
        Assert.Contains("0.275 0.51 0.706 RG", content);
        Assert.Contains("1.5 w", content);
        Assert.Contains("80 120 m", content);
        Assert.Contains("100 160 140 160 160 120 c", content);
        Assert.Contains("h", content);
        Assert.Contains("B", content);
    }

    [Fact]
    public void VectorPath_RendersQuadraticBezierCommandsAsPdfCubicCurves() {
        var shape = OfficeShape.Path(
            OfficePathCommand.MoveTo(0, 40),
            OfficePathCommand.QuadraticBezierTo(40, 0, 80, 40),
            OfficePathCommand.Close());
        shape.FillColor = OfficeColor.WhiteSmoke;
        shape.StrokeColor = OfficeColor.SteelBlue;
        shape.StrokeWidth = 1.5;

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20
            })
            .Shape(shape, align: PdfAlign.Center, spacingBefore: 4, spacingAfter: 6)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("80 120 m", content);
        Assert.Contains("106.667 146.667 133.333 146.667 160 120 c", content);
        Assert.Contains("h", content);
        Assert.Contains("B", content);
    }

    [Fact]
    public void VectorDrawing_RendersPositionedShapesFromSharedDrawingScene() {
        var background = OfficeShape.Rectangle(120, 60);
        background.FillColor = OfficeColor.WhiteSmoke;

        var marker = OfficeShape.Polygon(
            new OfficePoint(0, 30),
            new OfficePoint(40, 0),
            new OfficePoint(80, 30));
        marker.FillColor = OfficeColor.SteelBlue;
        marker.StrokeColor = OfficeColor.Black;
        marker.StrokeWidth = 1.25;

        var drawing = new OfficeDrawing(120, 60)
            .AddShape(background, 0, 0)
            .AddShape(marker, 20, 15);

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20
            })
            .Drawing(drawing, align: PdfAlign.Center, spacingBefore: 4, spacingAfter: 6)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("0.961 0.961 0.961 rg", content);
        Assert.Contains("60 100 120 60 re f", content);
        Assert.Contains("0.275 0.51 0.706 rg", content);
        Assert.Contains("0 0 0 RG", content);
        Assert.Contains("1.25 w", content);
        Assert.Contains("80 115 m", content);
        Assert.Contains("120 145 l", content);
        Assert.Contains("160 115 l", content);
        Assert.Contains("h B", content);
    }

    [Fact]
    public void VectorShape_RejectsInvalidGeometry() {
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfDocument.Create()
                .Rectangle(width: -1, height: 24));

        var invalidStroke = OfficeShape.Rectangle(90, 24);
        invalidStroke.StrokeWidth = -0.5;

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfDocument.Create()
                .Shape(invalidStroke));

        var invalidOpacity = OfficeShape.Rectangle(90, 24);
        invalidOpacity.FillOpacity = 1.1;

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfDocument.Create()
                .Shape(invalidOpacity));

        var invalidClipPath = OfficeShape.Rectangle(90, 24);
        invalidClipPath.ClipPath = OfficeClipPath.Rectangle(91, 24);

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfDocument.Create()
                .Shape(invalidClipPath));

        var invalidPolygon = new OfficeShape {
            Kind = OfficeShapeKind.Polygon,
            Width = 20,
            Height = 20
        };

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Shape(invalidPolygon));

        var invalidPath = new OfficeShape {
            Kind = OfficeShapeKind.Path,
            Width = 20,
            Height = 20
        };

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Shape(invalidPath));

        var invalidLine = new OfficeShape {
            Kind = OfficeShapeKind.Line,
            Width = 20,
            Height = 20
        };

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Shape(invalidLine));

        var invalidRoundedRectangle = OfficeShape.RoundedRectangle(40, 20, 4);
        invalidRoundedRectangle.CornerRadius = 11;

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfDocument.Create()
                .Shape(invalidRoundedRectangle));

        var emptyDrawing = new OfficeDrawing(40, 20);

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Drawing(emptyDrawing));

        var shapeSpacingBeforeException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Rectangle(width: 40, height: 20, spacingBefore: -1));

        Assert.Contains("Shape spacing before must be a non-negative finite value.", shapeSpacingBeforeException.Message, StringComparison.Ordinal);

        var shapeSpacingAfterException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Rectangle(width: 40, height: 20, spacingAfter: double.PositiveInfinity));

        Assert.Contains("Shape spacing after must be a non-negative finite value.", shapeSpacingAfterException.Message, StringComparison.Ordinal);

        var drawing = new OfficeDrawing(40, 20)
            .AddShape(OfficeShape.Rectangle(40, 20), 0, 0);

        var drawingSpacingBeforeException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Drawing(drawing, spacingBefore: double.NaN));

        Assert.Contains("Drawing spacing before must be a non-negative finite value.", drawingSpacingBeforeException.Message, StringComparison.Ordinal);

        var drawingSpacingAfterException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(100, column =>
                                    column.Drawing(drawing, spacingAfter: -1)))))));

        Assert.Contains("Drawing spacing after must be a non-negative finite value.", drawingSpacingAfterException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void VectorShapeAndDrawing_RejectFlowBlocksTallerThanContentArea() {
        var options = new PdfOptions {
            PageWidth = 220,
            PageHeight = 140,
            MarginLeft = 20,
            MarginRight = 20,
            MarginTop = 20,
            MarginBottom = 20
        };

        var shapeException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(options)
                .Rectangle(width: 80, height: 130)
                .ToBytes());
        Assert.Contains("Shape height exceeds the available page content height.", shapeException.Message, StringComparison.Ordinal);

        var drawing = new OfficeDrawing(80, 130)
            .AddShape(OfficeShape.Rectangle(80, 130), 0, 0);

        var drawingException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(options)
                .Drawing(drawing)
                .ToBytes());
        Assert.Contains("Drawing height exceeds the available page content height.", drawingException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void VectorShapeAndDrawing_RejectFlowBlocksWiderThanContentArea() {
        var options = new PdfOptions {
            PageWidth = 220,
            PageHeight = 180,
            MarginLeft = 20,
            MarginRight = 20,
            MarginTop = 20,
            MarginBottom = 20
        };

        var shapeException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(options)
                .Rectangle(width: 190, height: 40)
                .ToBytes());
        Assert.Contains("Shape width exceeds the available page content width.", shapeException.Message, StringComparison.Ordinal);

        var drawing = new OfficeDrawing(190, 40)
            .AddShape(OfficeShape.Rectangle(190, 40), 0, 0);

        var drawingException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(options)
                .Drawing(drawing)
                .ToBytes());
        Assert.Contains("Drawing width exceeds the available page content width.", drawingException.Message, StringComparison.Ordinal);
    }
}
