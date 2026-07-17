using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class DrawingInkTests {
    [Fact]
    public void InkDocumentOwnsFormatNeutralStrokeDataAndClonesIt() {
        var stroke = new OfficeInkStroke {
            Color = OfficeColor.SteelBlue,
            Width = 4,
            Height = 2,
            Opacity = 0.75,
            Bias = OfficeInkBias.Handwriting,
            RecognizedText = "hello"
        };
        stroke.AddPoint(10, 20, 0.25).AddPoint(30, 40, 1);
        stroke.RecognitionAlternatives.Add("hallo");

        var ink = new OfficeInkDocument().Add(stroke);
        OfficeInkDocument clone = ink.Clone();
        stroke.AddPoint(50, 60);

        OfficeInkStroke clonedStroke = Assert.Single(clone.Strokes);
        Assert.Equal(2, clonedStroke.Points.Count);
        Assert.Equal(OfficeColor.SteelBlue, clonedStroke.Color);
        Assert.Equal(OfficeInkBias.Handwriting, clonedStroke.Bias);
        Assert.Equal("hello", clonedStroke.RecognizedText);
        Assert.Equal("hallo", Assert.Single(clonedStroke.RecognitionAlternatives));
    }

    [Fact]
    public void InkBoundsIncludeTipSizeAndTransform() {
        var stroke = new OfficeInkStroke {
            Width = 4,
            Height = 6,
            Transform = OfficeTransform.Scale(2, 3)
        };
        stroke.AddPoint(5, 10).AddPoint(15, 20);

        OfficeInkBounds bounds = stroke.GetBounds();

        Assert.Equal(6, bounds.X, 6);
        Assert.Equal(21, bounds.Y, 6);
        Assert.Equal(28, bounds.Width, 6);
        Assert.Equal(48, bounds.Height, 6);
        Assert.Equal((8D, 18D), stroke.GetTransformedTipDimensions());
    }

    [Fact]
    public void InkRendererAppliesNonUniformTransformToEachTipAxis() {
        var dot = new OfficeInkStroke {
            Width = 4,
            Height = 6,
            Transform = OfficeTransform.Scale(2, 3)
        }.AddPoint(10, 10);
        var segment = new OfficeInkStroke {
            Width = 4,
            Height = 6,
            Transform = OfficeTransform.Scale(2, 3)
        }.AddPoint(10, 10).AddPoint(20, 10);

        OfficeDrawing dotDrawing = OfficeInkRenderer.Render(new OfficeInkDocument().Add(dot), 100, 100);
        OfficeDrawing segmentDrawing = OfficeInkRenderer.Render(new OfficeInkDocument().Add(segment), 100, 100);

        OfficeShape renderedDot = Assert.Single(dotDrawing.Shapes).Shape;
        Assert.Equal(8, renderedDot.Width, 6);
        Assert.Equal(18, renderedDot.Height, 6);
        OfficeShape renderedBody = Assert.Single(segmentDrawing.Shapes, item => item.Shape.Kind == OfficeShapeKind.Polygon).Shape;
        Assert.Equal(18, renderedBody.Height, 6);
    }

    [Fact]
    public void InkTipDimensionsUseCanvasAlignedBoundsForRotationAndShear() {
        var rotated = new OfficeInkStroke {
            Width = 4,
            Height = 6,
            Transform = OfficeTransform.RotateDegrees(90)
        };
        var shearedRectangle = new OfficeInkStroke {
            Width = 4,
            Height = 6,
            TipShape = OfficeInkTipShape.Rectangle,
            Transform = new OfficeTransform(1, 0, 0.5, 1, 0, 0)
        };
        var shearedEllipse = new OfficeInkStroke {
            Width = 4,
            Height = 6,
            Transform = new OfficeTransform(1, 0, 0.5, 1, 0, 0)
        };

        Assert.Equal((6D, 4D), rotated.GetTransformedTipDimensions());
        Assert.Equal((7D, 6D), shearedRectangle.GetTransformedTipDimensions());
        Assert.Equal(Math.Sqrt(25D), shearedEllipse.GetTransformedTipDimensions().Width, 6);
        Assert.Equal(6D, shearedEllipse.GetTransformedTipDimensions().Height, 6);
        Assert.Equal(5D, shearedEllipse.GetTransformedTipExtent(1D, 0D), 6);
        Assert.Equal(6D, shearedEllipse.GetTransformedTipExtent(0D, 1D), 6);
        OfficePoint support = shearedEllipse.GetTransformedTipSupport(1D, 0D);
        Assert.Equal(2.5D, support.X, 6);
        Assert.Equal(1.8D, support.Y, 6);
        Assert.False(shearedEllipse.IsTransformedTipAxisAligned());
        Assert.True(new OfficeInkStroke { Transform = OfficeTransform.Scale(2D, 3D) }.IsTransformedTipAxisAligned());
        Assert.True(new OfficeInkStroke { Width = 4D, Height = 4D, Transform = OfficeTransform.RotateDegrees(45D) }.IsTransformedTipAxisAligned());
        Assert.False(new OfficeInkStroke { TipShape = OfficeInkTipShape.Rectangle, Transform = OfficeTransform.RotateDegrees(45D) }.IsTransformedTipAxisAligned());
        Assert.Throws<ArgumentOutOfRangeException>(() => shearedEllipse.GetTransformedTipExtent(0D, 0D));
    }

    [Fact]
    public void InkRendererUsesTipExtentPerpendicularToEachSegment() {
        var horizontal = new OfficeInkStroke { Width = 4, Height = 6, Transform = OfficeTransform.Scale(2, 3) }
            .AddPoint(10, 10).AddPoint(40, 10);
        var vertical = new OfficeInkStroke { Width = 4, Height = 6, Transform = OfficeTransform.Scale(2, 3) }
            .AddPoint(10, 10).AddPoint(10, 40);

        OfficeShape horizontalBody = Assert.Single(
            OfficeInkRenderer.Render(new OfficeInkDocument().Add(horizontal), 100, 200).Shapes,
            item => item.Shape.Kind == OfficeShapeKind.Polygon).Shape;
        OfficeShape verticalBody = Assert.Single(
            OfficeInkRenderer.Render(new OfficeInkDocument().Add(vertical), 100, 200).Shapes,
            item => item.Shape.Kind == OfficeShapeKind.Polygon).Shape;

        Assert.Equal(18D, horizontalBody.Height, 6);
        Assert.Equal(8D, verticalBody.Width, 6);
    }

    [Fact]
    public void InkRendererRetainsAffineTipGeometryAtEndpointsAndDots() {
        var stroke = new OfficeInkStroke {
            Width = 4,
            Height = 6,
            Transform = new OfficeTransform(1, 0, 0.5, 1, 0, 0)
        }.AddPoint(10, 10).AddPoint(30, 10);
        var dot = new OfficeInkStroke {
            Width = 4,
            Height = 6,
            Transform = new OfficeTransform(1, 0, 0.5, 1, 0, 0)
        }.AddPoint(10, 10);

        OfficeDrawing strokeDrawing = OfficeInkRenderer.Render(new OfficeInkDocument().Add(stroke), 100, 100);
        OfficeDrawing dotDrawing = OfficeInkRenderer.Render(new OfficeInkDocument().Add(dot), 100, 100);

        Assert.Equal(2, strokeDrawing.Shapes.Count(item => item.Shape.Kind == OfficeShapeKind.Ellipse));
        Assert.All(strokeDrawing.Shapes.Where(item => item.Shape.Kind == OfficeShapeKind.Ellipse), item => Assert.True(item.Shape.Transform.HasValue));
        Assert.True(Assert.Single(dotDrawing.Shapes).Shape.Transform.HasValue);
    }

    [Fact]
    public void InkRendererProjectsPressureAndHighlighterOpacityIntoDrawingScene() {
        var stroke = new OfficeInkStroke {
            Color = OfficeColor.Yellow,
            Width = 8,
            Height = 8,
            IsHighlighter = true
        };
        stroke.AddPoint(10, 10, 0).AddPoint(40, 10, 1).AddPoint(70, 40, 0.5);
        var ink = new OfficeInkDocument().Add(stroke);

        OfficeDrawing drawing = OfficeInkRenderer.Render(ink, 100, 60);

        Assert.Equal(2, drawing.Shapes.Count);
        Assert.Equal(5, drawing.Shapes[0].Shape.StrokeWidth, 6);
        Assert.Equal(6.5, drawing.Shapes[1].Shape.StrokeWidth, 6);
        Assert.All(drawing.Shapes, item => Assert.Equal(0.4, item.Shape.StrokeOpacity!.Value, 6));
        Assert.All(drawing.Shapes, item => Assert.Equal(OfficeStrokeLineCap.Round, item.Shape.StrokeLineCap));
    }

    [Fact]
    public void InkRendererClipsSegmentsAtCanvasEdges() {
        var stroke = new OfficeInkStroke().AddPoint(-20, 20).AddPoint(20, 20).AddPoint(120, 20);
        var drawing = OfficeInkRenderer.Render(new OfficeInkDocument().Add(stroke), 100, 40);

        Assert.Equal(2, drawing.Shapes.Count);
        Assert.True(drawing.Shapes[0].X < 0D);
        Assert.Equal(20, drawing.Shapes[0].X + drawing.Shapes[0].Shape.Width, 6);
        Assert.Equal(20, drawing.Shapes[1].X, 6);
        Assert.True(drawing.Shapes[1].X + drawing.Shapes[1].Shape.Width > drawing.Width);
    }

    [Fact]
    public void InkRendererKeepsSegmentsWhosePaintedBoundsOverlapTheCanvas() {
        var stroke = new OfficeInkStroke().AddPoint(-10, -10).AddPoint(0, 0);

        OfficeDrawing drawing = OfficeInkRenderer.Render(new OfficeInkDocument().Add(stroke), 100, 40);

        Assert.Single(drawing.Shapes);
    }

    [Fact]
    public void InkRendererKeepsThickSegmentsAndDotsWithCentersOutsideTheCanvas() {
        var segment = new OfficeInkStroke { Width = 8, Height = 8 }.AddPoint(10, -1).AddPoint(90, -1);
        var dot = new OfficeInkStroke { Width = 8, Height = 8 }.AddPoint(-1, 20);

        OfficeDrawing segmentDrawing = OfficeInkRenderer.Render(new OfficeInkDocument().Add(segment), 100, 40);
        OfficeDrawing dotDrawing = OfficeInkRenderer.Render(new OfficeInkDocument().Add(dot), 100, 40);

        Assert.Equal(-1, Assert.Single(segmentDrawing.Shapes).Y, 6);
        Assert.Equal(-5, Assert.Single(dotDrawing.Shapes).X, 6);
    }

    [Fact]
    public void InkPointsValidatePressureAndCoordinates() {
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeInkPoint(double.NaN, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeInkPoint(0, 0, 1.01));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeInkPoint(0, 0, timestamp: TimeSpan.FromTicks(-1)));
    }
}
