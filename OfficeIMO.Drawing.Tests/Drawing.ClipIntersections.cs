using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingClipIntersectionTests {
    [Fact]
    public void TextInkInspectionAppliesEveryClipIntersection() {
        // The text sits inside the rectangular primary clip but outside its right-half intersection.
        var child = new OfficeDrawing(100, 100).AddFont("Ink", ManagedTextShapingTestAssets.CreateFont('A'));
        child.AddPositionedTextWithNaturalAdvance("A", 0, 20, 10, 40, new OfficeFontInfo("Ink", 30), OfficeColor.Black, 40);
        var parent = new OfficeDrawing(100, 100);
        parent.AddClippedDrawing(child, 0, 0, RectangleWithRightHalfIntersection());

        var results = new List<(bool HasInk, bool Measured, bool Clipped)>();
        new OfficeRasterCanvas(new OfficeRasterImage(1, 1)).InspectDrawingTextInk(parent, OfficeTransform.Identity,
            Array.Empty<OfficeTextInkClip>(), (ink, reason) => results.Add((ink.HasInk, ink.IsMeasured, ink.IsClipped)));

        var result = Assert.Single(results);
        Assert.True(result.Measured);
        Assert.True(result.Clipped);
        Assert.False(result.HasInk);
    }

    [Fact]
    public void PortableExportReportsARectangleClipWithIntersectionsAsUnsupported() {
        var child = new OfficeDrawing(100, 100);
        child.AddShape(OfficeShape.Rectangle(100, 100), 0, 0);
        var parent = new OfficeDrawing(100, 100);
        parent.AddClippedDrawing(child, 0, 0, RectangleWithRightHalfIntersection());

        Assert.False(OfficeDrawingPortableExporter.TryExportJson(parent, new OfficeDrawingPortableOptions("OfficeIMO Baseline Sans"),
            out string? json, out IReadOnlyList<OfficeDiagnostic> diagnostics));
        Assert.Null(json);
        Assert.Equal("DRAWING_CLIP_UNSUPPORTED", Assert.Single(diagnostics).Code);
    }

    private static OfficeClipPath RectangleWithRightHalfIntersection() =>
        OfficeClipPath.Rectangle(100, 100).WithIntersections(new[] {
            OfficeClipPath.Path(100, 100, new[] {
                OfficePathCommand.MoveTo(new OfficePoint(50, 0)),
                OfficePathCommand.LineTo(new OfficePoint(100, 0)),
                OfficePathCommand.LineTo(new OfficePoint(100, 100)),
                OfficePathCommand.LineTo(new OfficePoint(50, 100)),
                OfficePathCommand.Close()
            }, OfficeFillRule.NonZero)
        });
}
