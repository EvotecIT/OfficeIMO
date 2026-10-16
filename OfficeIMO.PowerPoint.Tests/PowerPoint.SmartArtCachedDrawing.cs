using System;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Pins SmartArt rendering from the drawing PowerPoint caches with each diagram.</summary>
public sealed class PowerPointSmartArtCachedDrawingTests {
    private static string RoadmapPath => System.IO.Path.Combine(RepositoryTestPaths.Find(), "Assets", "PowerPointTemplates",
        "PowerPointAdvancedRoadmap.pptx");

    [Fact]
    public void SmartArtDrawsTheShapesOfficeCached() {
        using PowerPointPresentation presentation = PowerPointPresentation.Load(RoadmapPath);
        int slideIndex = presentation.Slides.ToList().FindIndex(item => item.Shapes.OfType<PowerPointSmartArt>().Any());
        PowerPointSlide slide = presentation.Slides[slideIndex];
        PowerPointSmartArt smartArt = slide.Shapes.OfType<PowerPointSmartArt>().Single();

        Assert.True(smartArt.TryCreateCachedDrawing(out PowerPointGroupShape group, out string? reason), reason);
        PowerPointShape[] shapes = slide.GetGroupChildren(group).ToArray();
        string[] fills = shapes.Select(shape => shape.FillColor).Where(color => color != null).Select(color => color!).ToArray();
        Assert.Equal(4, fills.Count(color => color == "156082"));
        Assert.Equal(3, shapes.Count(shape => shape is PowerPointAutoShape { ShapeType: OfficePresetShapeType.RightArrow }));
        Assert.Equal(new[] { "Discover", "Represent", "Round-trip", "Validate" },
            shapes.OfType<PowerPointTextBox>().Select(text => text.Text.Trim()).Where(text => text.Length > 0).ToArray());
        Assert.Equal(smartArt.LeftPoints, group.LeftPoints, 3);
        Assert.Equal(smartArt.WidthPoints, group.WidthPoints, 3);

        PdfDocumentConversionResult result = presentation.ToPdfDocumentResult();
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "unsupported-smartart");
        Assert.Contains("Validate", PdfReadDocument.Open(result.ToBytes()).Pages[slideIndex].ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void EditedSmartArtFallsBackFromItsStaleCachedDrawing() {
        using PowerPointPresentation presentation = PowerPointPresentation.Load(RoadmapPath);
        PowerPointSmartArt smartArt = presentation.Slides.SelectMany(slide => slide.Shapes).OfType<PowerPointSmartArt>().Single();
        Assert.True(smartArt.TryCreateCachedDrawing(out _, out string? current), current);

        // Editing a node leaves the producer's cached drawing showing the old text.
        smartArt.SetNodeText(0, "Explore");
        Assert.False(smartArt.TryCreateCachedDrawing(out _, out string? reason));
        Assert.Equal("The SmartArt cached drawing no longer matches the diagram text.", reason);
    }
}
