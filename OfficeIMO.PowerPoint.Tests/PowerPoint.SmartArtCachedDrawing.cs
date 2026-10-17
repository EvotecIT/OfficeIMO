using System;
using System.Linq;
using System.IO;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
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
    public void EditedSmartArtRefreshesItsProducerDrawing() {
        using PowerPointPresentation presentation = PowerPointPresentation.Load(RoadmapPath);
        PowerPointSmartArt smartArt = presentation.Slides.SelectMany(slide => slide.Shapes).OfType<PowerPointSmartArt>().Single();
        Assert.True(smartArt.TryCreateCachedDrawing(out _, out string? current), current);

        smartArt.SetNodeText(0, "Explore");
        Assert.True(smartArt.TryCreateCachedDrawing(out _, out string? reason), reason);
        PdfDocumentConversionResult result = presentation.ToPdfDocumentResult();
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "unsupported-smartart");
        Assert.Contains("Explore", PdfReadDocument.Open(result.ToBytes()).ExtractText(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EditedSmartArtKeepsWhitespaceAndNodePositionsAfterReopen(bool swap) {
        using var original = PowerPointPresentation.Load(RoadmapPath);
        PowerPointSmartArt smartArt = original.Slides.SelectMany(slide => slide.Shapes).OfType<PowerPointSmartArt>().Single();
        smartArt.SetNodeText(0, swap ? "Represent" : "Dis cover");
        if (swap) smartArt.SetNodeText(1, "Discover");
        using var stream = new MemoryStream();
        original.Save(stream);
        stream.Position = 0;
        using var reopened = PowerPointPresentation.Load(stream);
        PowerPointSlide slide = reopened.Slides.Single(item => item.Shapes.OfType<PowerPointSmartArt>().Any());
        PowerPointSmartArt edited = slide.Shapes.OfType<PowerPointSmartArt>().Single();
        Assert.True(edited.TryCreateCachedDrawing(out PowerPointGroupShape group, out string? reason), reason);
        var labels = slide.GetGroupChildren(group).OfType<PowerPointTextBox>()
            .Where(item => !string.IsNullOrWhiteSpace(item.Text)).OrderBy(item => item.LeftPoints).ToArray();
        Assert.Equal(swap ? "Represent" : "Dis cover", labels[0].Text.Trim());
        Assert.Equal(swap ? "Discover" : "Represent", labels[1].Text.Trim());
        PdfDocumentConversionResult result = reopened.ToPdfDocumentResult();
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "unsupported-smartart");
        Assert.Contains(swap ? "Represent" : "Dis cover", PdfReadDocument.Open(result.ToBytes()).ExtractText(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CachedSmartArtRejectsWhitespaceEditsAndSwappedNodeLabels(bool swap) {
        using PowerPointPresentation presentation = PowerPointPresentation.Load(RoadmapPath);
        PowerPointSmartArt smartArt = presentation.Slides.SelectMany(slide => slide.Shapes).OfType<PowerPointSmartArt>().Single();
        Assert.True(smartArt.TryCreateCachedDrawing(out _, out string? current), current);

        // Simulate an externally edited package whose producer cache was not refreshed.
        DiagramDataPart part = presentation.Slides.SelectMany(slide => slide.SlidePart.DiagramDataParts).Single();
        XDocument data;
        using (Stream stream = part.GetStream(FileMode.Open, FileAccess.Read)) data = XDocument.Load(stream);
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        XElement first = data.Descendants(a + "t").Single(item => item.Value == "Discover");
        XElement second = data.Descendants(a + "t").Single(item => item.Value == "Represent");
        first.Value = swap ? "Represent" : "Dis cover";
        if (swap) second.Value = "Discover";
        using (Stream stream = part.GetStream(FileMode.Create, FileAccess.Write)) data.Save(stream);

        Assert.False(smartArt.TryCreateCachedDrawing(out _, out _));
        PdfDocumentConversionResult result = presentation.ToPdfDocumentResult();
        Assert.Contains(result.Warnings, warning => warning.Code == "unsupported-smartart");
    }
}
