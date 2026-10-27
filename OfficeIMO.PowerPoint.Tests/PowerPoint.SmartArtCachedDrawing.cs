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

        PdfDocumentConversionResult result = presentation.ConvertToPdfWithEvidence();
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "unsupported-smartart");
        Assert.Contains("Validate", PdfDocument.Load(new System.IO.MemoryStream(result.ToBytes())).Pages[slideIndex].ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void EditedSmartArtRefreshesItsProducerDrawing() {
        using PowerPointPresentation presentation = PowerPointPresentation.Load(RoadmapPath);
        PowerPointSmartArt smartArt = presentation.Slides.SelectMany(slide => slide.Shapes).OfType<PowerPointSmartArt>().Single();
        Assert.True(smartArt.TryCreateCachedDrawing(out _, out string? current), current);

        smartArt.SetNodeText(0, "Explore");
        Assert.True(smartArt.TryCreateCachedDrawing(out _, out string? reason), reason);
        PdfDocumentConversionResult result = presentation.ConvertToPdfWithEvidence();
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "unsupported-smartart");
        Assert.Contains("Explore", PdfDocument.Load(new System.IO.MemoryStream(result.ToBytes())).ExtractText(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void EditedSmartArtKeepsWhitespaceAndNodePositionsAfterReopen(bool swap, bool embeddedFont) {
        using var original = PowerPointPresentation.Load(RoadmapPath);
        PowerPointSlide originalSlide = original.Slides.Single(item => item.Shapes.OfType<PowerPointSmartArt>().Any());
        PowerPointSmartArt smartArt = originalSlide.Shapes.OfType<PowerPointSmartArt>().Single();
        Assert.True(smartArt.TryCreateCachedDrawing(out PowerPointGroupShape originalGroup, out string? current), current);
        var originalLabels = originalSlide.GetGroupChildren(originalGroup).OfType<PowerPointTextBox>()
            .Where(item => !string.IsNullOrWhiteSpace(item.Text)).OrderBy(item => item.LeftPoints).ToArray();
        string[] modelIds = ReadSmartArtModelIds(originalSlide.SlidePart);
        string[] expected = { swap ? "Represent" : "Dis cover", swap ? "Discover" : "Represent", "Round-trip", "Validate" };
        smartArt.SetNodeText(0, expected[0]);
        if (swap) smartArt.SetNodeText(1, expected[1]);
        using var stream = new MemoryStream();
        original.Save(stream);
        stream.Position = 0;
        using var reopened = PowerPointPresentation.Load(stream);
        PowerPointSlide slide = reopened.Slides.Single(item => item.Shapes.OfType<PowerPointSmartArt>().Any());
        PowerPointSmartArt edited = slide.Shapes.OfType<PowerPointSmartArt>().Single();
        Assert.Equal(modelIds, ReadSmartArtModelIds(slide.SlidePart));
        Assert.Equal(expected, edited.GetNodeTexts());
        Assert.True(edited.TryCreateCachedDrawing(out PowerPointGroupShape group, out string? reason), reason);
        var labels = slide.GetGroupChildren(group).OfType<PowerPointTextBox>()
            .Where(item => !string.IsNullOrWhiteSpace(item.Text)).OrderBy(item => item.LeftPoints).ToArray();
        Assert.Equal(expected, labels.Select(item => item.Text));
        Assert.Equal((originalGroup.LeftPoints, originalGroup.TopPoints, originalGroup.WidthPoints, originalGroup.HeightPoints),
            (group.LeftPoints, group.TopPoints, group.WidthPoints, group.HeightPoints));
        Assert.Equal(originalLabels.Select(item => (item.LeftPoints, item.TopPoints, item.WidthPoints, item.HeightPoints)),
            labels.Select(item => (item.LeftPoints, item.TopPoints, item.WidthPoints, item.HeightPoints)));

        // Resolve fonts explicitly: neither case depends on fonts installed on the test host.
        var pdfOptions = new PdfCreateOptions { DefaultFont = PdfStandardFont.Helvetica };
        if (embeddedFont) pdfOptions.UseFontFamily(new PdfEmbeddedFontFamily("SmartArtFixture",
            OfficeIMO.TestAssets.PdfTestFontAssets.LoadBundledOpenTypeCffFont()));
        PdfDocumentConversionResult result = reopened.ConvertToPdfWithEvidence(new PowerPointPdfExportOptions {
            PdfOptions = pdfOptions,
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        });
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "unsupported-smartart");
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(result.ToBytes());
        var page = pdf.GetPage(reopened.Slides.ToList().IndexOf(slide) + 1);
        for (int index = 0; index < labels.Length; index++) {
            PowerPointTextBox label = labels[index];
            double left = group.LeftPoints + label.LeftPoints;
            double top = (double)page.Height - group.TopPoints - label.TopPoints;
            // The cached group has unscaled child coordinates. Read each label's glyphs in
            // content order: font metrics may wrap a word across lines, and page-wide text
            // extraction interleaves adjacent nodes. Preserve actual spaces without normalization.
            var letters = page.Letters.Where(letter => letter.StartBaseLine.X >= left
                && letter.StartBaseLine.X <= left + label.WidthPoints
                && letter.StartBaseLine.Y <= top
                && letter.StartBaseLine.Y >= top - label.HeightPoints).ToArray();
            Assert.Equal(expected[index], string.Concat(letters.Select(letter => letter.Value)));
            Assert.All(letters, letter => Assert.Contains(embeddedFont ? "SmartArtFixture" : "Helvetica", letter.FontName,
                StringComparison.Ordinal));
        }
    }

    private static string[] ReadSmartArtModelIds(SlidePart slide) => slide.DiagramDataParts.Cast<OpenXmlPart>()
        .Concat(slide.Parts.Select(item => item.OpenXmlPart).OfType<DiagramPersistLayoutPart>())
        .SelectMany(part => {
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            return XDocument.Load(stream).Descendants().Attributes("modelId").Select(item => item.Value).ToArray();
        }).ToArray();

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
        PdfDocumentConversionResult result = presentation.ConvertToPdfWithEvidence();
        Assert.Contains(result.Warnings, warning => warning.Code == "unsupported-smartart");
    }
}
