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
        using PowerPointDocument presentation = PowerPointDocument.Load(RoadmapPath);
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
        Assert.Equal(smartArt.Left.ToPoints(), group.Left.ToPoints(), 3);
        Assert.Equal(smartArt.Width.ToPoints(), group.Width.ToPoints(), 3);

        PdfDocumentConversionResult result = presentation.ConvertToPdfWithEvidence();
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "unsupported-smartart");
        Assert.Contains("Validate", PdfDocument.Load(new System.IO.MemoryStream(result.ToBytes())).Pages[slideIndex].ExtractText(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CachedSmartArtRejectsExpandedCacheBeforeMaterialization(bool oversizedXml) {
        using PowerPointDocument presentation = PowerPointDocument.Load(RoadmapPath);
        PowerPointSlide slide = presentation.Slides.Single(item => item.Shapes.OfType<PowerPointSmartArt>().Any());
        PowerPointSmartArt smartArt = slide.Shapes.OfType<PowerPointSmartArt>().Single();
        DiagramPersistLayoutPart part = slide.SlidePart.Parts.Select(item => item.OpenXmlPart)
            .OfType<DiagramPersistLayoutPart>().Single();
        using (Stream stream = part.GetStream(FileMode.Create, FileAccess.Write))
        using (var writer = new StreamWriter(stream)) {
            writer.Write("<drawing><spTree>");
            if (oversizedXml) writer.Write(new string('A', checked((int)PowerPointXmlReader.MaximumPackageXmlCharacters + 1)));
            else for (int index = 0; index <= 4096; index++) writer.Write("<sp/>");
            writer.Write("</spTree></drawing>");
        }

        Assert.False(smartArt.TryCreateCachedDrawing(out _, out string? reason));
        Assert.Contains(oversizedXml ? "XML admission limits" : "supported shape count", reason, StringComparison.Ordinal);
    }

    [Fact]
    public void EditedSmartArtRefreshesItsProducerDrawing() {
        using PowerPointDocument presentation = PowerPointDocument.Load(RoadmapPath);
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
        using var original = PowerPointDocument.Load(RoadmapPath);
        PowerPointSlide originalSlide = original.Slides.Single(item => item.Shapes.OfType<PowerPointSmartArt>().Any());
        PowerPointSmartArt smartArt = originalSlide.Shapes.OfType<PowerPointSmartArt>().Single();
        Assert.True(smartArt.TryCreateCachedDrawing(out PowerPointGroupShape originalGroup, out string? current), current);
        var originalLabels = originalSlide.GetGroupChildren(originalGroup).OfType<PowerPointTextBox>()
            .Where(item => !string.IsNullOrWhiteSpace(item.Text)).OrderBy(item => item.Left.ToPoints()).ToArray();
        string[] modelIds = ReadSmartArtModelIds(originalSlide.SlidePart);
        string[] expected = { swap ? "Represent" : "Dis cover", swap ? "Discover" : "Represent", "Round-trip", "Validate" };
        smartArt.SetNodeText(0, expected[0]);
        if (swap) smartArt.SetNodeText(1, expected[1]);
        using var stream = new MemoryStream();
        original.Save(stream);
        stream.Position = 0;
        using var reopened = PowerPointDocument.Load(stream);
        PowerPointSlide slide = reopened.Slides.Single(item => item.Shapes.OfType<PowerPointSmartArt>().Any());
        PowerPointSmartArt edited = slide.Shapes.OfType<PowerPointSmartArt>().Single();
        Assert.Equal(modelIds, ReadSmartArtModelIds(slide.SlidePart));
        Assert.Equal(expected, edited.GetNodeTexts());
        Assert.True(edited.TryCreateCachedDrawing(out PowerPointGroupShape group, out string? reason), reason);
        var labels = slide.GetGroupChildren(group).OfType<PowerPointTextBox>()
            .Where(item => !string.IsNullOrWhiteSpace(item.Text)).OrderBy(item => item.Left.ToPoints()).ToArray();
        Assert.Equal(expected, labels.Select(item => item.Text));
        Assert.Equal((originalGroup.Left.ToPoints(), originalGroup.Top.ToPoints(), originalGroup.Width.ToPoints(), originalGroup.Height.ToPoints()),
            (group.Left.ToPoints(), group.Top.ToPoints(), group.Width.ToPoints(), group.Height.ToPoints()));
        Assert.Equal(originalLabels.Select(item => (item.Left.ToPoints(), item.Top.ToPoints(), item.Width.ToPoints(), item.Height.ToPoints())),
            labels.Select(item => (item.Left.ToPoints(), item.Top.ToPoints(), item.Width.ToPoints(), item.Height.ToPoints())));

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
            double left = group.Left.ToPoints() + label.Left.ToPoints();
            double top = (double)page.Height - group.Top.ToPoints() - label.Top.ToPoints();
            // The cached group has unscaled child coordinates. Read each label's glyphs in
            // content order: font metrics may wrap a word across lines, and page-wide text
            // extraction interleaves adjacent nodes. Preserve actual spaces without normalization.
            var letters = page.Letters.Where(letter => letter.StartBaseLine.X >= left
                && letter.StartBaseLine.X <= left + label.Width.ToPoints()
                && letter.StartBaseLine.Y <= top
                && letter.StartBaseLine.Y >= top - label.Height.ToPoints()).ToArray();
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
        using PowerPointDocument presentation = PowerPointDocument.Load(RoadmapPath);
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
