using System;
using System.IO;
using System.Linq;
using ChartForgeX.Core;
using ChartForgeX.Primitives;
using ChartForgeX.Rendering;
using ChartForgeX.Themes;
using ChartForgeX.VisualArtifacts;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.ChartForgeX.Tests;

public sealed partial class OfficeVisualIntegrationTests {
    [Fact]
    public void RepeatedArtifactIdentifierRetainsEachThemePayloadAndPlacementThroughSavedDocuments() {
        OfficeVisualConversionResult Create(VisualThemeMode mode) {
            var chart = Chart.Create().AddLine("Requests", new[] { new ChartPoint(1, 10), new ChartPoint(2, 20) });
            var artifact = chart.Prepare(new VisualRenderContext(
                layout: new VisualLayoutOptions(new VisualSize(400, 240)), themeMode: mode,
                frame: new VisualFrame(title: "Requests", showLegend: false, showSurface: true)))
                .ToArtifact("requests", VisualArtifactKind.Chart);
            artifact.Accessibility.WithTextAlternative("Requests", "Two request observations in " + mode + " mode.", "en");
            return artifact.ToOfficeVisual(new OfficeVisualConversionOptions { WidthPoints = 300D });
        }

        var light = Create(VisualThemeMode.Light);
        var dark = Create(VisualThemeMode.Dark);
        var expected = new[] { light, dark, light };
        Assert.False(light.GetPlacementBytes().SequenceEqual(dark.GetPlacementBytes()));
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-ChartForgeX-reuse-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try {
            string wordPath = Path.Combine(folder, "repeated.docx");
            using (var document = WordDocument.Create(wordPath)) {
                foreach (var visual in expected) document.AddParagraph().AddVisualArtifact(visual);
                document.Save();
            }
            using (var document = WordDocument.Load(wordPath)) {
                Assert.Equal(expected.Length, document.Images.Count);
                for (int index = 0; index < expected.Length; index++) {
                    var image = document.Images[index];
                    Assert.Equal(expected[index].GetPlacementBytes(), image.ToBytes());
                    Assert.Equal(expected[index].AlternativeText, image.Description);
                    Assert.Equal(400D, image.Width!.Value, 6);
                    Assert.Equal(240D, image.Height!.Value, 6);
                }
            }

            string excelPath = Path.Combine(folder, "repeated.xlsx");
            using (var document = ExcelDocument.Create(excelPath)) {
                var sheet = document.AddWorksheet("Requests");
                for (int index = 0; index < expected.Length; index++) sheet.AddVisualArtifact(2 + index * 14, 2, expected[index]);
                document.Save();
            }
            using (var document = ExcelDocument.Load(excelPath)) {
                var images = document.Sheets[0].Images.OrderBy(image => image.RowIndex).ToArray();
                Assert.Equal(expected.Length, images.Length);
                for (int index = 0; index < expected.Length; index++) {
                    Assert.Equal(expected[index].GetPlacementBytes(), images[index].ToBytes());
                    Assert.Equal(expected[index].AlternativeText, images[index].Description);
                    Assert.Equal(400, images[index].WidthPixels);
                    Assert.Equal(240, images[index].HeightPixels);
                }
            }

            string presentationPath = Path.Combine(folder, "repeated.pptx");
            using (var presentation = PowerPointPresentation.Create(presentationPath)) {
                foreach (var visual in expected) presentation.AddSlide().AddVisualArtifact(visual, 36D, 54D);
                presentation.Save();
            }
            using (var presentation = PowerPointPresentation.Load(presentationPath)) {
                Assert.Equal(expected.Length, presentation.Slides.Count);
                for (int index = 0; index < expected.Length; index++) {
                    var image = Assert.Single(presentation.Slides[index].Pictures);
                    Assert.Equal(expected[index].GetPlacementBytes(), image.GetImageBytes());
                    Assert.Equal(expected[index].AlternativeText, image.Description);
                    Assert.Equal(300D, image.WidthPoints, 6);
                    Assert.Equal(180D, image.HeightPoints, 6);
                }
            }
        } finally {
            if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true);
        }
    }
}
