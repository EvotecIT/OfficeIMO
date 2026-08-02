using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint.Pdf;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.PowerPoint.Benchmarks;

internal static class PowerPointRoadmapVisualCorpus {
    private static readonly (PowerPointSmartArtType Type, string Title)[] SmartArtScenarios = {
        (PowerPointSmartArtType.BasicProcess, "Process"),
        (PowerPointSmartArtType.BasicHierarchy, "Hierarchy"),
        (PowerPointSmartArtType.BasicCycle, "Cycle"),
        (PowerPointSmartArtType.BasicList, "List"),
        (PowerPointSmartArtType.BasicMatrix, "Matrix"),
        (PowerPointSmartArtType.BasicPyramid, "Pyramid"),
        (PowerPointSmartArtType.BasicRelationship, "Relationship")
    };

    internal static int Create(string outputDirectory) {
        if (string.IsNullOrWhiteSpace(outputDirectory)) {
            Console.Error.WriteLine(
                "Usage: --visual-corpus <output-directory>");
            return 2;
        }

        string root = Path.GetFullPath(outputDirectory);
        Directory.CreateDirectory(root);
        string deckPath = Path.Combine(root, "powerpoint-roadmap-visual-corpus.pptx");
        string imagesDirectory = Path.Combine(root, "images");
        Directory.CreateDirectory(imagesDirectory);
        string vectorsDirectory = Path.Combine(root, "vectors");
        Directory.CreateDirectory(vectorsDirectory);

        using (PowerPointPresentation presentation =
               PowerPointPresentation.Create(deckPath)) {
            presentation.SlideSize.SetSizePoints(960, 540);
            var visualSlides = new List<PowerPointSlide>();
            foreach ((PowerPointSmartArtType type, string title) in SmartArtScenarios) {
                visualSlides.Add(AddSmartArtSlide(presentation, type, title));
            }
            visualSlides.Add(AddCustomGeometrySlide(presentation));
            visualSlides.Add(AddChartAndTableSlide(presentation));

            var reviewer = new PowerPointCommentAuthor(
                "OfficeIMO visual review", "OVR", "officeimo-reviewer");
            presentation.AddClassicComment(visualSlides[0], reviewer,
                "Classic comment mutation and preservation proof.",
                PowerPointUnits.FromPoints(40), PowerPointUnits.FromPoints(40));
            PowerPointModernComment modern = presentation.AddModernComment(
                visualSlides[0], reviewer,
                "Modern threaded comment mutation and preservation proof.",
                x: PowerPointUnits.FromPoints(80),
                y: PowerPointUnits.FromPoints(80));
            modern.AddReply(reviewer, "Reply mutation proof.");
            presentation.AddCustomShow("Roadmap visual review", visualSlides);
            presentation.Save();
        }

        using PowerPointPresentation reopened = PowerPointPresentation.Load(deckPath);
        IReadOnlyList<string> validation = reopened.ValidateDocument()
            .Select(error => error.Description ?? error.ToString() ?? string.Empty)
            .ToArray();
        if (validation.Count > 0) {
            throw new InvalidOperationException(
                "Visual corpus Open XML validation failed: " +
                string.Join(" | ", validation.Take(10)));
        }
        if (reopened.CustomShows.Count != 1
            || reopened.GetClassicComments(reopened.Slides[0]).Count != 1
            || reopened.GetModernComments(reopened.Slides[0]).Single().Replies.Count != 1) {
            throw new InvalidOperationException(
                "Visual corpus review metadata or custom-show round trip failed.");
        }

        IReadOnlyList<OfficeImageExportResult> images = reopened.ExportImages(
            OfficeImageExportFormat.Png);
        IReadOnlyList<OfficeImageExportResult> vectors = reopened.ExportImages(
            OfficeImageExportFormat.Svg);
        if (vectors.Count != images.Count) {
            throw new InvalidOperationException(
                "Visual corpus PNG and SVG export counts differ.");
        }
        for (int index = 0; index < images.Count; index++) {
            OfficeImageExportResult image = images[index];
            OfficeImageExportDiagnostic[] failures = image.Diagnostics
                .Concat(vectors[index].Diagnostics)
                .Where(diagnostic =>
                    diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error
                    || (diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Warning
                        && diagnostic.Code != OfficeImageExportDiagnosticCodes.FontSubstituted))
                .ToArray();
            if (failures.Length > 0) {
                int withoutCharts = reopened.Slides[index].ExportImage(
                    OfficeImageExportFormat.Png,
                    new PowerPointImageExportOptions { IncludeCharts = false })
                    .Diagnostics.Count(diagnostic =>
                        diagnostic.Severity != OfficeImageExportDiagnosticSeverity.Info);
                int withoutTables = reopened.Slides[index].ExportImage(
                    OfficeImageExportFormat.Png,
                    new PowerPointImageExportOptions { IncludeTables = false })
                    .Diagnostics.Count(diagnostic =>
                        diagnostic.Severity != OfficeImageExportDiagnosticSeverity.Info);
                throw new InvalidOperationException(
                    $"Visual corpus slide {index + 1} image export reported: " +
                    string.Join(" | ", failures.Select(failure => failure.Message)) +
                    $" (without charts: {withoutCharts}; without tables: {withoutTables})");
            }
            File.WriteAllBytes(Path.Combine(imagesDirectory,
                $"slide-{index + 1:00}.png"), image.Bytes);
            File.WriteAllBytes(Path.Combine(vectorsDirectory,
                $"slide-{index + 1:00}.svg"), vectors[index].Bytes);
        }

        byte[] pdf = reopened.ToPdf();
        ValidatePdf(pdf, reopened.Slides.Count);
        File.WriteAllBytes(Path.Combine(root,
            "powerpoint-roadmap-visual-corpus.pdf"), pdf);
        Console.WriteLine($"Created {reopened.Slides.Count} validated slides in {root}");
        return 0;
    }

    private static void ValidatePdf(byte[] pdf, int expectedPageCount) {
        string[] expectedTitles = SmartArtScenarios
            .Select(scenario => scenario.Title + " SmartArt")
            .Concat(new[] { "Shared custom geometry", "Chart and table authoring" })
            .ToArray();
        PdfCore.PdfReadDocument parsed = PdfCore.PdfReadDocument.Open(pdf);
        if (parsed.Pages.Count != expectedPageCount
            || expectedTitles.Length != expectedPageCount) {
            throw new InvalidOperationException(
                $"Visual corpus PDF produced {parsed.Pages.Count} pages; expected {expectedPageCount}.");
        }
        for (int index = 0; index < parsed.Pages.Count; index++) {
            string text = parsed.Pages[index].ExtractText();
            if (text.IndexOf(expectedTitles[index],
                    StringComparison.Ordinal) < 0) {
                throw new InvalidOperationException(
                    $"Visual corpus PDF page {index + 1} lost expected title '{expectedTitles[index]}'.");
            }
        }

        IReadOnlyList<PdfCore.PdfPageRenderResult> rendered =
            PdfCore.PdfDocument.Open(pdf).Read.RenderPages(options:
                new PdfCore.PdfPageRenderOptions {
                    Dpi = 72D,
                    Format = PdfCore.PdfPageRenderFormat.Png,
                    MaxPages = expectedPageCount,
                    ContinueOnError = false,
                    MaxTotalOutputBytes = 256L * 1024L * 1024L
                });
        if (rendered.Count != expectedPageCount) {
            throw new InvalidOperationException(
                $"Visual corpus PDF rendered {rendered.Count} pages; expected {expectedPageCount}.");
        }
        for (int index = 0; index < rendered.Count; index++) {
            PdfCore.PdfPageRenderResult page = rendered[index];
            if (!page.Succeeded || page.Bytes == null
                || !OfficePngReader.TryDecode(page.Bytes,
                    out OfficeRasterImage? raster) || raster == null) {
                throw new InvalidOperationException(
                    $"Visual corpus PDF page {index + 1} could not be rerendered and decoded.");
            }
            OfficeColor background = OfficeColor.FromRgb(248, 250, 252);
            int visible = 0;
            for (int y = 0; y < raster.Height; y++) {
                for (int x = 0; x < raster.Width; x++) {
                    OfficeColor pixel = raster.GetPixel(x, y);
                    if (pixel.A > 0 && Math.Abs(pixel.R - background.R)
                        + Math.Abs(pixel.G - background.G)
                        + Math.Abs(pixel.B - background.B) > 12) {
                        visible++;
                    }
                }
            }
            if (visible < 1000) {
                throw new InvalidOperationException(
                    $"Visual corpus PDF page {index + 1} lost visible content.");
            }
        }
    }

    private static PowerPointSlide AddSmartArtSlide(
        PowerPointPresentation presentation, PowerPointSmartArtType type,
        string title) {
        PowerPointSlide slide = presentation.AddSlide();
        SetSlideBackground(slide);
        AddTitle(slide, title + " SmartArt");
        PowerPointSmartArt smartArt = slide.AddSmartArt(type,
            new[] { "Discover", "Design", "Build", "Validate" },
            PowerPointUnits.FromPoints(60), PowerPointUnits.FromPoints(100),
            PowerPointUnits.FromPoints(840), PowerPointUnits.FromPoints(360));
        smartArt.Name = title + " semantic diagram";
        AddFooter(slide, "Editable semantic diagram · create/save/open/render");
        return slide;
    }

    private static PowerPointSlide AddCustomGeometrySlide(
        PowerPointPresentation presentation) {
        PowerPointSlide slide = presentation.AddSlide();
        SetSlideBackground(slide);
        AddTitle(slide, "Shared custom geometry");

        OfficeShape curved = OfficeShape.Path(
            OfficePathCommand.MoveTo(0, 50),
            OfficePathCommand.QuadraticBezierTo(25, 0, 50, 25),
            OfficePathCommand.CubicBezierTo(70, 45, 82, 100, 100, 50),
            OfficePathCommand.LineTo(76, 100),
            OfficePathCommand.LineTo(24, 100),
            OfficePathCommand.Close());
        curved.FillRule = OfficeFillRule.NonZero;
        curved.FillColor = OfficeColor.FromRgb(14, 165, 233);
        curved.StrokeColor = OfficeColor.FromRgb(12, 74, 110);
        curved.StrokeWidth = 2.25D;
        slide.AddCustomGeometryPoints(curved, 90, 145, 300, 250,
            "Curved freeform");

        OfficeShape polygon = OfficeShape.Polygon(
            new OfficePoint(50, 0), new OfficePoint(100, 38),
            new OfficePoint(82, 100), new OfficePoint(18, 100),
            new OfficePoint(0, 38));
        polygon.FillColor = OfficeColor.FromRgb(168, 85, 247);
        polygon.StrokeColor = OfficeColor.FromRgb(88, 28, 135);
        polygon.StrokeWidth = 2.25D;
        slide.AddCustomGeometryPoints(polygon, 545, 145, 300, 250,
            "Polygon freeform");
        AddFooter(slide, "Shared Drawing geometry · editable PowerPoint custGeom");
        return slide;
    }

    private static PowerPointSlide AddChartAndTableSlide(
        PowerPointPresentation presentation) {
        PowerPointSlide slide = presentation.AddSlide();
        SetSlideBackground(slide);
        AddTitle(slide, "Chart and table authoring");

        PowerPointTable table = slide.AddTablePoints(4, 3, 45, 125, 330, 290);
        string[,] values = {
            { "Metric", "Current", "Target" },
            { "Quality", "94", "98" },
            { "Coverage", "91", "95" },
            { "Latency", "24 ms", "20 ms" }
        };
        for (int row = 0; row < values.GetLength(0); row++) {
            for (int column = 0; column < values.GetLength(1); column++) {
                PowerPointTableCell cell = table.GetCell(row, column);
                cell.Text = values[row, column];
                cell.FontName = "Arial";
                if (row == 0) {
                    cell.Bold = true;
                    cell.FillColor = "DBEAFE";
                }
            }
        }

        var data = new OfficeChartData(
            new[] { "Q1", "Q2", "Q3", "Q4" },
            new[] {
                new OfficeChartSeries("Actual", new[] { 18D, 24D, 33D, 42D }),
                new OfficeChartSeries("Target", new[] { 20D, 28D, 36D, 45D })
            });
        PowerPointChart chart = slide.AddChartPoints(
            OfficeChartKind.ColumnClustered, data, 420, 110, 490, 340);
        chart.SetTitle("Quarterly trajectory")
            .SetTitleTextStyle(fontName: "Arial")
            .SetLegendTextStyle(fontName: "Arial")
            .SetCategoryAxisLabelTextStyle(fontName: "Arial")
            .SetValueAxisLabelTextStyle(fontName: "Arial");
        AddFooter(slide, "Typed table and shared chart data · editable after reopen");
        return slide;
    }

    private static void SetSlideBackground(PowerPointSlide slide) {
        slide.BackgroundColor = "F8FAFC";
    }

    private static void AddTitle(PowerPointSlide slide, string text) {
        PowerPointTextBox title = slide.AddTextBoxPoints(text, 45, 28, 760, 46);
        title.FontName = "Arial";
        title.FontSize = 28;
        title.Bold = true;
        title.Color = "0F172A";
    }

    private static void AddFooter(PowerPointSlide slide, string text) {
        PowerPointTextBox footer = slide.AddTextBoxPoints(text, 45, 495, 760, 20);
        footer.FontName = "Arial";
        footer.FontSize = 10;
        footer.Color = "64748B";
    }
}
