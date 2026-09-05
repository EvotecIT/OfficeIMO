using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint.Pdf;
using PdfCore = OfficeIMO.Pdf;
using System.Globalization;
using System.Text;
using System.Xml.Linq;

namespace OfficeIMO.PowerPoint.Benchmarks;

internal static class PowerPointRoadmapVisualCorpus {
    private readonly struct VisualSubjectRegion {
        internal VisualSubjectRegion(string name, double left, double top,
            double width, double height, int minimumPaintedPixels) {
            Name = name;
            Left = left;
            Top = top;
            Width = width;
            Height = height;
            MinimumPaintedPixels = minimumPaintedPixels;
        }

        internal string Name { get; }
        internal double Left { get; }
        internal double Top { get; }
        internal double Width { get; }
        internal double Height { get; }
        internal int MinimumPaintedPixels { get; }
    }

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
        ValidateScenarioSemantics(reopened);

        IReadOnlyList<OfficeImageExportResult> images = reopened.ExportImages(
            OfficeImageExportFormat.Png);
        IReadOnlyList<OfficeImageExportResult> vectors = reopened.ExportImages(
            OfficeImageExportFormat.Svg);
        if (images.Count != reopened.Slides.Count
            || vectors.Count != reopened.Slides.Count) {
            throw new InvalidOperationException(
                $"Visual corpus expected {reopened.Slides.Count} PNG and SVG exports but received {images.Count} PNG and {vectors.Count} SVG artifacts.");
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
            ValidatePng(image.Bytes, index + 1);
            ValidateSvg(vectors[index].Bytes, index + 1);
            File.WriteAllBytes(Path.Combine(imagesDirectory,
                $"slide-{index + 1:00}.png"), image.Bytes);
            File.WriteAllBytes(Path.Combine(vectorsDirectory,
                $"slide-{index + 1:00}.svg"), vectors[index].Bytes);
        }

        byte[] pdf = reopened.ToPdfBytes();
        ValidatePdf(pdf, reopened.Slides.Count);
        File.WriteAllBytes(Path.Combine(root,
            "powerpoint-roadmap-visual-corpus.pdf"), pdf);
        Console.WriteLine($"Created {reopened.Slides.Count} validated slides in {root}");
        return 0;
    }

    private static void ValidatePng(byte[] png, int slideNumber) {
        if (!OfficePngReader.TryDecode(png, out OfficeRasterImage? raster)
            || raster == null || raster.Width <= 0 || raster.Height <= 0) {
            throw new InvalidOperationException(
                $"Visual corpus PNG slide {slideNumber} could not be decoded with valid dimensions.");
        }
        OfficeColor background = OfficeColor.FromRgb(248, 250, 252);
        int visible = PowerPointBenchmarkVisualValidator
            .CountPixelsDifferentFrom(raster, background,
                0D, 0D, 960D, 540D);
        if (visible < 1000) {
            throw new InvalidOperationException(
                $"Visual corpus PNG slide {slideNumber} lost visible content.");
        }
        ValidateSubjectRegions(raster, slideNumber, "PNG");
    }

    private static void ValidateSvg(byte[] svg, int slideNumber) {
        XDocument document;
        try {
            document = XDocument.Parse(Encoding.UTF8.GetString(svg),
                LoadOptions.None);
        } catch (Exception exception) when (exception is InvalidOperationException
            || exception is System.Xml.XmlException) {
            throw new InvalidOperationException(
                $"Visual corpus SVG slide {slideNumber} is not valid XML.",
                exception);
        }
        XElement? root = document.Root;
        string[] viewBox = (root?.Attribute("viewBox")?.Value
                ?? string.Empty).Split(new[] { ' ', ',' },
                StringSplitOptions.RemoveEmptyEntries);
        bool validDimensions = viewBox.Length == 4
            && double.TryParse(viewBox[2], NumberStyles.Float,
                CultureInfo.InvariantCulture, out double width) && width > 0D
            && double.TryParse(viewBox[3], NumberStyles.Float,
                CultureInfo.InvariantCulture, out double height) && height > 0D;
        if (root?.Name.LocalName != "svg" || !validDimensions) {
            throw new InvalidOperationException(
                $"Visual corpus SVG slide {slideNumber} has no valid canvas.");
        }
        if (!OfficeSvgDrawingReader.TryRead(svg, out OfficeDrawing? drawing,
                out int unsupportedFeatureCount) || drawing == null
            || unsupportedFeatureCount != 0) {
            throw new InvalidOperationException(
                $"Visual corpus SVG slide {slideNumber} could not be fully projected through the shared vector reader.");
        }
        OfficeColor background = OfficeColor.FromRgb(248, 250, 252);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing,
            scale: 1D, background: background);
        int visible = PowerPointBenchmarkVisualValidator
            .CountPixelsDifferentFrom(raster, background,
                0D, 0D, 960D, 540D);
        if (visible < 1000) {
            throw new InvalidOperationException(
                $"Visual corpus SVG slide {slideNumber} lost visible painted content.");
        }
        ValidateSubjectRegions(raster, slideNumber, "SVG");
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
            PdfCore.PdfDocument.Load(pdf).Render.Pages(options:
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
            ValidateSubjectRegions(raster, index + 1, "PDF");
        }
    }

    private static void ValidateScenarioSemantics(
        PowerPointPresentation presentation) {
        int expectedSlideCount = SmartArtScenarios.Length + 2;
        if (presentation.Slides.Count != expectedSlideCount) {
            throw new InvalidOperationException(
                $"Visual corpus contains {presentation.Slides.Count} slides; expected {expectedSlideCount}.");
        }
        string[] expectedNodes = { "Discover", "Design", "Build", "Validate" };
        for (int index = 0; index < SmartArtScenarios.Length; index++) {
            PowerPointSmartArt smartArt = presentation.Slides[index]
                .SmartArts.Single();
            if (!smartArt.GetNodeTexts().SequenceEqual(expectedNodes,
                    StringComparer.Ordinal)) {
                throw new InvalidOperationException(
                    $"Visual corpus SmartArt slide {index + 1} lost its semantic nodes.");
            }
        }

        PowerPointSlide geometrySlide = presentation
            .Slides[SmartArtScenarios.Length];
        string[] geometryNames = geometrySlide.Shapes
            .Select(shape => shape.Name ?? string.Empty)
            .Where(name => name is "Curved freeform" or "Polygon freeform")
            .OrderBy(name => name, StringComparer.Ordinal)
            .ToArray();
        if (!geometryNames.SequenceEqual(
                new[] { "Curved freeform", "Polygon freeform" },
                StringComparer.Ordinal)) {
            throw new InvalidOperationException(
                "Visual corpus custom-geometry slide lost an authored subject.");
        }

        PowerPointSlide chartSlide = presentation.Slides[expectedSlideCount - 1];
        PowerPointTable table = chartSlide.Tables.Single();
        if (!string.Equals(table.GetCell(1, 0).Text, "Quality",
                StringComparison.Ordinal)
            || !string.Equals(table.GetCell(3, 2).Text, "20 ms",
                StringComparison.Ordinal)
            || !chartSlide.Charts.Single().TryGetOfficeSnapshot(
                out OfficeChartSnapshot chart)
            || chart.Data.Series.Count != 2
            || chart.Data.Categories.Count != 4) {
            throw new InvalidOperationException(
                "Visual corpus chart-and-table slide lost its authored data.");
        }
    }

    private static void ValidateSubjectRegions(OfficeRasterImage raster,
        int slideNumber, string artifactKind) {
        foreach (VisualSubjectRegion region in GetSubjectRegions(slideNumber)) {
            OfficeColor background = OfficeColor.FromRgb(248, 250, 252);
            int visible = PowerPointBenchmarkVisualValidator
                .CountPixelsDifferentFrom(raster, background,
                    region.Left, region.Top, region.Width, region.Height);
            if (visible < region.MinimumPaintedPixels) {
                throw new InvalidOperationException(
                    $"Visual corpus {artifactKind} slide {slideNumber} lost the {region.Name} subject region ({visible} painted pixels; expected at least {region.MinimumPaintedPixels}).");
            }
        }
    }

    private static IReadOnlyList<VisualSubjectRegion> GetSubjectRegions(
        int slideNumber) {
        if (slideNumber >= 1 && slideNumber <= SmartArtScenarios.Length) {
            return new[] {
                new VisualSubjectRegion("SmartArt", 60, 100, 840, 360, 200)
            };
        }
        if (slideNumber == SmartArtScenarios.Length + 1) {
            return new[] {
                new VisualSubjectRegion("curved custom geometry", 90, 145,
                    300, 250, 200),
                new VisualSubjectRegion("polygon custom geometry", 545, 145,
                    300, 250, 200)
            };
        }
        if (slideNumber == SmartArtScenarios.Length + 2) {
            return new[] {
                new VisualSubjectRegion("table", 45, 125, 330, 290, 200),
                new VisualSubjectRegion("chart", 420, 110, 490, 340, 200)
            };
        }
        throw new ArgumentOutOfRangeException(nameof(slideNumber));
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
