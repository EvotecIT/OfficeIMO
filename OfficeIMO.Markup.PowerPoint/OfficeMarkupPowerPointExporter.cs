using System.Diagnostics;
using OfficeIMO.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Markup.PowerPoint;

public sealed partial class OfficeMarkupPowerPointExporter {
    private const double SlideWidth = 10.0;
    private const double SlideHeight = 5.625;

    private const double TitleLeft = 0.55;
    private const double TitleTop = 0.32;
    private const double TitleWidth = 8.9;
    private const double TitleHeight = 0.65;

    private const double BodyLeft = 0.75;
    private const double BodyTop = 1.15;
    private const double BodyWidth = 8.5;
    private const double BodyHeight = 3.85;

    private const double DesignerBodyLeft = 0.72;
    private const double DesignerBodyTop = 1.72;
    private const double DesignerBodyWidth = 8.55;
    private const double DesignerBodyHeight = 2.95;

    private readonly struct SlideCanvasMetrics {
        public SlideCanvasMetrics(double width, double height) {
            Width = width > 0 ? width : SlideWidth;
            Height = height > 0 ? height : SlideHeight;
        }

        public double Width { get; }
        public double Height { get; }
        private double ScaleX => Width / SlideWidth;
        private double ScaleY => Height / SlideHeight;

        public double Horizontal(double value) => value * ScaleX;
        public double Vertical(double value) => value * ScaleY;
    }

    public void Export(OfficeMarkupDocument document, OfficeMarkupPowerPointExportOptions options) {
        ExportWithReport(document, options);
    }

    /// <summary>
    /// Exports presentation markup and returns the shared machine-readable PowerPoint generation report.
    /// </summary>
    public PowerPointDeckPreflightReport ExportWithReport(OfficeMarkupDocument document,
        OfficeMarkupPowerPointExportOptions options) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        if (document.Profile != OfficeMarkupProfile.Presentation) {
            throw new InvalidOperationException("PowerPoint export requires the Presentation OfficeIMO markup profile.");
        }

        if (string.IsNullOrWhiteSpace(options.OutputPath)) {
            throw new InvalidOperationException("PowerPoint export requires an output path.");
        }

        var directory = Path.GetDirectoryName(Path.GetFullPath(options.OutputPath));
        if (!string.IsNullOrEmpty(directory)) {
            Directory.CreateDirectory(directory);
        }

        var styleResolver = OfficeMarkupStyleResolver.Create(document);
        var metrics = new SlideCanvasMetrics(options.SlideWidthInches, options.SlideHeightInches);
        using PowerPointPresentation presentation = PowerPointPresentation.Create(options.OutputPath);
        presentation.SlideSize.SetSizeInches(metrics.Width, metrics.Height);
        var deck = presentation.UseDesigner(CreateDeckDesign(document), applyTheme: true);
        string? activeSection = null;
        foreach (var slideBlock in GetSlides(document)) {
            ExportSlide(presentation, deck, slideBlock, options, metrics, styleResolver);
            var section = NormalizeSectionName(slideBlock.Section);
            if (section != null && !string.Equals(activeSection, section, StringComparison.Ordinal)) {
                presentation.AddSection(section, presentation.Slides.Count - 1);
                activeSection = section;
            }
        }

        PowerPointDeckPreflightOptions preflightOptions = options.PreflightOptions?.Clone()
            ?? new PowerPointDeckPreflightOptions();
        PowerPointDeckPreflightReport report = presentation.Preflight(preflightOptions);
        if (!string.IsNullOrWhiteSpace(options.PreflightReportPath)) {
            report.SaveJson(options.PreflightReportPath!);
        }
        if (options.FailOnPreflightFindings) {
            report.ThrowIfFindings(preflightOptions.FailureSeverity);
        }
        presentation.Save();
        return report;
    }
}
