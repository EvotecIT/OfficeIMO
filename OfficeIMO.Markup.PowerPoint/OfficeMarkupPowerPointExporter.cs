using System.Diagnostics;
using OfficeIMO.PowerPoint;
using OfficeIMO.Drawing.Internal;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Markup.PowerPoint;

internal sealed partial class OfficeMarkupPowerPointExporter {
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

    internal OfficeMarkupPowerPointConversionResult Build(
        OfficeMarkupDocument document,
        MarkupToPowerPointOptions options,
        IReadOnlyList<OfficeMarkupDiagnostic> diagnostics) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        if (document.Profile != OfficeMarkupProfile.Presentation) {
            throw new InvalidOperationException("PowerPoint export requires the Presentation OfficeIMO markup profile.");
        }

        var styleResolver = OfficeMarkupStyleResolver.Create(document);
        var metrics = new SlideCanvasMetrics(options.SlideWidthInches, options.SlideHeightInches);
        PowerPointPresentation presentation = PowerPointPresentation.Create();
        try {
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
                ?? new PowerPointDeckPreflightOptions {
                    // Markup can expand into many peer shapes. Callers may opt into collision diagnostics,
                    // but the default export path must not perform a quadratic pairwise scan.
                    DetectShapeCollisions = false
                };
            PowerPointDeckPreflightReport preflightReport = presentation.InspectPreflight(preflightOptions);
            if (options.FailOnPreflightFindings) {
                preflightReport.ThrowIfFindings(preflightOptions.FailureSeverity);
            }

            return new OfficeMarkupPowerPointConversionResult(presentation, diagnostics, preflightReport);
        } catch {
            presentation.Dispose();
            throw;
        }
    }
}
