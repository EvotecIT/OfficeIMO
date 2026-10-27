using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Plan-owned callback context for building a custom designed slide without manually placing every shape.
    ///     Regions and primitives come from the deck's layout grid and theme, so a custom slide shares the header,
    ///     type scale and tile designs of every other slide.
    /// </summary>
    public sealed class PowerPointSlideCompositionContext {
        private const double DefaultTopCm = 3.65;
        private const double DefaultBottomMarginCm = 1.65;
        private const double DefaultHorizontalMarginCm = 1.5;
        private readonly PowerPointSlide _slide;
        private readonly PowerPointDesignerSlideOptions _options;
        private readonly PowerPointLayoutContext _context;
        private readonly PowerPointLayoutWriter _writer;
        private PowerPointLayoutEngine.HeaderGeometry? _header;

        internal PowerPointSlideCompositionContext(PowerPointSlide slide, PowerPointLayoutContext context,
            PowerPointDesignerSlideOptions options) {
            _slide = slide;
            _context = context;
            _options = options;
            _writer = new PowerPointLayoutWriter(slide, context.Theme);
        }

        /// <summary>
        ///     Underlying editable slide for advanced callers that need to add custom OfficeIMO shapes.
        /// </summary>
        public PowerPointSlide Slide => _slide;

        private PowerPointLayoutRect Zone {
            get {
                PowerPointLayoutGrid grid = _context.Grid;
                int row = _header?.ContentRow ?? 4;
                return grid.Box(1, 12, row, PowerPointLayoutGrid.Rows + 1 - row);
            }
        }

        private static PowerPointLayoutBox ToBox(PowerPointLayoutRect rect) =>
            PowerPointLayoutBox.FromPoints(rect.X, rect.Y, rect.Width, rect.Height);

        private static PowerPointLayoutRect ToRect(PowerPointLayoutBox box) =>
            new(box.LeftPoints, box.TopPoints, box.WidthPoints, box.HeightPoints);

        /// <summary>
        ///     Returns the editable content area below the title and above the footer. With the default arguments this
        ///     is the layout grid's content zone; explicit arguments describe a custom area in centimetres.
        /// </summary>
        public PowerPointLayoutBox ContentArea(double topCm = DefaultTopCm, double bottomMarginCm = DefaultBottomMarginCm,
            double horizontalMarginCm = DefaultHorizontalMarginCm) {
            if (topCm < 0) throw new ArgumentOutOfRangeException(nameof(topCm));
            if (bottomMarginCm < 0) throw new ArgumentOutOfRangeException(nameof(bottomMarginCm));
            if (horizontalMarginCm < 0) throw new ArgumentOutOfRangeException(nameof(horizontalMarginCm));
            if (topCm == DefaultTopCm && bottomMarginCm == DefaultBottomMarginCm && horizontalMarginCm == DefaultHorizontalMarginCm) {
                return ToBox(Zone);
            }

            // Never let a custom area start inside the header the context already drew.
            if (_header != null) topCm = Math.Max(topCm, PowerPointUnits.ToCentimeters(PowerPointUnits.FromPoints(Zone.Y)));
            double heightCm = _slide.OwnerPresentation!.SlideSize.HeightCm - topCm - bottomMarginCm;
            double widthCm = _slide.OwnerPresentation!.SlideSize.WidthCm - horizontalMarginCm * 2;
            if (heightCm <= 0 || widthCm <= 0) {
                throw new ArgumentOutOfRangeException(nameof(topCm), "Content area does not fit on the slide.");
            }

            return PowerPointLayoutBox.FromCentimeters(horizontalMarginCm, topCm, widthCm, heightCm);
        }

        /// <summary>Returns columns inside the content area.</summary>
        public PowerPointLayoutBox[] ContentColumns(int columnCount, double gutterCm = 0.65,
            double topCm = DefaultTopCm, double bottomMarginCm = DefaultBottomMarginCm, double horizontalMarginCm = DefaultHorizontalMarginCm) {
            return ContentArea(topCm, bottomMarginCm, horizontalMarginCm).SplitColumnsCm(columnCount, gutterCm);
        }

        /// <summary>Returns rows inside the content area.</summary>
        public PowerPointLayoutBox[] ContentRows(int rowCount, double gutterCm = 0.55,
            double topCm = DefaultTopCm, double bottomMarginCm = DefaultBottomMarginCm, double horizontalMarginCm = DefaultHorizontalMarginCm) {
            return ContentArea(topCm, bottomMarginCm, horizontalMarginCm).SplitRowsCm(rowCount, gutterCm);
        }

        /// <summary>Returns a row/column grid inside the content area.</summary>
        public PowerPointLayoutBox[,] ContentGrid(int rowCount, int columnCount, double rowGutterCm = 0.55,
            double columnGutterCm = 0.55, double topCm = DefaultTopCm, double bottomMarginCm = DefaultBottomMarginCm,
            double horizontalMarginCm = DefaultHorizontalMarginCm) {
            return ContentArea(topCm, bottomMarginCm, horizontalMarginCm)
                .SplitGridCm(rowCount, columnCount, rowGutterCm, columnGutterCm);
        }

        /// <summary>Resolves a named composition preset into reusable regions of the content area.</summary>
        public PowerPointCompositionLayout UsePreset(PowerPointCompositionPreset preset = PowerPointCompositionPreset.Auto,
            double topCm = DefaultTopCm, double bottomMarginCm = DefaultBottomMarginCm, double horizontalMarginCm = DefaultHorizontalMarginCm) {
            return UsePreset(preset, PowerPointCompositionVariant.Auto, topCm, bottomMarginCm, horizontalMarginCm);
        }

        /// <summary>Resolves a named composition preset and variant into reusable regions of the content area.</summary>
        public PowerPointCompositionLayout UsePreset(PowerPointCompositionPreset preset,
            PowerPointCompositionVariant variant, double topCm = DefaultTopCm, double bottomMarginCm = DefaultBottomMarginCm,
            double horizontalMarginCm = DefaultHorizontalMarginCm) {
            PowerPointCompositionPreset resolvedPreset = preset == PowerPointCompositionPreset.Auto
                ? PowerPointCompositionPreset.BalancedColumns
                : preset;
            PowerPointCompositionVariant resolvedVariant = variant == PowerPointCompositionVariant.Auto
                ? PowerPointCompositionVariant.Standard
                : variant;
            return PowerPointCompositionLayout.Create(resolvedPreset, resolvedVariant,
                ContentArea(topCm, bottomMarginCm, horizontalMarginCm), _options.DesignIntent.Density);
        }

        /// <summary>Adds the deck's standard slide header: accent mark, eyebrow, title and subtitle.</summary>
        public void AddTitle(string title, string? subtitle = null) {
            if (title == null) throw new ArgumentNullException(nameof(title));
            var slide = new PowerPointLayoutSlide { Title = title, Subtitle = subtitle, Eyebrow = _options.Eyebrow };
            if (!string.IsNullOrWhiteSpace(_context.Deck.CurrentSection)) slide.Eyebrow = null;
            _header = PowerPointLayoutEngine.RenderHeader(_context, _writer, slide);
        }

        /// <summary>
        ///     Adds a text block at the requested bounds in the body font. Sizes below the theme's small-body minimum
        ///     are raised to it so text stays readable.
        /// </summary>
        public PowerPointTextBox AddTextBlock(string text, PowerPointLayoutBox bounds, int fontSize = 12,
            bool bold = false, string? color = null) {
            if (text == null) throw new ArgumentNullException(nameof(text));
            PowerPointTextStyleSpec style = _context.Style(PowerPointTextRole.BodySmall);
            double size = Math.Max(style.Minimum, fontSize);
            PowerPointLayoutParagraph paragraph = _context.Paragraph(text, PowerPointTextRole.BodySmall,
                _context.ResolveColor(color) ?? _context.Theme.Text, size, bold: bold);
            return _writer.Text("Text Block", ToRect(bounds), paragraph);
        }

        /// <summary>Adds the Cards composition in the content area.</summary>
        public void AddCardGrid(IEnumerable<PowerPointCardContent> cards, PowerPointCardGridSlideOptions? options = null) =>
            AddCardGrid(cards, ToBox(Zone), options);

        /// <summary>Adds the Cards composition inside a caller-selected region.</summary>
        public void AddCardGrid(IEnumerable<PowerPointCardContent> cards, PowerPointLayoutBox bounds,
            PowerPointCardGridSlideOptions? options = null) {
            List<PowerPointCardContent> list = PowerPointDesignExtensions.NormalizeCards(cards);
            var slide = new PowerPointLayoutSlide();
            foreach (PowerPointCardContent card in list) {
                slide.Blocks.Add(new PowerPointLayoutCard(card.Title, card.Text, card.Items, card.ImagePath, card.AccentColor));
            }

            PowerPointLayoutEngine.DrawCards(_context, _writer, slide, ToRect(bounds), full: false);
        }

        /// <summary>Adds the Process composition in the content area.</summary>
        public void AddProcessTimeline(IEnumerable<PowerPointProcessStep> steps, PowerPointProcessSlideOptions? options = null) =>
            AddProcessTimeline(steps, ToBox(Zone), options);

        /// <summary>Adds the Process composition inside a caller-selected region.</summary>
        public void AddProcessTimeline(IEnumerable<PowerPointProcessStep> steps, PowerPointLayoutBox bounds,
            PowerPointProcessSlideOptions? options = null) {
            List<PowerPointProcessStep> list = PowerPointDesignExtensions.NormalizeSteps(steps);
            var slide = new PowerPointLayoutSlide();
            foreach (PowerPointProcessStep step in list) slide.Blocks.Add(new PowerPointLayoutStep(step.Title, step.Body, step.Number));
            PowerPointLayoutEngine.DrawProcess(_context, _writer, slide, ToRect(bounds), full: false);
        }

        /// <summary>Adds KPI tiles inside explicit bounds.</summary>
        public void AddMetricStrip(IEnumerable<PowerPointMetric> metrics, PowerPointLayoutBox bounds) {
            AddMetricStrip(metrics, bounds, PowerPointMetricStripVariant.Auto);
        }

        /// <summary>Adds KPI tiles inside explicit bounds. The theme layouts draw one tile design; the variant is kept for compatibility.</summary>
        public void AddMetricStrip(IEnumerable<PowerPointMetric> metrics, PowerPointLayoutBox bounds,
            PowerPointMetricStripVariant variant) {
            if (metrics == null) throw new ArgumentNullException(nameof(metrics));
            var slide = new PowerPointLayoutSlide();
            foreach (PowerPointMetric metric in metrics.Where(m => m != null)) {
                slide.Blocks.Add(new PowerPointLayoutKpi(metric.Value, metric.Label, metric.Change, metric.Color));
            }

            if (slide.Blocks.Count == 0) return;
            PowerPointLayoutEngine.DrawKpi(_context, _writer, slide, ToRect(bounds), full: false);
        }

        /// <summary>Adds an image (contained) or, without one, a surface placeholder inside explicit bounds.</summary>
        public void AddVisualFrame(PowerPointLayoutBox bounds, string? imagePath = null) {
            AddVisualFrame(bounds, imagePath, PowerPointVisualFrameVariant.Auto);
        }

        /// <summary>Adds a surface placeholder inside explicit bounds.</summary>
        public void AddVisualFrame(PowerPointLayoutBox bounds, PowerPointVisualFrameVariant variant) {
            AddVisualFrame(bounds, null, variant);
        }

        /// <summary>Adds an image (contained) or, without one, a surface placeholder inside explicit bounds.</summary>
        public void AddVisualFrame(PowerPointLayoutBox bounds, string? imagePath, PowerPointVisualFrameVariant variant) {
            PowerPointLayoutRect rect = ToRect(bounds);
            if (!string.IsNullOrWhiteSpace(imagePath) && System.IO.File.Exists(imagePath)) {
                PowerPointLayoutEngine.RenderImage(_context, _writer, new PowerPointLayoutImage(imagePath!, null), rect, cover: false);
                return;
            }

            _writer.Rect("Visual Placeholder", rect, _context.Theme.Surface, outline: _context.Theme.Hairline, outlineWidth: 1D);
        }

        /// <summary>Adds a callout: a surface panel with an accent rule and lede text.</summary>
        public void AddCalloutBand(string text, PowerPointLayoutBox bounds, string? accentColor = null) {
            if (text == null) throw new ArgumentNullException(nameof(text));
            PowerPointLayoutRect rect = ToRect(bounds);
            _writer.Rect("Callout Panel", rect, _context.Theme.Surface);
            _writer.Rect("Callout Rule", new PowerPointLayoutRect(rect.X, rect.Y, 3D, rect.Height), _context.ResolveColor(accentColor) ?? _context.Theme.Accent1);
            PowerPointTextStyleSpec lede = _context.Style(PowerPointTextRole.Lede);
            PowerPointLayoutRect inner = rect.Inset(20D, 12D, 16D, 12D);
            double size = lede.Size;
            while (size - 1D >= lede.Minimum - 2D && _context.Height(text, lede, inner.Width, size) > inner.Height) size -= 1D;
            _writer.Text("Callout", inner, _context.Paragraph(text, PowerPointTextRole.Lede, size: size), anchor: "ctr");
        }

        /// <summary>Adds the Logos composition inside a caller-selected region.</summary>
        public void AddLogoWall(IEnumerable<PowerPointLogoItem> logos, PowerPointLayoutBox bounds,
            PowerPointLogoWallSlideOptions? options = null) {
            IReadOnlyList<PowerPointLogoItem> list = PowerPointDesignExtensions.NormalizeLogoItems(logos);
            var slide = new PowerPointLayoutSlide();
            foreach (PowerPointLogoItem logo in list) slide.Blocks.Add(new PowerPointLayoutCard(logo.Name, logo.Subtitle, image: logo.ImagePath, color: logo.AccentColor));
            PowerPointLayoutEngine.DrawLogos(_context, _writer, slide, ToRect(bounds), full: false);
        }

        /// <summary>Adds the Coverage composition (map with numbered pins and the location list) inside explicit bounds.</summary>
        public void AddCoverageMap(IEnumerable<PowerPointCoverageLocation> locations, PowerPointLayoutBox bounds,
            PowerPointCoverageSlideOptions? options = null) {
            IReadOnlyList<PowerPointCoverageLocation> list = PowerPointDesignExtensions.NormalizeLocations(locations);
            var slide = new PowerPointLayoutSlide { MapLabel = options?.MapLabel };
            foreach (PowerPointCoverageLocation location in list) {
                slide.Blocks.Add(new PowerPointLayoutLocation(location.Name, location.X, location.Y, location.Detail));
            }

            PowerPointLayoutEngine.DrawCoverage(_context, _writer, slide, ToRect(bounds), full: false);
        }
    }

    internal static partial class PowerPointDesignExtensions {
        /// <summary>
        ///     Creates a designed slide with the deck footer and exposes layout primitives for custom composition.
        ///     <paramref name="dark" /> is kept for compatibility; custom slides use the light content surface.
        /// </summary>
        public static PowerPointSlide ComposeDesignerSlide(this PowerPointPresentation presentation,
            Action<PowerPointSlideCompositionContext> compose, PowerPointDesignTheme? theme = null,
            PowerPointDesignerSlideOptions? options = null, bool dark = false) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            if (compose == null) throw new ArgumentNullException(nameof(compose));
            PowerPointDesignerSlideOptions resolvedOptions = options ?? new PowerPointDesignerSlideOptions();
            var slide = new PowerPointLayoutSlide { Layout = PowerPointLayoutKind.TitleAndContent, TemplateLayout = resolvedOptions.TemplateLayout };
            slide.Blocks.Add(new PowerPointLayoutDrawing((target, context, box) =>
                compose(new PowerPointSlideCompositionContext(target, context, resolvedOptions))));
            IReadOnlyList<PowerPointSlide> slides = PowerPointLayoutEngine.AddSlides(presentation, slide, LayoutThemeFor(presentation, theme));
            return FinalizeDesignerAccessibility(slides[0]);
        }
    }
}
