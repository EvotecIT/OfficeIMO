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
            new PowerPointLayoutBox(OfficeLength.Points(rect.X), OfficeLength.Points(rect.Y), OfficeLength.Points(rect.Width), OfficeLength.Points(rect.Height));

        private static PowerPointLayoutRect ToRect(PowerPointLayoutBox box) =>
            new(box.Left.ToPoints(), box.Top.ToPoints(), box.Width.ToPoints(), box.Height.ToPoints());

        /// <summary>
        ///     Returns the editable content area below the title and above the footer. With the default arguments this
        ///     is the layout grid's content zone; explicit arguments describe a custom area in centimetres.
        /// </summary>
        public PowerPointLayoutBox ContentArea(OfficeLength? top = null, OfficeLength? bottomMargin = null,
            OfficeLength? horizontalMargin = null) {
            OfficeLength resolvedTop = top ?? OfficeLength.Centimeters(DefaultTopCm);
            OfficeLength resolvedBottom = bottomMargin ?? OfficeLength.Centimeters(DefaultBottomMarginCm);
            OfficeLength resolvedHorizontal = horizontalMargin ?? OfficeLength.Centimeters(DefaultHorizontalMarginCm);
            if (resolvedTop < OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(top));
            if (resolvedBottom < OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(bottomMargin));
            if (resolvedHorizontal < OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(horizontalMargin));

            if (!top.HasValue && !bottomMargin.HasValue && !horizontalMargin.HasValue) return ToBox(Zone);
            OfficeLength height = _slide.OwnerDocument!.SlideSize.Height - resolvedTop - resolvedBottom;
            OfficeLength width = _slide.OwnerDocument!.SlideSize.Width - resolvedHorizontal * 2;
            if (height <= OfficeLength.Zero || width <= OfficeLength.Zero) {
                throw new ArgumentOutOfRangeException(nameof(top), "Content area does not fit on the slide.");
            }
            return new PowerPointLayoutBox(resolvedHorizontal, resolvedTop, width, height);
        }

        /// <summary>Returns columns inside the default content area. The default gutter is 0.65 cm.</summary>
        public PowerPointLayoutBox[] ContentColumns(int columnCount, OfficeLength? gutter = null,
            OfficeLength? top = null, OfficeLength? bottomMargin = null, OfficeLength? horizontalMargin = null) =>
            ContentArea(top, bottomMargin, horizontalMargin)
                .SplitColumns(columnCount, gutter ?? OfficeLength.Centimeters(0.65));

        /// <summary>Returns rows inside the default content area. The default gutter is 0.55 cm.</summary>
        public PowerPointLayoutBox[] ContentRows(int rowCount, OfficeLength? gutter = null,
            OfficeLength? top = null, OfficeLength? bottomMargin = null, OfficeLength? horizontalMargin = null) =>
            ContentArea(top, bottomMargin, horizontalMargin)
                .SplitRows(rowCount, gutter ?? OfficeLength.Centimeters(0.55));

        /// <summary>Returns a grid inside the default content area. Both default gutters are 0.55 cm.</summary>
        public PowerPointLayoutBox[,] ContentGrid(int rowCount, int columnCount, OfficeLength? rowGutter = null,
            OfficeLength? columnGutter = null, OfficeLength? top = null, OfficeLength? bottomMargin = null,
            OfficeLength? horizontalMargin = null) =>
            ContentArea(top, bottomMargin, horizontalMargin).SplitGrid(rowCount, columnCount,
                rowGutter ?? OfficeLength.Centimeters(0.55), columnGutter ?? OfficeLength.Centimeters(0.55));

        /// <summary>Resolves a named composition preset and variant into reusable regions for a raw designer slide.</summary>
        public PowerPointCompositionLayout UsePreset(PowerPointCompositionPreset preset = PowerPointCompositionPreset.Auto,
            PowerPointCompositionVariant variant = PowerPointCompositionVariant.Auto, OfficeLength? top = null,
            OfficeLength? bottomMargin = null, OfficeLength? horizontalMargin = null) {
            PowerPointCompositionPreset resolvedPreset = preset == PowerPointCompositionPreset.Auto ? PowerPointCompositionPreset.BalancedColumns : preset;
            PowerPointCompositionVariant resolvedVariant = variant == PowerPointCompositionVariant.Auto ? PowerPointCompositionVariant.Standard : variant;
            return PowerPointCompositionLayout.Create(resolvedPreset, resolvedVariant,
                ContentArea(top, bottomMargin, horizontalMargin), _options.DesignIntent.Density);
        }

        /// <summary>Adds the deck's standard slide header: accent mark, eyebrow, title and subtitle.</summary>
        public IReadOnlyList<PowerPointShape> AddTitle(string title, string? subtitle = null) {
            return CaptureAddedShapes(() => {
            if (title == null) throw new ArgumentNullException(nameof(title));
            var slide = new PowerPointLayoutSlide { Title = title, Subtitle = subtitle, Eyebrow = _options.Eyebrow };
            if (!string.IsNullOrWhiteSpace(_context.Deck.CurrentSection)) slide.Eyebrow = null;
            _header = PowerPointLayoutEngine.RenderHeader(_context, _writer, slide);
                    });
        }

        /// <summary>
        ///     Adds a text block at the requested bounds in the body font. Sizes below the theme's small-body minimum
        ///     are raised to it so text stays readable.
        /// </summary>
        public PowerPointTextBox AddTextBlock(string text, PowerPointLayoutBox bounds, OfficeLength? fontSize = null,
            bool bold = false, string? color = null) {
            if (text == null) throw new ArgumentNullException(nameof(text));
            if (fontSize.HasValue && fontSize.Value <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(fontSize));
            PowerPointTextStyleSpec style = _context.Style(PowerPointTextRole.BodySmall);
            double size = Math.Max(style.Minimum, (fontSize ?? OfficeLength.Points(12)).ToPoints());
            PowerPointLayoutParagraph paragraph = _context.Paragraph(text, PowerPointTextRole.BodySmall,
                _context.ResolveColor(PowerPointLayoutBlock.ParseColor(color)) ?? _context.Theme.Text, size, bold: bold);
            return _writer.Text("Text Block", ToRect(bounds), paragraph);
        }

        /// <summary>Adds the Cards composition in the content area.</summary>
        public IReadOnlyList<PowerPointShape> AddCardGrid(IEnumerable<PowerPointCardContent> cards, PowerPointCardGridSlideOptions? options = null) =>
            AddCardGrid(cards, ToBox(Zone), options);

        /// <summary>Adds the Cards composition inside a caller-selected region.</summary>
        public IReadOnlyList<PowerPointShape> AddCardGrid(IEnumerable<PowerPointCardContent> cards, PowerPointLayoutBox bounds,
            PowerPointCardGridSlideOptions? options = null) {
            return CaptureAddedShapes(() => {
            List<PowerPointCardContent> list = PowerPointDesignExtensions.NormalizeCards(cards);
            var slide = new PowerPointLayoutSlide();
            foreach (PowerPointCardContent card in list) {
                slide.Blocks.Add(new PowerPointLayoutCard(card.Title, card.Text, card.Items, card.ImagePath, PowerPointLayoutBlock.ParseColor(card.AccentColor)));
            }

            PowerPointLayoutEngine.DrawCards(_context, _writer, slide, ToRect(bounds), full: false, maxColumns: options?.MaxColumns);
                    });
        }

        /// <summary>Adds the Process composition in the content area.</summary>
        public IReadOnlyList<PowerPointShape> AddProcessTimeline(IEnumerable<PowerPointProcessStep> steps, PowerPointProcessSlideOptions? options = null) =>
            AddProcessTimeline(steps, ToBox(Zone), options);

        /// <summary>Adds the Process composition inside a caller-selected region.</summary>
        public IReadOnlyList<PowerPointShape> AddProcessTimeline(IEnumerable<PowerPointProcessStep> steps, PowerPointLayoutBox bounds,
            PowerPointProcessSlideOptions? options = null) {
            return CaptureAddedShapes(() => {
            List<PowerPointProcessStep> list = PowerPointDesignExtensions.NormalizeSteps(steps);
            var slide = new PowerPointLayoutSlide();
            foreach (PowerPointProcessStep step in list) slide.Blocks.Add(new PowerPointLayoutStep(step.Title, step.Body, step.Number));
            PowerPointLayoutEngine.DrawProcess(_context, _writer, slide, ToRect(bounds), full: false);
                    });
        }

        /// <summary>Adds KPI tiles inside explicit bounds.</summary>
        public IReadOnlyList<PowerPointShape> AddMetricStrip(IEnumerable<PowerPointMetric> metrics, PowerPointLayoutBox bounds) {
            return CaptureAddedShapes(() => {
            AddMetricStrip(metrics, bounds, PowerPointMetricStripVariant.Auto);
                    });
        }

        /// <summary>Adds KPI tiles inside explicit bounds. The theme layouts draw one tile design; the variant is kept for compatibility.</summary>
        public IReadOnlyList<PowerPointShape> AddMetricStrip(IEnumerable<PowerPointMetric> metrics, PowerPointLayoutBox bounds,
            PowerPointMetricStripVariant variant) {
            return CaptureAddedShapes(() => {
            if (metrics == null) throw new ArgumentNullException(nameof(metrics));
            var slide = new PowerPointLayoutSlide();
            foreach (PowerPointMetric metric in metrics.Where(m => m != null)) {
                slide.Blocks.Add(new PowerPointLayoutKpi(metric.Value, metric.Label, metric.Change, PowerPointLayoutBlock.ParseColor(metric.Color)));
            }

            if (slide.Blocks.Count == 0) return;
            PowerPointLayoutEngine.DrawKpi(_context, _writer, slide, ToRect(bounds), full: false);
                    });
        }

        /// <summary>Adds an image (contained) or, without one, a surface placeholder inside explicit bounds.</summary>
        public IReadOnlyList<PowerPointShape> AddVisualFrame(PowerPointLayoutBox bounds, string? imagePath = null) {
            return CaptureAddedShapes(() => {
            AddVisualFrame(bounds, imagePath, PowerPointVisualFrameVariant.Auto);
                    });
        }

        /// <summary>Adds a surface placeholder inside explicit bounds.</summary>
        public IReadOnlyList<PowerPointShape> AddVisualFrame(PowerPointLayoutBox bounds, PowerPointVisualFrameVariant variant) {
            return CaptureAddedShapes(() => {
            AddVisualFrame(bounds, null, variant);
                    });
        }

        /// <summary>Adds an image (contained) or, without one, a surface placeholder inside explicit bounds.</summary>
        public IReadOnlyList<PowerPointShape> AddVisualFrame(PowerPointLayoutBox bounds, string? imagePath, PowerPointVisualFrameVariant variant) {
            return CaptureAddedShapes(() => {
            PowerPointLayoutRect rect = ToRect(bounds);
            if (!string.IsNullOrWhiteSpace(imagePath) && System.IO.File.Exists(imagePath)) {
                PowerPointLayoutEngine.RenderImage(_context, _writer, new PowerPointLayoutImage(imagePath!, null), rect, cover: false);
                return;
            }

            _writer.Rect("Visual Placeholder", rect, _context.Theme.Surface, outline: _context.Theme.Hairline, outlineWidth: 1D);
                    });
        }

        /// <summary>Adds a callout: a surface panel with an accent rule and lede text.</summary>
        public IReadOnlyList<PowerPointShape> AddCalloutBand(string text, PowerPointLayoutBox bounds, string? accentColor = null) {
            return CaptureAddedShapes(() => {
            if (text == null) throw new ArgumentNullException(nameof(text));
            PowerPointLayoutRect rect = ToRect(bounds);
            _writer.Rect("Callout Panel", rect, _context.Theme.Surface);
            _writer.Rect("Callout Rule", new PowerPointLayoutRect(rect.X, rect.Y, 3D, rect.Height), _context.ResolveColor(PowerPointLayoutBlock.ParseColor(accentColor)) ?? _context.Theme.Accent1);
            PowerPointTextStyleSpec lede = _context.Style(PowerPointTextRole.Lede);
            PowerPointLayoutRect inner = rect.Inset(20D, 12D, 16D, 12D);
            double size = lede.Size;
            while (size - 1D >= lede.Minimum - 2D && _context.Height(text, lede, inner.Width, size) > inner.Height) size -= 1D;
            _writer.Text("Callout", inner, _context.Paragraph(text, PowerPointTextRole.Lede, size: size), anchor: "ctr");
                    });
        }

        /// <summary>Adds the Logos composition inside a caller-selected region.</summary>
        public IReadOnlyList<PowerPointShape> AddLogoWall(IEnumerable<PowerPointLogoItem> logos, PowerPointLayoutBox bounds,
            PowerPointLogoWallSlideOptions? options = null) {
            return CaptureAddedShapes(() => {
            IReadOnlyList<PowerPointLogoItem> list = PowerPointDesignExtensions.NormalizeLogoItems(logos);
            var slide = new PowerPointLayoutSlide();
            foreach (PowerPointLogoItem logo in list) slide.Blocks.Add(new PowerPointLayoutCard(logo.Name, logo.Subtitle, image: logo.ImagePath, color: PowerPointLayoutBlock.ParseColor(logo.AccentColor)));
            PowerPointLayoutEngine.DrawLogos(_context, _writer, slide, ToRect(bounds), full: false);
                    });
        }

        /// <summary>Adds the Coverage composition (map with numbered pins and the location list) inside explicit bounds.</summary>
        public IReadOnlyList<PowerPointShape> AddCoverageMap(IEnumerable<PowerPointCoverageLocation> locations, PowerPointLayoutBox bounds,
            PowerPointCoverageSlideOptions? options = null) {
            return CaptureAddedShapes(() => {
            IReadOnlyList<PowerPointCoverageLocation> list = PowerPointDesignExtensions.NormalizeLocations(locations);
            var slide = new PowerPointLayoutSlide { MapLabel = options?.MapLabel };
            foreach (PowerPointCoverageLocation location in list) {
                slide.Blocks.Add(new PowerPointLayoutLocation(location.Name, location.X, location.Y, location.Detail));
            }

            PowerPointLayoutEngine.DrawCoverage(_context, _writer, slide, ToRect(bounds), full: false);
                    });
        }
        /// <summary>Returns the attached shapes appended by one composite authoring operation.</summary>
        private IReadOnlyList<PowerPointShape> CaptureAddedShapes(Action add) {
            int first = _slide.Shapes.Count;
            add();
            return _slide.Shapes.Skip(first).ToArray();
        }
    }

    internal static partial class PowerPointDesignExtensions {
        /// <summary>
        ///     Creates a designed slide with the deck footer and exposes layout primitives for custom composition.
        ///     <paramref name="dark" /> is kept for compatibility; custom slides use the light content surface.
        /// </summary>
        public static PowerPointSlide ComposeDesignerSlide(this PowerPointDocument presentation,
            Action<PowerPointSlideCompositionContext> compose, PowerPointDesignTheme? theme = null,
            PowerPointDesignerSlideOptions? options = null, bool dark = false) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            if (compose == null) throw new ArgumentNullException(nameof(compose));
            PowerPointDesignerSlideOptions resolvedOptions = options ?? new PowerPointDesignerSlideOptions();
            var slide = new PowerPointLayoutSlide { Layout = PowerPointLayout.TitleAndContent, TemplateLayout = resolvedOptions.TemplateLayout };
            slide.Blocks.Add(new PowerPointLayoutDrawing((target, context, box) =>
                compose(new PowerPointSlideCompositionContext(target, context, resolvedOptions))));
            IReadOnlyList<PowerPointSlide> slides = PowerPointLayoutEngine.AddSlides(presentation, slide, LayoutThemeFor(presentation, theme));
            return FinalizeDesignerAccessibility(slides[0]);
        }
    }
}
