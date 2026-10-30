using System;
using System.Collections.Generic;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Plan-owned callback context for building a custom designed slide without manually placing every shape.
    /// </summary>
    public sealed class PowerPointSlideCompositionContext {
        private readonly PowerPointSlide _slide;
        private readonly PowerPointDesignTheme _theme;
        private readonly PowerPointDesignerSlideOptions _options;
        private readonly double _slideWidthCm;
        private readonly double _slideHeightCm;
        private readonly bool _dark;

        internal PowerPointSlideCompositionContext(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointDesignerSlideOptions options, double slideWidthCm, double slideHeightCm, bool dark) {
            _slide = slide;
            _theme = theme;
            _options = options;
            _slideWidthCm = slideWidthCm;
            _slideHeightCm = slideHeightCm;
            _dark = dark;
        }

        /// <summary>
        ///     Underlying editable slide for advanced callers that need to add custom OfficeIMO shapes.
        /// </summary>
        public PowerPointSlide Slide => _slide;

        /// <summary>
        ///     Returns the default editable content area below the title and above the footer.
        /// </summary>
        public PowerPointLayoutBox ContentArea(OfficeLength? top = null, OfficeLength? bottomMargin = null,
            OfficeLength? horizontalMargin = null) {
            OfficeLength resolvedTop = top ?? OfficeLength.Centimeters(3.65);
            OfficeLength resolvedBottom = bottomMargin ?? OfficeLength.Centimeters(1.65);
            OfficeLength resolvedHorizontal = horizontalMargin ?? OfficeLength.Centimeters(1.5);
            if (resolvedTop < OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(top));
            if (resolvedBottom < OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(bottomMargin));
            if (resolvedHorizontal < OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(horizontalMargin));

            OfficeLength height = OfficeLength.Centimeters(_slideHeightCm) - resolvedTop - resolvedBottom;
            OfficeLength width = OfficeLength.Centimeters(_slideWidthCm) - resolvedHorizontal * 2;
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
            PowerPointCompositionPreset resolvedPreset = ResolveCompositionPreset(preset);
            PowerPointCompositionVariant resolvedVariant = ResolveCompositionVariant(resolvedPreset, variant);
            return PowerPointCompositionLayout.Create(resolvedPreset, resolvedVariant,
                ContentArea(top, bottomMargin, horizontalMargin), _options.DesignIntent.Density);
        }

        /// <summary>
        ///     Adds a standard title block using the active theme and slide surface.
        /// </summary>
        /// <returns>The editable shapes added by this composition, in drawing order.</returns>
        public IReadOnlyList<PowerPointShape> AddTitle(string title, string? subtitle = null) {
            return CaptureAddedShapes(() => {
                if (title == null) {
                    throw new ArgumentNullException(nameof(title));
                }

                string titleColor = _dark ? _theme.AccentContrastColor : _theme.PrimaryTextColor;
                string subtitleColor = _dark ? _theme.AccentLightColor : _theme.SecondaryTextColor;
                PowerPointDesignExtensions.AddText(_slide, title, 1.65, 1.45, _slideWidthCm * 0.62, 1.05, 31,
                    titleColor, _theme.HeadingFontName, bold: true);

                if (!string.IsNullOrWhiteSpace(subtitle)) {
                    PowerPointDesignExtensions.AddText(_slide, subtitle!, 1.7, 2.65, _slideWidthCm * 0.58, 0.55, 12,
                        subtitleColor, _theme.BodyFontName, bold: true);
                }

            });
        }

        /// <summary>
        ///     Adds a themed text block at the requested bounds. The default font size is 12pt.
        /// </summary>
        public PowerPointTextBox AddTextBlock(string text, PowerPointLayoutBox bounds, OfficeLength? fontSize = null,
            bool bold = false, string? color = null) {
            if (text == null) {
                throw new ArgumentNullException(nameof(text));
            }

            double resolvedFontSize = (fontSize ?? OfficeLength.Points(12)).ToPoints();
            _ = PowerPointTextDefaults.ToDrawingFontSize(resolvedFontSize, nameof(fontSize));
            string resolvedColor = color ?? (_dark ? _theme.AccentLightColor : _theme.SecondaryTextColor);
            return PowerPointDesignExtensions.AddText(_slide, text, bounds.Left.ToCentimeters(), bounds.Top.ToCentimeters(), bounds.Width.ToCentimeters(),
                bounds.Height.ToCentimeters(), resolvedFontSize, resolvedColor, _theme.BodyFontName, bold);
        }

        /// <summary>
        ///     Adds a designer card grid using the supplied semantic card content.
        /// </summary>
        /// <returns>The editable shapes added by this composition, in drawing order.</returns>
        public IReadOnlyList<PowerPointShape> AddCardGrid(IEnumerable<PowerPointCardContent> cards, PowerPointCardGridSlideOptions? options = null) {
            return CaptureAddedShapes(() => {
                PowerPointCardGridSlideOptions resolved = options ?? new PowerPointCardGridSlideOptions();
                InheritDesignIntent(resolved);
                List<PowerPointCardContent> cardList = PowerPointDesignExtensions.NormalizeCards(cards);
                PowerPointCardGridLayoutVariant variant =
                    PowerPointDesignExtensions.ResolveCardGridVariant(resolved, cardList);

                PowerPointDesignExtensions.AddCardGrid(_slide, _theme, cardList, resolved, variant,
                    _slideWidthCm, _slideHeightCm);

            });
        }

        /// <summary>
        ///     Adds a designer card grid inside a caller-selected region.
        /// </summary>
        /// <returns>The editable shapes added by this composition, in drawing order.</returns>
        public IReadOnlyList<PowerPointShape> AddCardGrid(IEnumerable<PowerPointCardContent> cards, PowerPointLayoutBox bounds,
            PowerPointCardGridSlideOptions? options = null) {
            return CaptureAddedShapes(() => {
                PowerPointCardGridSlideOptions resolved = options ?? new PowerPointCardGridSlideOptions();
                InheritDesignIntent(resolved);
                List<PowerPointCardContent> cardList = PowerPointDesignExtensions.NormalizeCards(cards);
                PowerPointCardGridLayoutVariant variant =
                    PowerPointDesignExtensions.ResolveCardGridVariant(resolved, cardList);

                PowerPointDesignExtensions.AddCardGrid(_slide, _theme, cardList, resolved, variant, bounds);

            });
        }

        /// <summary>
        ///     Adds a process/timeline primitive using the supplied semantic steps.
        /// </summary>
        /// <returns>The editable shapes added by this composition, in drawing order.</returns>
        public IReadOnlyList<PowerPointShape> AddProcessTimeline(IEnumerable<PowerPointProcessStep> steps, PowerPointProcessSlideOptions? options = null) {
            return CaptureAddedShapes(() => {
                PowerPointProcessSlideOptions resolved = options ?? new PowerPointProcessSlideOptions();
                InheritDesignIntent(resolved);
                List<PowerPointProcessStep> stepList = PowerPointDesignExtensions.NormalizeSteps(steps);
                PowerPointDesignExtensions.AddProcessTimeline(_slide, _theme, stepList, resolved,
                    _slideWidthCm, _slideHeightCm);

            });
        }

        /// <summary>
        ///     Adds a process/timeline primitive inside a caller-selected region.
        /// </summary>
        /// <returns>The editable shapes added by this composition, in drawing order.</returns>
        public IReadOnlyList<PowerPointShape> AddProcessTimeline(IEnumerable<PowerPointProcessStep> steps, PowerPointLayoutBox bounds,
            PowerPointProcessSlideOptions? options = null) {
            return CaptureAddedShapes(() => {
                PowerPointProcessSlideOptions resolved = options ?? new PowerPointProcessSlideOptions();
                InheritDesignIntent(resolved);
                List<PowerPointProcessStep> stepList = PowerPointDesignExtensions.NormalizeSteps(steps);
                PowerPointDesignExtensions.AddProcessTimeline(_slide, _theme, stepList, resolved, bounds);

            });
        }

        /// <summary>
        ///     Adds a themed metric strip inside explicit bounds.
        /// </summary>
        /// <returns>The editable shapes added by this composition, in drawing order.</returns>
        public IReadOnlyList<PowerPointShape> AddMetricStrip(IEnumerable<PowerPointMetric> metrics, PowerPointLayoutBox bounds) {
            return AddMetricStrip(metrics, bounds, PowerPointMetricStripVariant.SolidBand);
        }

        /// <summary>
        ///     Adds a themed metric strip inside explicit bounds using a selected surface variant.
        /// </summary>
        /// <returns>The editable shapes added by this composition, in drawing order.</returns>
        public IReadOnlyList<PowerPointShape> AddMetricStrip(IEnumerable<PowerPointMetric> metrics, PowerPointLayoutBox bounds,
            PowerPointMetricStripVariant variant) {
            return CaptureAddedShapes(() => {
                if (metrics == null) {
                    throw new ArgumentNullException(nameof(metrics));
                }

                List<PowerPointMetric> metricList = metrics.Where(metric => metric != null).ToList();
                if (metricList.Count == 0) {
                    return;
                }

                PowerPointMetricStripVariant resolvedVariant = ResolveMetricStripVariant(variant);
                if (resolvedVariant == PowerPointMetricStripVariant.SeparatedTiles) {
                    AddMetricTiles(metricList, bounds);
                    return;
                }
                if (resolvedVariant == PowerPointMetricStripVariant.Underlined) {
                    AddUnderlinedMetrics(metricList, bounds);
                    return;
                }

                PowerPointAutoShape band = _slide.AddRectangle(bounds.Left, bounds.Top, bounds.Width,
                    bounds.Height, "Composer Metric Band");
                band.FillColor = _dark ? _theme.AccentColor : _theme.AccentColor;
                band.FillTransparency = _dark ? 55 : 0;
                band.OutlineColor = _dark ? _theme.AccentLightColor : _theme.AccentColor;
                band.OutlineWidth = global::OfficeIMO.OfficeLength.Points(0.45);
                band.SetShadow("000000", blur: OfficeIMO.OfficeLength.Points(3), distance: OfficeIMO.OfficeLength.Points(0.8), angleDegrees: 90, transparencyPercent: 90);

                PowerPointDesignExtensions.AddMetrics(_slide, _theme, metricList, bounds.Left.ToCentimeters(), bounds.Top.ToCentimeters(),
                    bounds.Width.ToCentimeters(), bounds.Height.ToCentimeters());

            });
        }

        /// <summary>
        ///     Adds a polished image or editable placeholder surface inside explicit bounds.
        /// </summary>
        /// <returns>The editable shapes added by this composition, in drawing order.</returns>
        public IReadOnlyList<PowerPointShape> AddVisualFrame(PowerPointLayoutBox bounds, string? imagePath = null) {
            return AddVisualFrame(bounds, imagePath, PowerPointVisualFrameVariant.Auto);
        }

        /// <summary>
        ///     Adds a polished editable placeholder surface inside explicit bounds using a selected visual variant.
        /// </summary>
        /// <returns>The editable shapes added by this composition, in drawing order.</returns>
        public IReadOnlyList<PowerPointShape> AddVisualFrame(PowerPointLayoutBox bounds, PowerPointVisualFrameVariant variant) {
            return AddVisualFrame(bounds, null, variant);
        }

        /// <summary>
        ///     Adds a polished image or editable placeholder surface inside explicit bounds using a selected visual variant.
        /// </summary>
        /// <returns>The editable shapes added by this composition, in drawing order.</returns>
        public IReadOnlyList<PowerPointShape> AddVisualFrame(PowerPointLayoutBox bounds, string? imagePath, PowerPointVisualFrameVariant variant) {
            return CaptureAddedShapes(() => {
                PowerPointDesignExtensions.AddVisualFrame(_slide, _theme, imagePath, bounds.Left.ToCentimeters(), bounds.Top.ToCentimeters(),
                    bounds.Width.ToCentimeters(), bounds.Height.ToCentimeters(), variant, _options.DesignIntent);

            });
        }

        /// <summary>
        ///     Adds a lightweight callout band inside explicit bounds.
        /// </summary>
        /// <returns>The editable shapes added by this composition, in drawing order.</returns>
        public IReadOnlyList<PowerPointShape> AddCalloutBand(string text, PowerPointLayoutBox bounds, string? accentColor = null) {
            return CaptureAddedShapes(() => {
                if (text == null) {
                    throw new ArgumentNullException(nameof(text));
                }

                string accent = accentColor ?? _theme.AccentColor;
                PowerPointAutoShape band = _slide.AddRectangle(bounds.Left, bounds.Top, bounds.Width,
                    bounds.Height, "Composer Callout Band");
                band.FillColor = _dark ? _theme.AccentColor : _theme.PanelColor;
                band.FillTransparency = _dark ? 60 : 0;
                band.OutlineColor = _dark ? _theme.AccentLightColor : _theme.PanelBorderColor;
                band.OutlineWidth = global::OfficeIMO.OfficeLength.Points(0.45);
                band.SetShadow("000000", blur: OfficeIMO.OfficeLength.Points(3), distance: OfficeIMO.OfficeLength.Points(0.8), angleDegrees: 90, transparencyPercent: 90);

                PowerPointAutoShape rule = _slide.AddRectangle(bounds.Left, bounds.Top, OfficeIMO.OfficeLength.Centimeters(0.13), bounds.Height,
                    "Composer Callout Accent");
                rule.FillColor = accent;
                rule.OutlineColor = accent;

                PowerPointDesignExtensions.AddText(_slide, text, bounds.Left.ToCentimeters() + 0.55, bounds.Top.ToCentimeters() + 0.28,
                    bounds.Width.ToCentimeters() - 0.85, bounds.Height.ToCentimeters() - 0.45, 13,
                    _dark ? _theme.AccentContrastColor : _theme.SecondaryTextColor, _theme.BodyFontName, bold: true);

            });
        }

        /// <summary>
        ///     Adds a logo, partner, or certification wall inside a caller-selected region.
        /// </summary>
        /// <returns>The editable shapes added by this composition, in drawing order.</returns>
        public IReadOnlyList<PowerPointShape> AddLogoWall(IEnumerable<PowerPointLogoItem> logos, PowerPointLayoutBox bounds,
            PowerPointLogoWallSlideOptions? options = null) {
            return CaptureAddedShapes(() => {
                PowerPointLogoWallSlideOptions resolved = options ?? new PowerPointLogoWallSlideOptions();
                InheritDesignIntent(resolved);
                IReadOnlyList<PowerPointLogoItem> logoList = PowerPointDesignExtensions.NormalizeLogoItems(logos);
                PowerPointLogoWallLayoutVariant variant =
                    PowerPointDesignExtensions.ResolveLogoWallVariant(resolved, logoList);

                PowerPointDesignExtensions.AddLogoWall(_slide, _theme, logoList, resolved, variant, bounds);

            });
        }

        /// <summary>
        ///     Adds an editable coverage map-like surface with normalized location pins inside explicit bounds.
        /// </summary>
        /// <returns>The editable shapes added by this composition, in drawing order.</returns>
        public IReadOnlyList<PowerPointShape> AddCoverageMap(IEnumerable<PowerPointCoverageLocation> locations, PowerPointLayoutBox bounds,
            PowerPointCoverageSlideOptions? options = null) {
            return CaptureAddedShapes(() => {
                PowerPointCoverageSlideOptions resolved = options ?? new PowerPointCoverageSlideOptions();
                InheritDesignIntent(resolved);
                IReadOnlyList<PowerPointCoverageLocation> locationList = PowerPointDesignExtensions.NormalizeLocations(locations);
                PowerPointDesignExtensions.AddCoverageMap(_slide, _theme, locationList, bounds, resolved);

            });
        }

        /// <summary>Returns the attached shapes appended by one composite authoring operation.</summary>
        private IReadOnlyList<PowerPointShape> CaptureAddedShapes(Action add) {
            int first = _slide.Shapes.Count;
            add();
            return _slide.Shapes.Skip(first).ToArray();
        }

        private void InheritDesignIntent(PowerPointDesignerSlideOptions childOptions) {
            if (childOptions.DesignIntent.Seed == null &&
                childOptions.DesignIntent.Mood == PowerPointDesignMood.Corporate &&
                childOptions.DesignIntent.Density == PowerPointSlideDensity.Balanced &&
                childOptions.DesignIntent.VisualStyle == PowerPointVisualStyle.Geometric) {
                childOptions.DesignIntent = _options.DesignIntent;
            }
        }

        private PowerPointCompositionPreset ResolveCompositionPreset(PowerPointCompositionPreset preset) {
            if (preset != PowerPointCompositionPreset.Auto) {
                return preset;
            }

            if (_options.DesignIntent.LayoutStrategy == PowerPointAutoLayoutStrategy.Compact) {
                return PowerPointCompositionPreset.DashboardGrid;
            }
            if (_options.DesignIntent.LayoutStrategy == PowerPointAutoLayoutStrategy.VisualFirst) {
                return StablePick(_options.DesignIntent.Seed ?? "visual", 2) == 0
                    ? PowerPointCompositionPreset.VisualSplit
                    : PowerPointCompositionPreset.MetricStory;
            }
            if (_options.DesignIntent.VisualStyle == PowerPointVisualStyle.Minimal) {
                return PowerPointCompositionPreset.BalancedColumns;
            }
            if (_options.DesignIntent.LayoutStrategy == PowerPointAutoLayoutStrategy.DesignFirst) {
                PowerPointCompositionPreset[] choices = {
                    PowerPointCompositionPreset.BalancedColumns,
                    PowerPointCompositionPreset.VisualSplit,
                    PowerPointCompositionPreset.MetricStory,
                    PowerPointCompositionPreset.DashboardGrid
                };
                return choices[StablePick(_options.DesignIntent.Seed ?? "composition", choices.Length)];
            }

            return _options.DesignIntent.Mood == PowerPointDesignMood.Energetic
                ? PowerPointCompositionPreset.MetricStory
                : PowerPointCompositionPreset.BalancedColumns;
        }

        private PowerPointCompositionVariant ResolveCompositionVariant(PowerPointCompositionPreset preset,
            PowerPointCompositionVariant variant) {
            if (variant != PowerPointCompositionVariant.Auto) {
                return variant;
            }

            if (_options.DesignIntent.VisualStyle == PowerPointVisualStyle.Minimal) {
                return PowerPointCompositionVariant.Standard;
            }
            if (_options.DesignIntent.LayoutStrategy == PowerPointAutoLayoutStrategy.VisualFirst) {
                return PowerPointCompositionVariant.VisualLead;
            }
            if (_options.DesignIntent.LayoutStrategy == PowerPointAutoLayoutStrategy.Compact ||
                preset == PowerPointCompositionPreset.DashboardGrid) {
                return PowerPointCompositionVariant.EvidenceLead;
            }

            PowerPointCompositionVariant[] choices = {
                PowerPointCompositionVariant.Standard,
                PowerPointCompositionVariant.Mirrored,
                PowerPointCompositionVariant.VisualLead,
                PowerPointCompositionVariant.EvidenceLead
            };
            string seed = _options.DesignIntent.Seed ?? preset.ToString();
            return choices[StablePick(seed + ":variant", choices.Length)];
        }

        private PowerPointMetricStripVariant ResolveMetricStripVariant(PowerPointMetricStripVariant variant) {
            if (variant != PowerPointMetricStripVariant.Auto) {
                return variant;
            }
            if (_options.DesignIntent.VisualStyle == PowerPointVisualStyle.Minimal) {
                return PowerPointMetricStripVariant.Underlined;
            }
            if (_options.DesignIntent.VisualStyle == PowerPointVisualStyle.Soft ||
                _options.DesignIntent.Mood == PowerPointDesignMood.Editorial ||
                _options.DesignIntent.LayoutStrategy == PowerPointAutoLayoutStrategy.Compact) {
                return PowerPointMetricStripVariant.SeparatedTiles;
            }

            PowerPointMetricStripVariant[] choices = {
                PowerPointMetricStripVariant.SolidBand,
                PowerPointMetricStripVariant.SeparatedTiles,
                PowerPointMetricStripVariant.Underlined
            };
            string seed = _options.DesignIntent.Seed ?? "metric-strip";
            return choices[StablePick(seed + ":metric-surface", choices.Length)];
        }

        private void AddMetricTiles(IReadOnlyList<PowerPointMetric> metrics, PowerPointLayoutBox bounds) {
            int count = Math.Min(metrics.Count, 3);
            PowerPointLayoutBox[] boxes = bounds.SplitColumns(count,OfficeIMO.OfficeLength.Centimeters(0.34));
            for (int i = 0; i < count; i++) {
                PowerPointLayoutBox box = boxes[i];
                string accent = GetComposerAccent(i);
                PowerPointAutoShape tile = _slide.AddRectangle(box.Left, box.Top, box.Width, box.Height,
                    "Composer Metric Tile " + (i + 1));
                tile.FillColor = accent;
                tile.FillTransparency = _dark ? 55 : 8;
                tile.OutlineColor = accent;
                tile.OutlineWidth = global::OfficeIMO.OfficeLength.Points(0.45);
                tile.SetShadow("000000", blur: OfficeIMO.OfficeLength.Points(2.8), distance: OfficeIMO.OfficeLength.Points(0.7), angleDegrees: 90, transparencyPercent: 90);
            }

            PowerPointDesignExtensions.AddMetrics(_slide, _theme, metrics, bounds.Left.ToCentimeters(), bounds.Top.ToCentimeters(),
                bounds.Width.ToCentimeters(), bounds.Height.ToCentimeters());
        }

        private void AddUnderlinedMetrics(IReadOnlyList<PowerPointMetric> metrics, PowerPointLayoutBox bounds) {
            int count = Math.Min(metrics.Count, 3);
            PowerPointLayoutBox[] boxes = bounds.SplitColumns(count,OfficeIMO.OfficeLength.Centimeters(0.45));
            for (int i = 0; i < count; i++) {
                PowerPointLayoutBox box = boxes[i];
                PowerPointMetric metric = metrics[i];
                string valueColor = _dark ? _theme.AccentContrastColor : _theme.AccentDarkColor;
                string labelColor = _dark ? _theme.AccentLightColor : _theme.SecondaryTextColor;
                int valueFontSize = ResolveComposerMetricFontSize(metric.Value, box.Width.ToCentimeters(), box.Height.ToCentimeters() < 1.4 ? 21 : 25);
                double valueHeight = Math.Max(0.46, box.Height.ToCentimeters() * 0.42);
                valueFontSize = PowerPointDesignExtensions.FitSingleLineFontToHeight(valueFontSize, valueHeight);
                int labelFontSize = ResolveComposerMetricFontSize(metric.Label, box.Width.ToCentimeters(), 9);

                PowerPointTextBox value = PowerPointDesignExtensions.AddText(_slide, metric.Value,
                    box.Left.ToCentimeters(), box.Top.ToCentimeters() + Math.Max(0.05, box.Height.ToCentimeters() * 0.08), box.Width.ToCentimeters(),
                    valueHeight, valueFontSize, valueColor, _theme.HeadingFontName, bold: true);
                CenterComposerText(value);

                double underlineWidth = Math.Min(box.Width.ToCentimeters() * 0.46, 1.4);
                PowerPointAutoShape underline = _slide.AddRectangle(
                    OfficeIMO.OfficeLength.Centimeters(box.Left.ToCentimeters() + (box.Width.ToCentimeters() - underlineWidth) / 2),
                    OfficeIMO.OfficeLength.Centimeters(box.Top.ToCentimeters() + box.Height.ToCentimeters() * 0.54), OfficeIMO.OfficeLength.Centimeters(underlineWidth), OfficeIMO.OfficeLength.Centimeters(0.05),
                    "Composer Metric Underline " + (i + 1));
                underline.FillColor = GetComposerAccent(i);
                underline.OutlineColor = underline.FillColor;

                PowerPointTextBox label = PowerPointDesignExtensions.AddText(_slide, metric.Label,
                    box.Left.ToCentimeters() + 0.08, box.Top.ToCentimeters() + box.Height.ToCentimeters() * 0.62, box.Width.ToCentimeters() - 0.16,
                    Math.Max(0.35, box.Height.ToCentimeters() * 0.28), labelFontSize, labelColor, _theme.BodyFontName, bold: true);
                CenterComposerText(label);
            }
        }

        private string GetComposerAccent(int index) {
            switch (index % 4) {
                case 1:
                    return _theme.Accent3Color;
                case 2:
                    return _theme.Accent2Color;
                case 3:
                    return _theme.WarningColor;
                default:
                    return _theme.AccentColor;
            }
        }

        private static int ResolveComposerMetricFontSize(string text, double widthCm, int preferredFontSize) {
            int length = string.IsNullOrWhiteSpace(text) ? 0 : text.Trim().Length;
            if (length <= 4) {
                return preferredFontSize;
            }

            int estimate = (int)Math.Floor(widthCm * 7.0 / Math.Max(1, length));
            return Math.Max(8, Math.Min(preferredFontSize, estimate));
        }

        private static void CenterComposerText(PowerPointTextBox textBox) {
            foreach (PowerPointParagraph paragraph in textBox.Paragraphs) {
                paragraph.Alignment = PowerPointTextAlignment.Center;
            }
            textBox.TextVerticalAlignment = PowerPointTextVerticalAlignment.Center;
        }

        private static int StablePick(string value, int choices) {
            unchecked {
                int hash = (int)2166136261;
                for (int i = 0; i < value.Length; i++) {
                    hash ^= value[i];
                    hash *= 16777619;
                }

                return (hash & int.MaxValue) % choices;
            }
        }
    }

    internal static partial class PowerPointDesignExtensions {
        /// <summary>
        ///     Creates a designer slide and exposes reusable primitives for custom composition.
        /// </summary>
        public static PowerPointSlide ComposeDesignerSlide(this PowerPointDocument presentation,
            Action<PowerPointSlideCompositionContext> compose, PowerPointDesignTheme? theme = null,
            PowerPointDesignerSlideOptions? options = null, bool dark = false) {
            if (presentation == null) {
                throw new ArgumentNullException(nameof(presentation));
            }
            if (compose == null) {
                throw new ArgumentNullException(nameof(compose));
            }

            PowerPointDesignTheme resolvedTheme = ResolveTheme(theme);
            PowerPointDesignerSlideOptions resolvedOptions = options ?? new PowerPointDesignerSlideOptions();
            PowerPointSlide slide = presentation.AddSlide();
            double width = presentation.SlideSize.Width.ToCentimeters();
            double height = presentation.SlideSize.Height.ToCentimeters();

            slide.BackgroundColor = dark ? resolvedTheme.AccentDarkColor : resolvedTheme.BackgroundColor;
            if (dark && resolvedOptions.DesignIntent.VisualStyle == PowerPointVisualStyle.Geometric) {
                AddDiagonalPlanes(slide, resolvedTheme, width, height, dark: true);
            } else if (!dark && resolvedOptions.DesignIntent.VisualStyle != PowerPointVisualStyle.Minimal) {
                AddSubtleLightBackground(slide, resolvedTheme, width, height);
            }

            AddChrome(slide, resolvedTheme, width, height, dark, resolvedOptions);
            compose(new PowerPointSlideCompositionContext(slide, resolvedTheme, resolvedOptions, width, height, dark));
            return slide;
        }
    }
}
