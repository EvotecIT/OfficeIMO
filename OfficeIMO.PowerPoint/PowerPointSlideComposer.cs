using System;
using System.Collections.Generic;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Structured composition surface for building custom designer slides without manual placement for every shape.
    /// </summary>
    public sealed class PowerPointSlideComposer {
        private readonly PowerPointSlide _slide;
        private readonly PowerPointDesignTheme _theme;
        private readonly PowerPointDesignerSlideOptions _options;
        private readonly double _slideWidthCm;
        private readonly double _slideHeightCm;
        private readonly bool _dark;

        internal PowerPointSlideComposer(PowerPointSlide slide, PowerPointDesignTheme theme,
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
        public PowerPointLayoutBox ContentArea(double topCm = 3.65, double bottomMarginCm = 1.65,
            double horizontalMarginCm = 1.5) {
            if (topCm < 0) {
                throw new ArgumentOutOfRangeException(nameof(topCm));
            }
            if (bottomMarginCm < 0) {
                throw new ArgumentOutOfRangeException(nameof(bottomMarginCm));
            }
            if (horizontalMarginCm < 0) {
                throw new ArgumentOutOfRangeException(nameof(horizontalMarginCm));
            }

            double height = _slideHeightCm - topCm - bottomMarginCm;
            double width = _slideWidthCm - horizontalMarginCm * 2;
            if (height <= 0 || width <= 0) {
                throw new ArgumentOutOfRangeException(nameof(topCm), "Content area does not fit on the slide.");
            }

            return PowerPointLayoutBox.FromCentimeters(horizontalMarginCm, topCm, width, height);
        }

        /// <summary>
        ///     Returns columns inside the default content area.
        /// </summary>
        public PowerPointLayoutBox[] ContentColumns(int columnCount, double gutterCm = 0.65,
            double topCm = 3.65, double bottomMarginCm = 1.65, double horizontalMarginCm = 1.5) {
            return ContentArea(topCm, bottomMarginCm, horizontalMarginCm).SplitColumnsCm(columnCount, gutterCm);
        }

        /// <summary>
        ///     Returns rows inside the default content area.
        /// </summary>
        public PowerPointLayoutBox[] ContentRows(int rowCount, double gutterCm = 0.55,
            double topCm = 3.65, double bottomMarginCm = 1.65, double horizontalMarginCm = 1.5) {
            return ContentArea(topCm, bottomMarginCm, horizontalMarginCm).SplitRowsCm(rowCount, gutterCm);
        }

        /// <summary>
        ///     Returns a row/column grid inside the default content area.
        /// </summary>
        public PowerPointLayoutBox[,] ContentGrid(int rowCount, int columnCount, double rowGutterCm = 0.55,
            double columnGutterCm = 0.55, double topCm = 3.65, double bottomMarginCm = 1.65,
            double horizontalMarginCm = 1.5) {
            return ContentArea(topCm, bottomMarginCm, horizontalMarginCm)
                .SplitGridCm(rowCount, columnCount, rowGutterCm, columnGutterCm);
        }

        /// <summary>
        ///     Resolves a named composition preset into reusable regions for a raw designer slide.
        /// </summary>
        public PowerPointCompositionLayout UsePreset(PowerPointCompositionPreset preset = PowerPointCompositionPreset.Auto,
            double topCm = 3.65, double bottomMarginCm = 1.65, double horizontalMarginCm = 1.5) {
            return UsePreset(preset, PowerPointCompositionVariant.Auto, topCm, bottomMarginCm, horizontalMarginCm);
        }

        /// <summary>
        ///     Resolves a named composition preset and variant into reusable regions for a raw designer slide.
        /// </summary>
        public PowerPointCompositionLayout UsePreset(PowerPointCompositionPreset preset,
            PowerPointCompositionVariant variant, double topCm = 3.65, double bottomMarginCm = 1.65,
            double horizontalMarginCm = 1.5) {
            PowerPointCompositionPreset resolvedPreset = ResolveCompositionPreset(preset);
            PowerPointCompositionVariant resolvedVariant = ResolveCompositionVariant(resolvedPreset, variant);
            return PowerPointCompositionLayout.Create(resolvedPreset,
                resolvedVariant,
                ContentArea(topCm, bottomMarginCm, horizontalMarginCm), _options.DesignIntent.Density);
        }

        /// <summary>
        ///     Adds a standard title block using the active theme and slide surface.
        /// </summary>
        public void AddTitle(string title, string? subtitle = null) {
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
        }

        /// <summary>
        ///     Adds a raw themed text block at the requested bounds.
        /// </summary>
        public PowerPointTextBox AddTextBlock(string text, PowerPointLayoutBox bounds, int fontSize = 12,
            bool bold = false, string? color = null) {
            if (text == null) {
                throw new ArgumentNullException(nameof(text));
            }

            string resolvedColor = color ?? (_dark ? _theme.AccentLightColor : _theme.SecondaryTextColor);
            return PowerPointDesignExtensions.AddText(_slide, text, bounds.LeftCm, bounds.TopCm, bounds.WidthCm,
                bounds.HeightCm, fontSize, resolvedColor, _theme.BodyFontName, bold);
        }

        /// <summary>
        ///     Adds a designer card grid using the supplied semantic card content.
        /// </summary>
        public void AddCardGrid(IEnumerable<PowerPointCardContent> cards, PowerPointCardGridSlideOptions? options = null) {
            PowerPointCardGridSlideOptions resolved = options ?? new PowerPointCardGridSlideOptions();
            InheritDesignIntent(resolved);
            List<PowerPointCardContent> cardList = PowerPointDesignExtensions.NormalizeCards(cards);
            PowerPointCardGridLayoutVariant variant =
                PowerPointDesignExtensions.ResolveCardGridVariant(resolved, cardList);

            PowerPointDesignExtensions.AddCardGrid(_slide, _theme, cardList, resolved, variant,
                _slideWidthCm, _slideHeightCm);
        }

        /// <summary>
        ///     Adds a designer card grid inside a caller-selected region.
        /// </summary>
        public void AddCardGrid(IEnumerable<PowerPointCardContent> cards, PowerPointLayoutBox bounds,
            PowerPointCardGridSlideOptions? options = null) {
            PowerPointCardGridSlideOptions resolved = options ?? new PowerPointCardGridSlideOptions();
            InheritDesignIntent(resolved);
            List<PowerPointCardContent> cardList = PowerPointDesignExtensions.NormalizeCards(cards);
            PowerPointCardGridLayoutVariant variant =
                PowerPointDesignExtensions.ResolveCardGridVariant(resolved, cardList);

            PowerPointDesignExtensions.AddCardGrid(_slide, _theme, cardList, resolved, variant, bounds);
        }

        /// <summary>
        ///     Adds a process/timeline primitive using the supplied semantic steps.
        /// </summary>
        public void AddProcessTimeline(IEnumerable<PowerPointProcessStep> steps, PowerPointProcessSlideOptions? options = null) {
            PowerPointProcessSlideOptions resolved = options ?? new PowerPointProcessSlideOptions();
            InheritDesignIntent(resolved);
            List<PowerPointProcessStep> stepList = PowerPointDesignExtensions.NormalizeSteps(steps);
            PowerPointDesignExtensions.AddProcessTimeline(_slide, _theme, stepList, resolved,
                _slideWidthCm, _slideHeightCm);
        }

        /// <summary>
        ///     Adds a process/timeline primitive inside a caller-selected region.
        /// </summary>
        public void AddProcessTimeline(IEnumerable<PowerPointProcessStep> steps, PowerPointLayoutBox bounds,
            PowerPointProcessSlideOptions? options = null) {
            PowerPointProcessSlideOptions resolved = options ?? new PowerPointProcessSlideOptions();
            InheritDesignIntent(resolved);
            List<PowerPointProcessStep> stepList = PowerPointDesignExtensions.NormalizeSteps(steps);
            PowerPointDesignExtensions.AddProcessTimeline(_slide, _theme, stepList, resolved, bounds);
        }

        /// <summary>
        ///     Adds a themed metric strip inside explicit bounds.
        /// </summary>
        public void AddMetricStrip(IEnumerable<PowerPointMetric> metrics, PowerPointLayoutBox bounds) {
            if (metrics == null) {
                throw new ArgumentNullException(nameof(metrics));
            }

            PowerPointAutoShape band = _slide.AddRectangleCm(bounds.LeftCm, bounds.TopCm, bounds.WidthCm,
                bounds.HeightCm, "Composer Metric Band");
            band.FillColor = _dark ? _theme.AccentColor : _theme.AccentColor;
            band.FillTransparency = _dark ? 55 : 0;
            band.OutlineColor = _dark ? _theme.AccentLightColor : _theme.AccentColor;
            band.OutlineWidthPoints = 0.45;
            band.SetShadow("000000", blurPoints: 3, distancePoints: 0.8, angleDegrees: 90, transparencyPercent: 90);

            PowerPointDesignExtensions.AddMetrics(_slide, _theme, metrics.Where(metric => metric != null).ToList(),
                bounds.LeftCm, bounds.TopCm, bounds.WidthCm, bounds.HeightCm);
        }

        /// <summary>
        ///     Adds a polished image or editable placeholder surface inside explicit bounds.
        /// </summary>
        public void AddVisualFrame(PowerPointLayoutBox bounds, string? imagePath = null) {
            PowerPointDesignExtensions.AddVisualFrame(_slide, _theme, imagePath, bounds.LeftCm, bounds.TopCm,
                bounds.WidthCm, bounds.HeightCm, _options.DesignIntent);
        }

        /// <summary>
        ///     Adds a lightweight callout band inside explicit bounds.
        /// </summary>
        public void AddCalloutBand(string text, PowerPointLayoutBox bounds, string? accentColor = null) {
            if (text == null) {
                throw new ArgumentNullException(nameof(text));
            }

            string accent = accentColor ?? _theme.AccentColor;
            PowerPointAutoShape band = _slide.AddRectangleCm(bounds.LeftCm, bounds.TopCm, bounds.WidthCm,
                bounds.HeightCm, "Composer Callout Band");
            band.FillColor = _dark ? _theme.AccentColor : _theme.PanelColor;
            band.FillTransparency = _dark ? 60 : 0;
            band.OutlineColor = _dark ? _theme.AccentLightColor : _theme.PanelBorderColor;
            band.OutlineWidthPoints = 0.45;
            band.SetShadow("000000", blurPoints: 3, distancePoints: 0.8, angleDegrees: 90, transparencyPercent: 90);

            PowerPointAutoShape rule = _slide.AddRectangleCm(bounds.LeftCm, bounds.TopCm, 0.13, bounds.HeightCm,
                "Composer Callout Accent");
            rule.FillColor = accent;
            rule.OutlineColor = accent;

            PowerPointDesignExtensions.AddText(_slide, text, bounds.LeftCm + 0.55, bounds.TopCm + 0.28,
                bounds.WidthCm - 0.85, bounds.HeightCm - 0.45, 13,
                _dark ? _theme.AccentContrastColor : _theme.SecondaryTextColor, _theme.BodyFontName, bold: true);
        }

        /// <summary>
        ///     Adds a logo, partner, or certification wall inside a caller-selected region.
        /// </summary>
        public void AddLogoWall(IEnumerable<PowerPointLogoItem> logos, PowerPointLayoutBox bounds,
            PowerPointLogoWallSlideOptions? options = null) {
            PowerPointLogoWallSlideOptions resolved = options ?? new PowerPointLogoWallSlideOptions();
            InheritDesignIntent(resolved);
            IReadOnlyList<PowerPointLogoItem> logoList = PowerPointDesignExtensions.NormalizeLogoItems(logos);
            PowerPointLogoWallLayoutVariant variant =
                PowerPointDesignExtensions.ResolveLogoWallVariant(resolved, logoList);

            PowerPointDesignExtensions.AddLogoWall(_slide, _theme, logoList, resolved, variant, bounds);
        }

        /// <summary>
        ///     Adds an editable coverage map-like surface with normalized location pins inside explicit bounds.
        /// </summary>
        public void AddCoverageMap(IEnumerable<PowerPointCoverageLocation> locations, PowerPointLayoutBox bounds,
            PowerPointCoverageSlideOptions? options = null) {
            PowerPointCoverageSlideOptions resolved = options ?? new PowerPointCoverageSlideOptions();
            InheritDesignIntent(resolved);
            IReadOnlyList<PowerPointCoverageLocation> locationList = PowerPointDesignExtensions.NormalizeLocations(locations);
            PowerPointDesignExtensions.AddCoverageMap(_slide, _theme, locationList, bounds, resolved);
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

    public static partial class PowerPointDesignExtensions {
        /// <summary>
        ///     Creates a designer slide and exposes reusable primitives for custom composition.
        /// </summary>
        public static PowerPointSlide ComposeDesignerSlide(this PowerPointPresentation presentation,
            Action<PowerPointSlideComposer> compose, PowerPointDesignTheme? theme = null,
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
            double width = presentation.SlideSize.WidthCm;
            double height = presentation.SlideSize.HeightCm;

            slide.BackgroundColor = dark ? resolvedTheme.AccentDarkColor : resolvedTheme.BackgroundColor;
            if (dark && resolvedOptions.DesignIntent.VisualStyle == PowerPointVisualStyle.Geometric) {
                AddDiagonalPlanes(slide, resolvedTheme, width, height, dark: true);
            } else if (!dark && resolvedOptions.DesignIntent.VisualStyle != PowerPointVisualStyle.Minimal) {
                AddSubtleLightBackground(slide, resolvedTheme, width, height);
            }

            AddChrome(slide, resolvedTheme, width, height, dark, resolvedOptions);
            compose(new PowerPointSlideComposer(slide, resolvedTheme, resolvedOptions, width, height, dark));
            return slide;
        }
    }
}
