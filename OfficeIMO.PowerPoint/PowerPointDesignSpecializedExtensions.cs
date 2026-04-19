using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public static partial class PowerPointDesignExtensions {
        /// <summary>
        ///     Adds a logo, partner, or certification wall slide with optional proof/certificate emphasis.
        /// </summary>
        public static PowerPointSlide AddDesignerLogoWallSlide(this PowerPointPresentation presentation,
            string title, string? subtitle, IEnumerable<PowerPointLogoItem> logos,
            PowerPointDesignTheme? theme = null, PowerPointLogoWallSlideOptions? options = null) {
            if (presentation == null) {
                throw new ArgumentNullException(nameof(presentation));
            }
            if (title == null) {
                throw new ArgumentNullException(nameof(title));
            }

            PowerPointDesignTheme resolvedTheme = ResolveTheme(theme);
            PowerPointLogoWallSlideOptions resolvedOptions = options ?? new PowerPointLogoWallSlideOptions();
            List<PowerPointLogoItem> logoList = NormalizeLogoItems(logos);

            PowerPointSlide slide = presentation.AddSlide();
            double width = presentation.SlideSize.WidthCm;
            double height = presentation.SlideSize.HeightCm;
            slide.BackgroundColor = resolvedTheme.BackgroundColor;

            AddSubtleLightBackground(slide, resolvedTheme, width, height);
            AddChrome(slide, resolvedTheme, width, height, dark: false, resolvedOptions);
            AddText(slide, title, 1.5, 1.45, width * 0.58, 1.0, 29,
                resolvedTheme.PrimaryTextColor, resolvedTheme.HeadingFontName, bold: true);
            if (!string.IsNullOrWhiteSpace(subtitle)) {
                AddText(slide, subtitle!, 1.55, 2.7, width * 0.58, 0.55, 12,
                    resolvedTheme.SecondaryTextColor, resolvedTheme.BodyFontName, bold: true);
            }

            PowerPointLogoWallLayoutVariant variant = ResolveLogoWallVariant(resolvedOptions, logoList);
            if (variant == PowerPointLogoWallLayoutVariant.CertificateFeature) {
                AddLogoCertificateFeature(slide, resolvedTheme, logoList, resolvedOptions, width, height);
            } else {
                AddLogoMosaic(slide, resolvedTheme, logoList, resolvedOptions, width, height);
            }

            return slide;
        }

        /// <summary>
        ///     Adds a coverage/location slide with editable pins and a structured location list.
        /// </summary>
        public static PowerPointSlide AddDesignerCoverageSlide(this PowerPointPresentation presentation,
            string title, string? subtitle, IEnumerable<PowerPointCoverageLocation> locations,
            PowerPointDesignTheme? theme = null, PowerPointCoverageSlideOptions? options = null) {
            if (presentation == null) {
                throw new ArgumentNullException(nameof(presentation));
            }
            if (title == null) {
                throw new ArgumentNullException(nameof(title));
            }

            PowerPointDesignTheme resolvedTheme = ResolveTheme(theme);
            PowerPointCoverageSlideOptions resolvedOptions = options ?? new PowerPointCoverageSlideOptions();
            List<PowerPointCoverageLocation> locationList = NormalizeLocations(locations);

            PowerPointSlide slide = presentation.AddSlide();
            double width = presentation.SlideSize.WidthCm;
            double height = presentation.SlideSize.HeightCm;
            slide.BackgroundColor = resolvedTheme.BackgroundColor;

            AddSubtleLightBackground(slide, resolvedTheme, width, height);
            AddChrome(slide, resolvedTheme, width, height, dark: false, resolvedOptions);
            AddText(slide, title, 1.5, 1.45, width * 0.58, 1.0, 29,
                resolvedTheme.PrimaryTextColor, resolvedTheme.HeadingFontName, bold: true);
            if (!string.IsNullOrWhiteSpace(subtitle)) {
                AddText(slide, subtitle!, 1.55, 2.7, width * 0.58, 0.55, 12,
                    resolvedTheme.SecondaryTextColor, resolvedTheme.BodyFontName, bold: true);
            }

            PowerPointCoverageLayoutVariant variant = ResolveCoverageVariant(resolvedOptions, locationList);
            if (variant == PowerPointCoverageLayoutVariant.ListMap) {
                PowerPointLayoutBox[] columns = PowerPointLayoutBox
                    .FromCentimeters(1.5, 4.05, width - 3.0, height - 5.9)
                    .SplitColumnsCm(2, 0.85);
                AddCoverageList(slide, resolvedTheme, locationList, columns[0], resolvedOptions.SupportingText);
                AddCoverageMap(slide, resolvedTheme, locationList, columns[1], resolvedOptions);
            } else {
                AddCoverageMap(slide, resolvedTheme, locationList,
                    PowerPointLayoutBox.FromCentimeters(2.0, 4.0, width - 4.0, height - 7.1), resolvedOptions);
                AddCoverageStrip(slide, resolvedTheme, locationList,
                    PowerPointLayoutBox.FromCentimeters(2.0, height - 2.55, width - 4.0, 1.05));
            }

            return slide;
        }

        /// <summary>
        ///     Adds a capability/content slide with structured narrative sections and optional visual support.
        /// </summary>
        public static PowerPointSlide AddDesignerCapabilitySlide(this PowerPointPresentation presentation,
            string title, string? subtitle, IEnumerable<PowerPointCapabilitySection> sections,
            PowerPointDesignTheme? theme = null, PowerPointCapabilitySlideOptions? options = null) {
            if (presentation == null) {
                throw new ArgumentNullException(nameof(presentation));
            }
            if (title == null) {
                throw new ArgumentNullException(nameof(title));
            }

            PowerPointDesignTheme resolvedTheme = ResolveTheme(theme);
            PowerPointCapabilitySlideOptions resolvedOptions = options ?? new PowerPointCapabilitySlideOptions();
            List<PowerPointCapabilitySection> sectionList = NormalizeCapabilitySections(sections);

            PowerPointSlide slide = presentation.AddSlide();
            double width = presentation.SlideSize.WidthCm;
            double height = presentation.SlideSize.HeightCm;
            slide.BackgroundColor = resolvedTheme.BackgroundColor;

            AddSubtleLightBackground(slide, resolvedTheme, width, height);
            AddChrome(slide, resolvedTheme, width, height, dark: false, resolvedOptions);
            AddText(slide, title, 1.5, 1.45, width * 0.60, 1.0, 29,
                resolvedTheme.PrimaryTextColor, resolvedTheme.HeadingFontName, bold: true);
            if (!string.IsNullOrWhiteSpace(subtitle)) {
                AddText(slide, subtitle!, 1.55, 2.7, width * 0.64, 0.55, 12,
                    resolvedTheme.SecondaryTextColor, resolvedTheme.BodyFontName, bold: true);
            }

            PowerPointCapabilityLayoutVariant variant = ResolveCapabilityVariant(resolvedOptions, sectionList);
            if (variant == PowerPointCapabilityLayoutVariant.Stacked) {
                AddCapabilityStacked(slide, resolvedTheme, sectionList, resolvedOptions, width, height);
            } else {
                AddCapabilitySplit(slide, resolvedTheme, sectionList, resolvedOptions, width, height,
                    visualFirst: variant == PowerPointCapabilityLayoutVariant.VisualText);
            }

            return slide;
        }

        internal static void AddLogoWall(PowerPointSlide slide, PowerPointDesignTheme theme,
            IReadOnlyList<PowerPointLogoItem> logos, PowerPointLogoWallSlideOptions options,
            PowerPointLogoWallLayoutVariant variant, PowerPointLayoutBox bounds) {
            if (variant == PowerPointLogoWallLayoutVariant.CertificateFeature) {
                PowerPointLayoutBox[] columns = bounds.SplitColumnsCm(2, 0.7);
                AddLogoGrid(slide, theme, logos, columns[0], options);
                AddCertificatePanel(slide, theme, columns[1], options);
                return;
            }

            AddLogoGrid(slide, theme, logos, bounds, options);
        }

        internal static void AddCoverageMap(PowerPointSlide slide, PowerPointDesignTheme theme,
            IReadOnlyList<PowerPointCoverageLocation> locations, PowerPointLayoutBox bounds,
            PowerPointCoverageSlideOptions options) {
            PowerPointAutoShape panel = slide.AddRectangleCm(bounds.LeftCm, bounds.TopCm, bounds.WidthCm,
                bounds.HeightCm, "Coverage Map Panel");
            panel.FillColor = theme.AccentDarkColor;
            panel.FillTransparency = 0;
            panel.OutlineColor = theme.AccentDarkColor;
            panel.OutlineWidthPoints = 0;
            panel.SetShadow("000000", blurPoints: 5, distancePoints: 1.2, angleDegrees: 90, transparencyPercent: 86);

            PowerPointAutoShape wash = slide.AddShapeCm(A.ShapeTypeValues.Parallelogram, bounds.LeftCm + bounds.WidthCm * 0.16,
                bounds.TopCm, bounds.WidthCm * 0.28, bounds.HeightCm, "Coverage Map Wash");
            wash.FillColor = theme.AccentLightColor;
            wash.FillTransparency = 72;
            wash.OutlineColor = theme.AccentLightColor;
            wash.OutlineWidthPoints = 0;

            PowerPointAutoShape gridLine1 = slide.AddLineCm(bounds.LeftCm + bounds.WidthCm * 0.16,
                bounds.TopCm + bounds.HeightCm * 0.33, bounds.LeftCm + bounds.WidthCm * 0.88,
                bounds.TopCm + bounds.HeightCm * 0.33, "Coverage Map Latitude 1");
            gridLine1.OutlineColor = theme.AccentLightColor;
            gridLine1.OutlineWidthPoints = 0.35;
            PowerPointAutoShape gridLine2 = slide.AddLineCm(bounds.LeftCm + bounds.WidthCm * 0.18,
                bounds.TopCm + bounds.HeightCm * 0.66, bounds.LeftCm + bounds.WidthCm * 0.84,
                bounds.TopCm + bounds.HeightCm * 0.66, "Coverage Map Latitude 2");
            gridLine2.OutlineColor = theme.AccentLightColor;
            gridLine2.OutlineWidthPoints = 0.35;
            PowerPointAutoShape gridLine3 = slide.AddLineCm(bounds.LeftCm + bounds.WidthCm * 0.38,
                bounds.TopCm + bounds.HeightCm * 0.08, bounds.LeftCm + bounds.WidthCm * 0.34,
                bounds.TopCm + bounds.HeightCm * 0.92, "Coverage Map Longitude 1");
            gridLine3.OutlineColor = theme.AccentLightColor;
            gridLine3.OutlineWidthPoints = 0.35;
            PowerPointAutoShape gridLine4 = slide.AddLineCm(bounds.LeftCm + bounds.WidthCm * 0.64,
                bounds.TopCm + bounds.HeightCm * 0.12, bounds.LeftCm + bounds.WidthCm * 0.58,
                bounds.TopCm + bounds.HeightCm * 0.88, "Coverage Map Longitude 2");
            gridLine4.OutlineColor = theme.AccentLightColor;
            gridLine4.OutlineWidthPoints = 0.35;

            AddCoverageRegion(slide, theme, bounds, 0.12, 0.23, 0.31, 0.27, "Coverage Region North", theme.AccentLightColor, 78);
            AddCoverageRegion(slide, theme, bounds, 0.43, 0.18, 0.36, 0.25, "Coverage Region East", theme.Accent2Color, 74);
            AddCoverageRegion(slide, theme, bounds, 0.20, 0.55, 0.34, 0.24, "Coverage Region South", theme.AccentLightColor, 82);
            AddCoverageRegion(slide, theme, bounds, 0.58, 0.52, 0.28, 0.25, "Coverage Region Central", theme.Accent2Color, 78);

            int routePoints = Math.Min(locations.Count, 6);
            for (int i = 0; i < routePoints - 1; i++) {
                AddCoverageRoute(slide, theme, bounds, locations[i], locations[i + 1], i);
            }

            int maxPins = Math.Min(locations.Count, 18);
            for (int i = 0; i < maxPins; i++) {
                AddCoveragePin(slide, theme, bounds, locations[i], i);
            }

            if (!string.IsNullOrWhiteSpace(options.MapLabel)) {
                AddText(slide, options.MapLabel!, bounds.LeftCm + 0.55, bounds.TopCm + bounds.HeightCm - 0.85,
                    bounds.WidthCm - 1.1, 0.45, 10, theme.AccentContrastColor, theme.BodyFontName, bold: true);
            }
        }

        private static void AddCapabilitySplit(PowerPointSlide slide, PowerPointDesignTheme theme,
            IReadOnlyList<PowerPointCapabilitySection> sections, PowerPointCapabilitySlideOptions options,
            double slideWidthCm, double slideHeightCm, bool visualFirst) {
            PowerPointLayoutBox[] columns = PowerPointLayoutBox
                .FromCentimeters(1.5, 4.0, slideWidthCm - 3.0, slideHeightCm - 5.7)
                .SplitColumnsCm(2, 0.85);
            PowerPointLayoutBox visual = visualFirst ? columns[0] : columns[1];
            PowerPointLayoutBox narrative = visualFirst ? columns[1] : columns[0];

            double metricReserve = options.Metrics.Count > 0 ? 1.45 : 0;
            PowerPointLayoutBox sectionBounds = metricReserve > 0
                ? PowerPointLayoutBox.FromCentimeters(narrative.LeftCm, narrative.TopCm,
                    narrative.WidthCm, narrative.HeightCm - metricReserve)
                : narrative;

            AddCapabilitySections(slide, theme, sections, sectionBounds);
            AddCapabilityVisual(slide, theme, options, visual);

            if (options.Metrics.Count > 0) {
                PowerPointLayoutBox metricBox = PowerPointLayoutBox.FromCentimeters(
                    narrative.LeftCm,
                    narrative.BottomCm - 1.05,
                    narrative.WidthCm,
                    1.05);
                PowerPointAutoShape band = slide.AddRectangleCm(metricBox.LeftCm, metricBox.TopCm,
                    metricBox.WidthCm, metricBox.HeightCm, "Capability Metric Band");
                band.FillColor = theme.AccentColor;
                band.OutlineColor = theme.AccentColor;
                AddMetrics(slide, theme, options.Metrics.ToList(), metricBox.LeftCm + 0.25, metricBox.TopCm + 0.12,
                    metricBox.WidthCm - 0.5, metricBox.HeightCm - 0.18);
            }
        }

        private static void AddCapabilityStacked(PowerPointSlide slide, PowerPointDesignTheme theme,
            IReadOnlyList<PowerPointCapabilitySection> sections, PowerPointCapabilitySlideOptions options,
            double slideWidthCm, double slideHeightCm) {
            double metricReserve = options.Metrics.Count > 0 ? 1.45 : 0;
            PowerPointLayoutBox bounds = PowerPointLayoutBox.FromCentimeters(
                1.5, 4.0, slideWidthCm - 3.0, slideHeightCm - 5.55 - metricReserve);
            AddCapabilitySections(slide, theme, sections, bounds, maxColumns: Math.Min(3, sections.Count));

            if (options.Metrics.Count > 0) {
                PowerPointLayoutBox metricBox = PowerPointLayoutBox.FromCentimeters(
                    1.5, slideHeightCm - 2.55, slideWidthCm - 3.0, 1.15);
                PowerPointAutoShape band = slide.AddRectangleCm(metricBox.LeftCm, metricBox.TopCm,
                    metricBox.WidthCm, metricBox.HeightCm, "Capability Metric Band");
                band.FillColor = theme.AccentColor;
                band.OutlineColor = theme.AccentColor;
                AddMetrics(slide, theme, options.Metrics.ToList(), metricBox.LeftCm + 0.35, metricBox.TopCm + 0.13,
                    metricBox.WidthCm - 0.7, metricBox.HeightCm - 0.2);
            }
        }

        private static void AddCapabilitySections(PowerPointSlide slide, PowerPointDesignTheme theme,
            IReadOnlyList<PowerPointCapabilitySection> sections, PowerPointLayoutBox bounds, int maxColumns = 1) {
            if (maxColumns > 1) {
                int columns = Math.Min(maxColumns, sections.Count);
                int rows = (int)Math.Ceiling(sections.Count / (double)columns);
                PowerPointLayoutBox[,] grid = bounds.SplitGridCm(rows, columns, 0.35, 0.45);
                for (int i = 0; i < sections.Count; i++) {
                    AddCapabilityPanel(slide, theme, sections[i], grid[i / columns, i % columns], i);
                }
                return;
            }

            PowerPointLayoutBox[] rowsBox = bounds.SplitRowsCm(sections.Count, 0.30);
            for (int i = 0; i < sections.Count; i++) {
                AddCapabilityPanel(slide, theme, sections[i], rowsBox[i], i);
            }
        }

        private static void AddCapabilityPanel(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointCapabilitySection section, PowerPointLayoutBox box, int index) {
            PowerPointAutoShape panel = slide.AddRectangleCm(box.LeftCm, box.TopCm, box.WidthCm, box.HeightCm,
                "Capability Section " + (index + 1));
            panel.FillColor = theme.PanelColor;
            panel.OutlineColor = theme.PanelBorderColor;
            panel.OutlineWidthPoints = 0.45;
            panel.SetShadow("000000", blurPoints: 2.5, distancePoints: 0.6, angleDegrees: 90, transparencyPercent: 92);

            string accent = section.AccentColor ?? GetAccent(theme, index);
            PowerPointAutoShape accentRule = slide.AddRectangleCm(box.LeftCm, box.TopCm, 0.12, box.HeightCm,
                "Capability Section Accent " + (index + 1));
            accentRule.FillColor = accent;
            accentRule.OutlineColor = accent;

            AddText(slide, section.Heading, box.LeftCm + 0.45, box.TopCm + 0.32, box.WidthCm - 0.75,
                0.45, 12, theme.PrimaryTextColor, theme.HeadingFontName, bold: true);

            double currentTop = box.TopCm + 0.88;
            if (!string.IsNullOrWhiteSpace(section.Body)) {
                PowerPointTextBox body = AddText(slide, section.Body!, box.LeftCm + 0.45, currentTop,
                    box.WidthCm - 0.75, Math.Min(0.75, box.HeightCm - 1.1), 9,
                    theme.SecondaryTextColor, theme.BodyFontName);
                body.TextAutoFitOptions = new PowerPointTextAutoFitOptions(fontScalePercent: 82,
                    lineSpaceReductionPercent: 18);
                currentTop += 0.78;
            }

            if (section.Items.Count > 0 && currentTop < box.BottomCm - 0.3) {
                PowerPointTextBox bullets = slide.AddTextBox("", PowerPointLayoutBox.FromCentimeters(
                    box.LeftCm + 0.52, currentTop, box.WidthCm - 0.85, box.BottomCm - currentTop - 0.25));
                bullets.SetTextMarginsCm(0, 0, 0, 0);
                bullets.TextAutoFit = PowerPointTextAutoFit.Normal;
                bullets.TextAutoFitOptions = new PowerPointTextAutoFitOptions(fontScalePercent: 78,
                    lineSpaceReductionPercent: 20);
                bullets.SetBullets(section.Items.Select(item => " " + item), configure: paragraph => {
                    paragraph.SetFontName(theme.BodyFontName)
                        .SetFontSize(9)
                        .SetColor(theme.SecondaryTextColor)
                        .SetHangingPoints(14)
                        .SetSpaceAfterPoints(3)
                        .SetBulletSizePercent(70);
                });
            }
        }

        private static void AddCapabilityVisual(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointCapabilitySlideOptions options, PowerPointLayoutBox bounds) {
            if (options.VisualKind == PowerPointCapabilityVisualKind.CoverageMap && options.Locations.Count > 0) {
                AddCoverageMap(slide, theme, options.Locations.ToList(), bounds,
                    new PowerPointCoverageSlideOptions {
                        MapLabel = options.VisualLabel,
                        DesignIntent = options.DesignIntent
                    });
                return;
            }

            if (options.VisualKind == PowerPointCapabilityVisualKind.LogoWall && options.Logos.Count > 0) {
                AddLogoWall(slide, theme, options.Logos.ToList(), new PowerPointLogoWallSlideOptions {
                    MaxColumns = Math.Min(3, Math.Max(1, options.Logos.Count)),
                    Variant = PowerPointLogoWallLayoutVariant.LogoMosaic,
                    DesignIntent = options.DesignIntent
                }, PowerPointLogoWallLayoutVariant.LogoMosaic, bounds);
                return;
            }

            AddVisualFrame(slide, theme, options.VisualImagePath, bounds.LeftCm, bounds.TopCm, bounds.WidthCm,
                bounds.HeightCm, options.DesignIntent);
            if (!string.IsNullOrWhiteSpace(options.VisualLabel)) {
                AddText(slide, options.VisualLabel!, bounds.LeftCm + 0.35, bounds.BottomCm - 0.55,
                    bounds.WidthCm - 0.7, 0.35, 9, theme.AccentContrastColor, theme.BodyFontName, bold: true);
            }
        }

        private static void AddCaseStudyVisualHero(PowerPointSlide slide, PowerPointDesignTheme theme,
            string clientTitle, IReadOnlyList<PowerPointCaseStudySection> sections, PowerPointCaseStudySlideOptions options,
            IReadOnlyList<PowerPointMetric> metrics, double slideWidthCm, double slideHeightCm) {
            AddText(slide, clientTitle, 1.45, 1.55, slideWidthCm * 0.62, 1.25, 27,
                theme.PrimaryTextColor, theme.HeadingFontName, bold: true);

            PowerPointAutoShape titleRule = slide.AddRectangleCm(1.48, 3.03, slideWidthCm * 0.20, 0.08,
                "Case Study Visual Hero Rule");
            titleRule.FillColor = theme.AccentColor;
            titleRule.OutlineColor = theme.AccentColor;

            double visualLeft = 1.45;
            double visualTop = 3.55;
            double visualWidth = slideWidthCm * 0.48;
            double visualHeight = slideHeightCm - 5.35;
            AddVisualFrame(slide, theme, options.VisualImagePath, visualLeft, visualTop, visualWidth, visualHeight,
                options.DesignIntent);

            PowerPointLayoutBox right = PowerPointLayoutBox.FromCentimeters(
                visualLeft + visualWidth + 0.85,
                visualTop,
                slideWidthCm - visualLeft - visualWidth - 2.3,
                visualHeight);

            double metricsHeight = metrics.Count > 0 ? 1.55 : 0;
            double sectionHeight = right.HeightCm - metricsHeight - (metrics.Count > 0 ? 0.55 : 0);
            int sectionCount = Math.Min(sections.Count, 4);
            PowerPointLayoutBox[] rows = PowerPointLayoutBox
                .FromCentimeters(right.LeftCm, right.TopCm, right.WidthCm, sectionHeight)
                .SplitRowsCm(sectionCount, 0.25);

            for (int i = 0; i < sectionCount; i++) {
                PowerPointLayoutBox row = rows[i];
                PowerPointCaseStudySection section = sections[i];
                PowerPointAutoShape accent = slide.AddRectangleCm(row.LeftCm, row.TopCm + 0.05,
                    0.12, row.HeightCm - 0.1, "Case Study Visual Hero Accent " + (i + 1));
                accent.FillColor = GetAccent(theme, i);
                accent.OutlineColor = GetAccent(theme, i);

                AddText(slide, section.Heading, row.LeftCm + 0.45, row.TopCm + 0.08,
                    row.WidthCm - 0.45, 0.42, 11, theme.PrimaryTextColor, theme.HeadingFontName, bold: true);
                PowerPointTextBox body = AddText(slide, section.Body, row.LeftCm + 0.45,
                    row.TopCm + 0.62, row.WidthCm - 0.45, row.HeightCm - 0.7, 9,
                    theme.SecondaryTextColor, theme.BodyFontName);
                body.TextAutoFitOptions = new PowerPointTextAutoFitOptions(fontScalePercent: 82,
                    lineSpaceReductionPercent: 18);
            }

            if (metrics.Count > 0) {
                PowerPointLayoutBox metricBox = right.TakeBottomCm(metricsHeight);
                PowerPointAutoShape metricBand = slide.AddRectangleCm(metricBox.LeftCm, metricBox.TopCm,
                    metricBox.WidthCm, metricBox.HeightCm, "Case Study Visual Hero Metric Band");
                metricBand.FillColor = theme.AccentColor;
                metricBand.OutlineColor = theme.AccentColor;
                metricBand.SetShadow("000000", blurPoints: 3, distancePoints: 0.8, angleDegrees: 90, transparencyPercent: 88);
                AddMetrics(slide, theme, metrics, metricBox.LeftCm + 0.35, metricBox.TopCm + 0.2,
                    metricBox.WidthCm - 0.7, metricBox.HeightCm - 0.3);
            }

            AddTags(slide, theme, options.Tags, 1.45, slideHeightCm - 1.85, visualWidth, 0.55);
        }

        private static void AddLogoMosaic(PowerPointSlide slide, PowerPointDesignTheme theme,
            IReadOnlyList<PowerPointLogoItem> logos, PowerPointLogoWallSlideOptions options,
            double slideWidthCm, double slideHeightCm) {
            double bottomReserve = string.IsNullOrWhiteSpace(options.SupportingText) ? 1.65 : 3.05;
            PowerPointLayoutBox bounds = PowerPointLayoutBox
                .FromCentimeters(1.5, 4.0, slideWidthCm - 3.0, slideHeightCm - 4.0 - bottomReserve);
            AddLogoGrid(slide, theme, logos, bounds, options);

            if (!string.IsNullOrWhiteSpace(options.SupportingText)) {
                PowerPointAutoShape band = slide.AddRectangleCm(1.55, slideHeightCm - 2.65,
                    slideWidthCm - 3.1, 1.2, "Logo Wall Supporting Band");
                band.FillColor = theme.PanelColor;
                band.OutlineColor = theme.PanelBorderColor;
                band.SetShadow("000000", blurPoints: 3, distancePoints: 0.8, angleDegrees: 90, transparencyPercent: 90);
                AddText(slide, options.SupportingText!, 2.0, slideHeightCm - 2.25, slideWidthCm - 4.0, 0.45,
                    12, theme.SecondaryTextColor, theme.BodyFontName, bold: true);
            }
        }

        private static void AddLogoCertificateFeature(PowerPointSlide slide, PowerPointDesignTheme theme,
            IReadOnlyList<PowerPointLogoItem> logos, PowerPointLogoWallSlideOptions options,
            double slideWidthCm, double slideHeightCm) {
            double height = string.IsNullOrWhiteSpace(options.SupportingText)
                ? slideHeightCm - 5.25
                : slideHeightCm - 7.6;
            PowerPointLayoutBox[] columns = PowerPointLayoutBox
                .FromCentimeters(1.5, 4.0, slideWidthCm - 3.0, height)
                .SplitColumnsCm(2, 0.9);
            AddLogoGrid(slide, theme, logos, columns[0], options);
            AddCertificatePanel(slide, theme, columns[1], options);

            if (!string.IsNullOrWhiteSpace(options.SupportingText)) {
                AddText(slide, options.SupportingText!, 1.55, slideHeightCm - 3.05,
                    slideWidthCm * 0.55, 0.5, 11, theme.SecondaryTextColor, theme.BodyFontName, bold: true);
            }
        }

        private static void AddLogoGrid(PowerPointSlide slide, PowerPointDesignTheme theme,
            IReadOnlyList<PowerPointLogoItem> logos, PowerPointLayoutBox bounds, PowerPointLogoWallSlideOptions options) {
            int maxColumns = Math.Max(1, options.MaxColumns);
            if (bounds.WidthCm < 12) {
                maxColumns = Math.Min(maxColumns, 4);
            }
            int columns = Math.Min(maxColumns, logos.Count);
            int rows = (int)Math.Ceiling(logos.Count / (double)columns);
            PowerPointLayoutBox[,] grid = bounds.SplitGridCm(rows, columns, 0.36, 0.36);

            for (int i = 0; i < logos.Count; i++) {
                int row = i / columns;
                int column = i % columns;
                AddLogoTile(slide, theme, logos[i], grid[row, column], i);
            }
        }

        private static void AddLogoTile(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointLogoItem logo, PowerPointLayoutBox box, int index) {
            string accent = logo.AccentColor ?? GetAccent(theme, index);
            PowerPointAutoShape tile = slide.AddRectangleCm(box.LeftCm, box.TopCm, box.WidthCm, box.HeightCm,
                "Logo Wall Tile " + (index + 1));
            tile.FillColor = theme.PanelColor;
            tile.OutlineColor = theme.PanelBorderColor;
            tile.OutlineWidthPoints = 0.45;
            tile.SetShadow("000000", blurPoints: 2.5, distancePoints: 0.6, angleDegrees: 90, transparencyPercent: 91);

            PowerPointAutoShape accentRule = slide.AddRectangleCm(box.LeftCm, box.TopCm, box.WidthCm, 0.08,
                "Logo Wall Accent " + (index + 1));
            accentRule.FillColor = accent;
            accentRule.OutlineColor = accent;

            if (!string.IsNullOrWhiteSpace(logo.ImagePath) && File.Exists(logo.ImagePath!)) {
                AddPictureIfExists(slide, logo.ImagePath!, box.LeftCm + 0.35, box.TopCm + 0.35,
                    box.WidthCm - 0.7, box.HeightCm - 0.85, crop: false);
            } else {
                int fontSize = box.WidthCm < 2.25 ? 12 : box.WidthCm < 2.7 ? 14 : 16;
                PowerPointTextBox name = AddText(slide, logo.Name, box.LeftCm + 0.25, box.TopCm + 0.45,
                    box.WidthCm - 0.5, box.HeightCm * 0.45, fontSize, theme.PrimaryTextColor,
                    theme.HeadingFontName, bold: true);
                name.TextAutoFitOptions = new PowerPointTextAutoFitOptions(fontScalePercent: 78, lineSpaceReductionPercent: 20);
                CenterText(name);
            }

            if (!string.IsNullOrWhiteSpace(logo.Subtitle)) {
                PowerPointTextBox subtitle = AddText(slide, logo.Subtitle!, box.LeftCm + 0.35,
                    box.TopCm + box.HeightCm - 0.55, box.WidthCm - 0.7, 0.3, 8,
                    theme.MutedTextColor, theme.BodyFontName);
                CenterText(subtitle);
            }
        }

        private static void AddCertificatePanel(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointLayoutBox bounds, PowerPointLogoWallSlideOptions options) {
            PowerPointAutoShape frame = slide.AddRectangleCm(bounds.LeftCm, bounds.TopCm, bounds.WidthCm,
                bounds.HeightCm, "Logo Wall Certificate Frame");
            frame.FillColor = theme.PanelColor;
            frame.OutlineColor = theme.PanelBorderColor;
            frame.OutlineWidthPoints = 0.45;
            frame.SetShadow("000000", blurPoints: 4, distancePoints: 1.0, angleDegrees: 90, transparencyPercent: 88);

            if (!string.IsNullOrWhiteSpace(options.FeaturedImagePath) && File.Exists(options.FeaturedImagePath!)) {
                AddPictureIfExists(slide, options.FeaturedImagePath!, bounds.LeftCm + 0.35, bounds.TopCm + 0.35,
                    bounds.WidthCm - 0.7, bounds.HeightCm - 0.95, crop: false);
            } else {
                PowerPointAutoShape rail = slide.AddRectangleCm(bounds.LeftCm + 0.46, bounds.TopCm + 0.46,
                    0.10, bounds.HeightCm - 0.92, "Logo Wall Certificate Rail");
                rail.FillColor = theme.AccentColor;
                rail.OutlineColor = theme.AccentColor;

                PowerPointAutoShape watermark = slide.AddShapeCm(A.ShapeTypeValues.Parallelogram,
                    bounds.LeftCm + bounds.WidthCm * 0.60, bounds.TopCm + 0.08,
                    bounds.WidthCm * 0.23, bounds.HeightCm - 0.16, "Logo Wall Certificate Watermark");
                watermark.FillColor = theme.AccentLightColor;
                watermark.FillTransparency = 66;
                watermark.OutlineColor = theme.AccentLightColor;
                watermark.OutlineWidthPoints = 0;

                PowerPointAutoShape document = slide.AddRectangleCm(bounds.LeftCm + bounds.WidthCm * 0.24,
                    bounds.TopCm + 0.55, bounds.WidthCm * 0.52, bounds.HeightCm - 1.45,
                    "Logo Wall Certificate Document");
                document.FillColor = theme.BackgroundColor;
                document.OutlineColor = theme.PanelBorderColor;
                document.OutlineWidthPoints = 0.55;
                document.SetShadow("000000", blurPoints: 2.5, distancePoints: 0.5, angleDegrees: 90, transparencyPercent: 92);

                PowerPointAutoShape documentHeader = slide.AddRectangleCm(bounds.LeftCm + bounds.WidthCm * 0.29,
                    bounds.TopCm + 0.92, bounds.WidthCm * 0.42, 0.08, "Logo Wall Certificate Header");
                documentHeader.FillColor = theme.AccentColor;
                documentHeader.OutlineColor = theme.AccentColor;

                PowerPointAutoShape seal = slide.AddEllipseCm(bounds.LeftCm + bounds.WidthCm * 0.42,
                    bounds.TopCm + bounds.HeightCm * 0.43, bounds.WidthCm * 0.16, bounds.WidthCm * 0.16,
                    "Logo Wall Certificate Seal");
                seal.FillColor = theme.AccentColor;
                seal.FillTransparency = 15;
                seal.OutlineColor = theme.AccentColor;
                seal.OutlineWidthPoints = 0.8;

                PowerPointAutoShape sealCenter = slide.AddEllipseCm(bounds.LeftCm + bounds.WidthCm * 0.455,
                    bounds.TopCm + bounds.HeightCm * 0.43 + bounds.WidthCm * 0.035,
                    bounds.WidthCm * 0.09, bounds.WidthCm * 0.09, "Logo Wall Certificate Seal Center");
                sealCenter.FillColor = theme.AccentContrastColor;
                sealCenter.FillTransparency = 42;
                sealCenter.OutlineColor = theme.AccentContrastColor;

                for (int i = 0; i < 5; i++) {
                    PowerPointAutoShape line = slide.AddRectangleCm(bounds.LeftCm + bounds.WidthCm * 0.32,
                        bounds.TopCm + 1.25 + i * 0.38, bounds.WidthCm * (i == 0 ? 0.34 : 0.25), 0.032,
                        "Logo Wall Certificate Line " + (i + 1));
                    line.FillColor = i == 0 ? theme.AccentColor : theme.AccentLightColor;
                    line.FillTransparency = i == 0 ? 0 : 18;
                    line.OutlineColor = line.FillColor;
                }

                for (int i = 0; i < 3; i++) {
                    PowerPointAutoShape signatureLine = slide.AddRectangleCm(bounds.LeftCm + bounds.WidthCm * (0.32 + i * 0.12),
                        bounds.TopCm + bounds.HeightCm - 1.18, bounds.WidthCm * 0.08, 0.025,
                        "Logo Wall Certificate Signature " + (i + 1));
                    signatureLine.FillColor = theme.PanelBorderColor;
                    signatureLine.OutlineColor = theme.PanelBorderColor;
                }
            }

            if (!string.IsNullOrWhiteSpace(options.FeatureTitle)) {
                AddText(slide, options.FeatureTitle!, bounds.LeftCm + 0.45, bounds.TopCm + bounds.HeightCm - 0.55,
                    bounds.WidthCm - 0.9, 0.35, 9, theme.SecondaryTextColor, theme.BodyFontName, bold: true);
            }
        }

        private static void AddCoverageList(PowerPointSlide slide, PowerPointDesignTheme theme,
            IReadOnlyList<PowerPointCoverageLocation> locations, PowerPointLayoutBox bounds, string? supportingText) {
            if (!string.IsNullOrWhiteSpace(supportingText)) {
                PowerPointAutoShape band = slide.AddRectangleCm(bounds.LeftCm, bounds.TopCm,
                    bounds.WidthCm, 1.1, "Coverage Supporting Band");
                band.FillColor = theme.PanelColor;
                band.OutlineColor = theme.PanelBorderColor;
                AddText(slide, supportingText!, bounds.LeftCm + 0.35, bounds.TopCm + 0.35,
                    bounds.WidthCm - 0.7, 0.4, 11, theme.SecondaryTextColor, theme.BodyFontName, bold: true);
                bounds = PowerPointLayoutBox.FromCentimeters(bounds.LeftCm, bounds.TopCm + 1.35,
                    bounds.WidthCm, bounds.HeightCm - 1.35);
            }

            int count = Math.Min(locations.Count, 8);
            PowerPointLayoutBox[] rows = bounds.SplitRowsCm(count, 0.16);
            for (int i = 0; i < count; i++) {
                PowerPointCoverageLocation location = locations[i];
                PowerPointLayoutBox row = rows[i];
                PowerPointAutoShape marker = slide.AddEllipseCm(row.LeftCm, row.TopCm + 0.08,
                    0.32, 0.32, "Coverage List Marker " + (i + 1));
                marker.FillColor = GetAccent(theme, i);
                marker.OutlineColor = GetAccent(theme, i);

                string text = string.IsNullOrWhiteSpace(location.Detail)
                    ? location.Name
                    : location.Name + " - " + location.Detail;
                AddText(slide, text, row.LeftCm + 0.55, row.TopCm, row.WidthCm - 0.55,
                    row.HeightCm, 10, theme.SecondaryTextColor, theme.BodyFontName, bold: i < 3);
            }
        }

        private static void AddCoverageStrip(PowerPointSlide slide, PowerPointDesignTheme theme,
            IReadOnlyList<PowerPointCoverageLocation> locations, PowerPointLayoutBox bounds) {
            PowerPointAutoShape band = slide.AddRectangleCm(bounds.LeftCm, bounds.TopCm, bounds.WidthCm,
                bounds.HeightCm, "Coverage Location Strip");
            band.FillColor = theme.PanelColor;
            band.OutlineColor = theme.PanelBorderColor;
            band.SetShadow("000000", blurPoints: 3, distancePoints: 0.7, angleDegrees: 90, transparencyPercent: 91);

            string text = string.Join(", ", locations.Take(12).Select(location => location.Name));
            AddText(slide, text, bounds.LeftCm + 0.55, bounds.TopCm + 0.32, bounds.WidthCm - 1.1,
                0.38, 10, theme.SecondaryTextColor, theme.BodyFontName, bold: true);
        }

        private static void AddCoverageRegion(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointLayoutBox bounds, double x, double y, double width, double height, string name,
            string fillColor, int fillTransparency) {
            PowerPointAutoShape region = slide.AddShapeCm(A.ShapeTypeValues.Parallelogram,
                bounds.LeftCm + bounds.WidthCm * x,
                bounds.TopCm + bounds.HeightCm * y,
                bounds.WidthCm * width,
                bounds.HeightCm * height,
                name);
            region.FillColor = fillColor;
            region.FillTransparency = fillTransparency;
            region.OutlineColor = theme.AccentLightColor;
            region.OutlineWidthPoints = 0.25;
        }

        private static void AddCoverageRoute(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointLayoutBox bounds, PowerPointCoverageLocation start, PowerPointCoverageLocation end, int index) {
            PowerPointAutoShape route = slide.AddLineCm(
                bounds.LeftCm + bounds.WidthCm * start.X,
                bounds.TopCm + bounds.HeightCm * start.Y,
                bounds.LeftCm + bounds.WidthCm * end.X,
                bounds.TopCm + bounds.HeightCm * end.Y,
                "Coverage Route " + (index + 1));
            route.OutlineColor = theme.AccentLightColor;
            route.OutlineWidthPoints = 0.45;
        }

        private static void AddCoveragePin(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointLayoutBox bounds, PowerPointCoverageLocation location, int index) {
            double x = bounds.LeftCm + bounds.WidthCm * location.X;
            double y = bounds.TopCm + bounds.HeightCm * location.Y;
            string accent = GetAccent(theme, index);

            PowerPointAutoShape halo = slide.AddEllipseCm(x - 0.22, y - 0.22, 0.44, 0.44,
                "Coverage Pin Halo " + (index + 1));
            halo.FillColor = theme.AccentContrastColor;
            halo.FillTransparency = 72;
            halo.OutlineColor = theme.AccentContrastColor;
            halo.OutlineWidthPoints = 0;

            PowerPointAutoShape pinOuter = slide.AddEllipseCm(x - 0.16, y - 0.16, 0.32, 0.32,
                "Coverage Pin Outer " + (index + 1));
            pinOuter.FillColor = theme.AccentDarkColor;
            pinOuter.FillTransparency = 35;
            pinOuter.OutlineColor = theme.AccentContrastColor;
            pinOuter.OutlineWidthPoints = 0.45;

            PowerPointAutoShape pin = slide.AddEllipseCm(x - 0.09, y - 0.09, 0.18, 0.18,
                "Coverage Pin " + (index + 1));
            pin.FillColor = accent;
            pin.OutlineColor = theme.AccentContrastColor;
            pin.OutlineWidthPoints = 0.25;
        }

        internal static List<PowerPointLogoItem> NormalizeLogoItems(IEnumerable<PowerPointLogoItem> logos) {
            if (logos == null) {
                throw new ArgumentNullException(nameof(logos));
            }

            List<PowerPointLogoItem> list = logos.Where(logo => logo != null).ToList();
            if (list.Count == 0) {
                throw new ArgumentException("At least one logo item is required.", nameof(logos));
            }
            if (list.Count > 24) {
                throw new ArgumentOutOfRangeException(nameof(logos), "This composition supports up to 24 logo items.");
            }

            return list;
        }

        internal static List<PowerPointCoverageLocation> NormalizeLocations(IEnumerable<PowerPointCoverageLocation> locations) {
            if (locations == null) {
                throw new ArgumentNullException(nameof(locations));
            }

            List<PowerPointCoverageLocation> list = locations.Where(location => location != null).ToList();
            if (list.Count == 0) {
                throw new ArgumentException("At least one location is required.", nameof(locations));
            }
            if (list.Count > 24) {
                throw new ArgumentOutOfRangeException(nameof(locations), "This composition supports up to 24 locations.");
            }

            foreach (PowerPointCoverageLocation location in list) {
                if (location.X < 0 || location.X > 1) {
                    throw new ArgumentOutOfRangeException(nameof(locations), "Location X must be between 0 and 1.");
                }
                if (location.Y < 0 || location.Y > 1) {
                    throw new ArgumentOutOfRangeException(nameof(locations), "Location Y must be between 0 and 1.");
                }
            }

            return list;
        }

        private static List<PowerPointCapabilitySection> NormalizeCapabilitySections(
            IEnumerable<PowerPointCapabilitySection> sections) {
            if (sections == null) {
                throw new ArgumentNullException(nameof(sections));
            }

            List<PowerPointCapabilitySection> list = sections.Where(section => section != null).ToList();
            if (list.Count == 0) {
                throw new ArgumentException("At least one capability section is required.", nameof(sections));
            }
            if (list.Count > 6) {
                throw new ArgumentOutOfRangeException(nameof(sections), "This composition supports up to 6 sections.");
            }

            return list;
        }

        internal static PowerPointLogoWallLayoutVariant ResolveLogoWallVariant(PowerPointLogoWallSlideOptions options,
            IReadOnlyList<PowerPointLogoItem> logos) {
            if (options.Variant != PowerPointLogoWallLayoutVariant.Auto) {
                return options.Variant;
            }

            bool hasFeaturedProof = !string.IsNullOrWhiteSpace(options.FeaturedImagePath) ||
                !string.IsNullOrWhiteSpace(options.FeatureTitle);
            if (hasFeaturedProof && logos.Count <= 12) {
                return PowerPointLogoWallLayoutVariant.CertificateFeature;
            }

            if (logos.Count > Math.Max(6, options.MaxColumns) ||
                options.DesignIntent.VisualStyle == PowerPointVisualStyle.Minimal) {
                return PowerPointLogoWallLayoutVariant.LogoMosaic;
            }

            if (string.IsNullOrWhiteSpace(options.DesignIntent.Seed)) {
                return PowerPointLogoWallLayoutVariant.LogoMosaic;
            }

            return options.DesignIntent.Pick(2, "logo-wall") == 0
                ? PowerPointLogoWallLayoutVariant.LogoMosaic
                : PowerPointLogoWallLayoutVariant.CertificateFeature;
        }

        internal static PowerPointCoverageLayoutVariant ResolveCoverageVariant(PowerPointCoverageSlideOptions options,
            IReadOnlyList<PowerPointCoverageLocation> locations) {
            if (options.Variant != PowerPointCoverageLayoutVariant.Auto) {
                return options.Variant;
            }

            if (locations.Count > 6 ||
                !string.IsNullOrWhiteSpace(options.SupportingText) ||
                options.DesignIntent.VisualStyle == PowerPointVisualStyle.Soft ||
                options.DesignIntent.VisualStyle == PowerPointVisualStyle.Minimal) {
                return PowerPointCoverageLayoutVariant.ListMap;
            }

            if (string.IsNullOrWhiteSpace(options.DesignIntent.Seed)) {
                return PowerPointCoverageLayoutVariant.PinBoard;
            }

            return options.DesignIntent.Pick(2, "coverage") == 0
                ? PowerPointCoverageLayoutVariant.PinBoard
                : PowerPointCoverageLayoutVariant.ListMap;
        }

        internal static PowerPointCapabilityLayoutVariant ResolveCapabilityVariant(PowerPointCapabilitySlideOptions options,
            IReadOnlyList<PowerPointCapabilitySection> sections) {
            if (options.Variant != PowerPointCapabilityLayoutVariant.Auto) {
                return options.Variant;
            }

            bool hasCoverageVisual = options.VisualKind == PowerPointCapabilityVisualKind.CoverageMap &&
                options.Locations.Count > 0;
            bool hasLogoVisual = options.VisualKind == PowerPointCapabilityVisualKind.LogoWall &&
                options.Logos.Count > 0;
            bool hasImageVisual = !string.IsNullOrWhiteSpace(options.VisualImagePath);
            bool hasVisualEvidence = hasCoverageVisual || hasLogoVisual || hasImageVisual;

            if (sections.Count >= 4) {
                return PowerPointCapabilityLayoutVariant.Stacked;
            }

            if (options.DesignIntent.VisualStyle == PowerPointVisualStyle.Minimal && sections.Count >= 3) {
                return PowerPointCapabilityLayoutVariant.Stacked;
            }

            if (hasVisualEvidence && sections.Count <= 3) {
                return options.DesignIntent.Mood == PowerPointDesignMood.Editorial
                    ? PowerPointCapabilityLayoutVariant.VisualText
                    : PowerPointCapabilityLayoutVariant.TextVisual;
            }

            if (string.IsNullOrWhiteSpace(options.DesignIntent.Seed)) {
                return PowerPointCapabilityLayoutVariant.TextVisual;
            }

            return options.DesignIntent.Pick(3, "capability") switch {
                0 => PowerPointCapabilityLayoutVariant.TextVisual,
                1 => PowerPointCapabilityLayoutVariant.VisualText,
                _ => PowerPointCapabilityLayoutVariant.Stacked
            };
        }
    }
}
