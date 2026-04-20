using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     High-level slide composition helpers for building polished decks without hand-placing every shape.
    /// </summary>
    public static partial class PowerPointDesignExtensions {
        /// <summary>
        ///     Applies the designer theme colors and fonts to all slide masters.
        /// </summary>
        public static PowerPointPresentation ApplyDesignerTheme(this PowerPointPresentation presentation,
            PowerPointDesignTheme? theme = null) {
            if (presentation == null) {
                throw new ArgumentNullException(nameof(presentation));
            }

            PowerPointDesignTheme resolvedTheme = ResolveTheme(theme);
            presentation.ThemeName = resolvedTheme.Name;
            presentation.SetThemeColorsForAllMasters(new Dictionary<PowerPointThemeColor, string> {
                [PowerPointThemeColor.Dark1] = resolvedTheme.PrimaryTextColor,
                [PowerPointThemeColor.Light1] = resolvedTheme.BackgroundColor,
                [PowerPointThemeColor.Dark2] = resolvedTheme.AccentDarkColor,
                [PowerPointThemeColor.Light2] = resolvedTheme.SurfaceColor,
                [PowerPointThemeColor.Accent1] = resolvedTheme.AccentColor,
                [PowerPointThemeColor.Accent2] = resolvedTheme.Accent2Color,
                [PowerPointThemeColor.Accent3] = resolvedTheme.Accent3Color,
                [PowerPointThemeColor.Accent4] = resolvedTheme.WarningColor,
                [PowerPointThemeColor.Accent5] = resolvedTheme.PanelBorderColor,
                [PowerPointThemeColor.Accent6] = resolvedTheme.MutedTextColor
            });
            presentation.SetThemeFontsForAllMasters(new PowerPointThemeFontSet(
                resolvedTheme.HeadingFontName,
                resolvedTheme.BodyFontName,
                null,
                null,
                null,
                null));

            return presentation;
        }

        /// <summary>
        ///     Adds a full-bleed section/title slide with diagonal planes and optional footer chrome.
        /// </summary>
        public static PowerPointSlide AddDesignerSectionSlide(this PowerPointPresentation presentation, string title,
            string? subtitle = null, PowerPointDesignTheme? theme = null,
            PowerPointDesignerSlideOptions? options = null) {
            if (presentation == null) {
                throw new ArgumentNullException(nameof(presentation));
            }
            if (title == null) {
                throw new ArgumentNullException(nameof(title));
            }

            PowerPointDesignTheme resolvedTheme = ResolveTheme(theme);
            PowerPointDesignerSlideOptions resolvedOptions = options ?? new PowerPointDesignerSlideOptions();
            PowerPointSlide slide = presentation.AddSlide();
            double width = presentation.SlideSize.WidthCm;
            double height = presentation.SlideSize.HeightCm;
            PowerPointSectionLayoutVariant variant = ResolveSectionVariant(resolvedOptions);

            if (variant == PowerPointSectionLayoutVariant.EditorialRail) {
                AddSectionEditorialRail(slide, resolvedTheme, resolvedOptions, title, subtitle, width, height);
            } else if (variant == PowerPointSectionLayoutVariant.Poster) {
                AddSectionPoster(slide, resolvedTheme, resolvedOptions, title, subtitle, width, height);
            } else {
                AddSectionGeometricCover(slide, resolvedTheme, resolvedOptions, title, subtitle, width, height);
            }

            return slide;
        }

        /// <summary>
        ///     Adds a case-study slide with summary columns, a strong visual band, metrics, and optional tags.
        /// </summary>
        public static PowerPointSlide AddDesignerCaseStudySlide(this PowerPointPresentation presentation,
            string clientTitle, IEnumerable<PowerPointCaseStudySection> sections,
            IEnumerable<PowerPointMetric>? metrics = null,
            PowerPointDesignTheme? theme = null,
            PowerPointCaseStudySlideOptions? options = null) {
            if (presentation == null) {
                throw new ArgumentNullException(nameof(presentation));
            }
            if (clientTitle == null) {
                throw new ArgumentNullException(nameof(clientTitle));
            }

            PowerPointDesignTheme resolvedTheme = ResolveTheme(theme);
            PowerPointCaseStudySlideOptions resolvedOptions = options ?? new PowerPointCaseStudySlideOptions();
            List<PowerPointCaseStudySection> sectionList = NormalizeSections(sections, 4, nameof(sections));
            List<PowerPointMetric> metricList = (metrics ?? Enumerable.Empty<PowerPointMetric>()).Where(m => m != null).ToList();

            PowerPointSlide slide = presentation.AddSlide();
            double width = presentation.SlideSize.WidthCm;
            double height = presentation.SlideSize.HeightCm;
            slide.BackgroundColor = resolvedTheme.BackgroundColor;

            PowerPointCaseStudyLayoutVariant variant = ResolveCaseStudyVariant(resolvedOptions, sectionList, metricList);
            if (variant == PowerPointCaseStudyLayoutVariant.EditorialSplit) {
                AddSubtleLightBackground(slide, resolvedTheme, width, height);
                AddChrome(slide, resolvedTheme, width, height, dark: false, resolvedOptions);
                AddCaseStudyEditorialSplit(slide, resolvedTheme, clientTitle, sectionList, resolvedOptions, metricList, width, height);
            } else if (variant == PowerPointCaseStudyLayoutVariant.VisualHero) {
                AddSubtleLightBackground(slide, resolvedTheme, width, height);
                AddChrome(slide, resolvedTheme, width, height, dark: false, resolvedOptions);
                AddCaseStudyVisualHero(slide, resolvedTheme, clientTitle, sectionList, resolvedOptions, metricList, width, height);
            } else {
                AddChrome(slide, resolvedTheme, width, height, dark: false, resolvedOptions);
                AddCaseStudyColumns(slide, resolvedTheme, clientTitle, sectionList, width);
                AddCaseStudyBand(slide, resolvedTheme, resolvedOptions, metricList, width, height);
            }

            return slide;
        }

        /// <summary>
        ///     Adds a card grid slide that automatically chooses rows and columns for the supplied content.
        /// </summary>
        public static PowerPointSlide AddDesignerCardGridSlide(this PowerPointPresentation presentation,
            string title, string? subtitle, IEnumerable<PowerPointCardContent> cards,
            PowerPointDesignTheme? theme = null,
            PowerPointCardGridSlideOptions? options = null) {
            if (presentation == null) {
                throw new ArgumentNullException(nameof(presentation));
            }
            if (title == null) {
                throw new ArgumentNullException(nameof(title));
            }

            PowerPointDesignTheme resolvedTheme = ResolveTheme(theme);
            PowerPointCardGridSlideOptions resolvedOptions = options ?? new PowerPointCardGridSlideOptions();
            List<PowerPointCardContent> cardList = NormalizeCards(cards);

            PowerPointSlide slide = presentation.AddSlide();
            double width = presentation.SlideSize.WidthCm;
            double height = presentation.SlideSize.HeightCm;
            slide.BackgroundColor = resolvedTheme.BackgroundColor;

            AddSubtleLightBackground(slide, resolvedTheme, width, height);
            AddChrome(slide, resolvedTheme, width, height, dark: false, resolvedOptions);
            AddText(slide, title, 1.5, 1.45, width * 0.6, 1.0, 29,
                resolvedTheme.PrimaryTextColor, resolvedTheme.HeadingFontName, bold: true);
            if (!string.IsNullOrWhiteSpace(subtitle)) {
                AddText(slide, subtitle!, 1.55, 2.7, width * 0.58, 0.5, 12,
                    resolvedTheme.SecondaryTextColor, resolvedTheme.BodyFontName, bold: true);
            }

            AddCardGrid(slide, resolvedTheme, cardList, resolvedOptions,
                ResolveCardGridVariant(resolvedOptions, cardList), width, height);

            return slide;
        }

        /// <summary>
        ///     Adds a dark process slide with a readable timeline and automatic spacing.
        /// </summary>
        public static PowerPointSlide AddDesignerProcessSlide(this PowerPointPresentation presentation,
            string title, string? subtitle, IEnumerable<PowerPointProcessStep> steps,
            PowerPointDesignTheme? theme = null,
            PowerPointProcessSlideOptions? options = null) {
            if (presentation == null) {
                throw new ArgumentNullException(nameof(presentation));
            }
            if (title == null) {
                throw new ArgumentNullException(nameof(title));
            }

            PowerPointDesignTheme resolvedTheme = ResolveTheme(theme);
            PowerPointProcessSlideOptions resolvedOptions = options ?? new PowerPointProcessSlideOptions();
            List<PowerPointProcessStep> stepList = NormalizeSteps(steps);

            PowerPointSlide slide = presentation.AddSlide();
            double width = presentation.SlideSize.WidthCm;
            double height = presentation.SlideSize.HeightCm;
            slide.BackgroundColor = resolvedTheme.AccentDarkColor;

            if (resolvedOptions.ShowDiagonalPlanes) {
                AddDiagonalPlanes(slide, resolvedTheme, width, height, dark: true);
            }
            AddChrome(slide, resolvedTheme, width, height, dark: true, resolvedOptions);
            AddText(slide, title, 1.85, 1.45, width * 0.52, 1.1, 33,
                resolvedTheme.AccentContrastColor, resolvedTheme.HeadingFontName, bold: true);
            if (!string.IsNullOrWhiteSpace(subtitle)) {
                AddText(slide, subtitle!, 1.9, 2.78, width * 0.58, 0.5, 13,
                    resolvedTheme.AccentLightColor, resolvedTheme.BodyFontName, bold: true);
            }

            AddProcessTimeline(slide, resolvedTheme, stepList, resolvedOptions, width, height);
            return slide;
        }

        private static PowerPointDesignTheme ResolveTheme(PowerPointDesignTheme? theme) {
            PowerPointDesignTheme resolved = theme ?? PowerPointDesignTheme.ModernBlue;
            resolved.Validate();
            return resolved;
        }

        private static List<PowerPointCaseStudySection> NormalizeSections(IEnumerable<PowerPointCaseStudySection> sections,
            int maxCount, string paramName) {
            if (sections == null) {
                throw new ArgumentNullException(paramName);
            }

            List<PowerPointCaseStudySection> list = sections.Where(section => section != null).ToList();
            if (list.Count == 0) {
                throw new ArgumentException("At least one section is required.", paramName);
            }
            if (list.Count > maxCount) {
                throw new ArgumentOutOfRangeException(paramName, $"This composition supports up to {maxCount} sections.");
            }

            return list;
        }

        internal static List<PowerPointCardContent> NormalizeCards(IEnumerable<PowerPointCardContent> cards) {
            if (cards == null) {
                throw new ArgumentNullException(nameof(cards));
            }

            List<PowerPointCardContent> list = cards.Where(card => card != null).ToList();
            if (list.Count == 0) {
                throw new ArgumentException("At least one card is required.", nameof(cards));
            }

            return list;
        }

        internal static List<PowerPointProcessStep> NormalizeSteps(IEnumerable<PowerPointProcessStep> steps) {
            if (steps == null) {
                throw new ArgumentNullException(nameof(steps));
            }

            List<PowerPointProcessStep> list = steps.Where(step => step != null).ToList();
            if (list.Count == 0) {
                throw new ArgumentException("At least one step is required.", nameof(steps));
            }
            if (list.Count > 8) {
                throw new ArgumentOutOfRangeException(nameof(steps), "This composition supports up to 8 steps.");
            }

            return list;
        }

        internal static PowerPointSectionLayoutVariant ResolveSectionVariant(PowerPointDesignerSlideOptions options) {
            if (options.SectionVariant != PowerPointSectionLayoutVariant.Auto) {
                return options.SectionVariant;
            }

            if (options.DesignIntent.LayoutStrategy == PowerPointAutoLayoutStrategy.Compact) {
                return PowerPointSectionLayoutVariant.EditorialRail;
            }
            if (options.DesignIntent.LayoutStrategy == PowerPointAutoLayoutStrategy.VisualFirst) {
                return PowerPointSectionLayoutVariant.Poster;
            }
            if (string.IsNullOrWhiteSpace(options.DesignIntent.Seed)) {
                return PowerPointSectionLayoutVariant.GeometricCover;
            }

            return options.DesignIntent.Pick(3, "section") switch {
                0 => PowerPointSectionLayoutVariant.GeometricCover,
                1 => PowerPointSectionLayoutVariant.EditorialRail,
                _ => PowerPointSectionLayoutVariant.Poster
            };
        }

        internal static PowerPointTitleAccentStyle ResolveTitleAccentStyle(PowerPointDesignerSlideOptions options,
            PowerPointSectionLayoutVariant variant) {
            if (options.TitleAccentStyle != PowerPointTitleAccentStyle.Auto) {
                return options.TitleAccentStyle;
            }

            PowerPointDesignIntent intent = options.DesignIntent;
            if (string.IsNullOrWhiteSpace(intent.Seed) ||
                intent.VisualStyle == PowerPointVisualStyle.Minimal) {
                return PowerPointTitleAccentStyle.None;
            }
            if (intent.Mood == PowerPointDesignMood.Editorial) {
                return PowerPointTitleAccentStyle.KickerRule;
            }
            if (intent.Mood == PowerPointDesignMood.Energetic ||
                intent.LayoutStrategy == PowerPointAutoLayoutStrategy.VisualFirst ||
                variant == PowerPointSectionLayoutVariant.Poster) {
                return PowerPointTitleAccentStyle.Underline;
            }
            if (intent.VisualStyle == PowerPointVisualStyle.Soft) {
                return PowerPointTitleAccentStyle.SideRule;
            }

            return intent.Pick(3, "title-accent") switch {
                0 => PowerPointTitleAccentStyle.Underline,
                1 => PowerPointTitleAccentStyle.SideRule,
                _ => PowerPointTitleAccentStyle.KickerRule
            };
        }

        internal static PowerPointCaseStudyLayoutVariant ResolveCaseStudyVariant(PowerPointCaseStudySlideOptions options,
            IReadOnlyList<PowerPointCaseStudySection> sections, IReadOnlyList<PowerPointMetric> metrics) {
            if (options.Variant != PowerPointCaseStudyLayoutVariant.Auto) {
                return options.Variant;
            }

            if (options.DesignIntent.LayoutStrategy == PowerPointAutoLayoutStrategy.Compact ||
                sections.Count >= 4) {
                return PowerPointCaseStudyLayoutVariant.EditorialSplit;
            }
            if (options.DesignIntent.LayoutStrategy == PowerPointAutoLayoutStrategy.VisualFirst &&
                sections.Count <= 3) {
                return PowerPointCaseStudyLayoutVariant.VisualHero;
            }
            if (string.IsNullOrWhiteSpace(options.DesignIntent.Seed)) {
                return PowerPointCaseStudyLayoutVariant.VisualBand;
            }
            if (options.DesignIntent.LayoutStrategy == PowerPointAutoLayoutStrategy.DesignFirst) {
                return options.DesignIntent.Pick(3, "case-study") switch {
                    0 => PowerPointCaseStudyLayoutVariant.VisualBand,
                    1 => PowerPointCaseStudyLayoutVariant.EditorialSplit,
                    _ => PowerPointCaseStudyLayoutVariant.VisualHero
                };
            }
            if (options.DesignIntent.VisualStyle == PowerPointVisualStyle.Soft ||
                options.DesignIntent.VisualStyle == PowerPointVisualStyle.Minimal) {
                return PowerPointCaseStudyLayoutVariant.EditorialSplit;
            }
            if (metrics.Count > 0 && sections.Count <= 3) {
                return PowerPointCaseStudyLayoutVariant.VisualHero;
            }

            return options.DesignIntent.Pick(3, "case-study") switch {
                0 => PowerPointCaseStudyLayoutVariant.VisualBand,
                1 => PowerPointCaseStudyLayoutVariant.EditorialSplit,
                _ => PowerPointCaseStudyLayoutVariant.VisualHero
            };
        }

        internal static PowerPointCardGridLayoutVariant ResolveCardGridVariant(PowerPointCardGridSlideOptions options,
            IReadOnlyList<PowerPointCardContent> cards) {
            if (options.Variant != PowerPointCardGridLayoutVariant.Auto) {
                return options.Variant;
            }

            if (options.DesignIntent.LayoutStrategy == PowerPointAutoLayoutStrategy.Compact ||
                cards.Count > 4) {
                return PowerPointCardGridLayoutVariant.AccentTop;
            }
            if (string.IsNullOrWhiteSpace(options.DesignIntent.Seed)) {
                return PowerPointCardGridLayoutVariant.AccentTop;
            }
            if (options.DesignIntent.LayoutStrategy == PowerPointAutoLayoutStrategy.DesignFirst ||
                options.DesignIntent.LayoutStrategy == PowerPointAutoLayoutStrategy.VisualFirst) {
                return options.DesignIntent.Pick(2, "card-grid") == 0
                    ? PowerPointCardGridLayoutVariant.AccentTop
                    : PowerPointCardGridLayoutVariant.SoftTiles;
            }
            if (options.DesignIntent.VisualStyle == PowerPointVisualStyle.Soft ||
                options.DesignIntent.VisualStyle == PowerPointVisualStyle.Minimal) {
                return PowerPointCardGridLayoutVariant.SoftTiles;
            }
            if (options.DesignIntent.Density == PowerPointSlideDensity.Compact) {
                return PowerPointCardGridLayoutVariant.AccentTop;
            }

            return options.DesignIntent.Pick(2, "card-grid") == 0
                ? PowerPointCardGridLayoutVariant.AccentTop
                : PowerPointCardGridLayoutVariant.SoftTiles;
        }

        internal static PowerPointCardSurfaceStyle ResolveCardSurfaceStyle(PowerPointCardGridSlideOptions options,
            PowerPointCardGridLayoutVariant variant) {
            if (options.SurfaceStyle != PowerPointCardSurfaceStyle.Auto) {
                return options.SurfaceStyle;
            }

            PowerPointDesignIntent intent = options.DesignIntent;
            if (string.IsNullOrWhiteSpace(intent.Seed)) {
                return PowerPointCardSurfaceStyle.Elevated;
            }
            if (intent.VisualStyle == PowerPointVisualStyle.Minimal) {
                return PowerPointCardSurfaceStyle.Flat;
            }
            if (intent.Mood == PowerPointDesignMood.Editorial) {
                return PowerPointCardSurfaceStyle.Hairline;
            }
            if (intent.Mood == PowerPointDesignMood.Energetic ||
                intent.LayoutStrategy == PowerPointAutoLayoutStrategy.DesignFirst) {
                return PowerPointCardSurfaceStyle.AccentWash;
            }
            if (variant == PowerPointCardGridLayoutVariant.SoftTiles) {
                return PowerPointCardSurfaceStyle.Flat;
            }

            return intent.Pick(4, "card-surface") switch {
                0 => PowerPointCardSurfaceStyle.Elevated,
                1 => PowerPointCardSurfaceStyle.Flat,
                2 => PowerPointCardSurfaceStyle.Hairline,
                _ => PowerPointCardSurfaceStyle.AccentWash
            };
        }

        internal static PowerPointProcessLayoutVariant ResolveProcessVariant(PowerPointProcessSlideOptions options,
            IReadOnlyList<PowerPointProcessStep> steps) {
            if (options.Variant != PowerPointProcessLayoutVariant.Auto) {
                return options.Variant;
            }

            if (steps.Count > PowerPointDeckPlanLimits.DenseProcessSteps ||
                options.DesignIntent.VisualStyle == PowerPointVisualStyle.Minimal) {
                return PowerPointProcessLayoutVariant.Rail;
            }
            if (options.DesignIntent.LayoutStrategy == PowerPointAutoLayoutStrategy.Compact) {
                return PowerPointProcessLayoutVariant.NumberedColumns;
            }
            if (string.IsNullOrWhiteSpace(options.DesignIntent.Seed)) {
                return PowerPointProcessLayoutVariant.Rail;
            }
            if (options.DesignIntent.LayoutStrategy == PowerPointAutoLayoutStrategy.DesignFirst ||
                options.DesignIntent.LayoutStrategy == PowerPointAutoLayoutStrategy.VisualFirst ||
                options.DesignIntent.Density != PowerPointSlideDensity.Compact) {
                return options.DesignIntent.Pick(2, "process") == 0
                    ? PowerPointProcessLayoutVariant.Rail
                    : PowerPointProcessLayoutVariant.NumberedColumns;
            }

            return PowerPointProcessLayoutVariant.NumberedColumns;
        }

        internal static PowerPointProcessConnectorStyle ResolveProcessConnectorStyle(PowerPointProcessSlideOptions options,
            IReadOnlyList<PowerPointProcessStep> steps) {
            if (options.ConnectorStyle != PowerPointProcessConnectorStyle.Auto) {
                return options.ConnectorStyle;
            }

            if (options.DesignIntent.VisualStyle == PowerPointVisualStyle.Minimal) {
                return PowerPointProcessConnectorStyle.None;
            }

            if (options.DesignIntent.Density == PowerPointSlideDensity.Compact || steps.Count > 5) {
                return PowerPointProcessConnectorStyle.ContinuousRail;
            }

            if (options.DesignIntent.Mood == PowerPointDesignMood.Energetic) {
                return PowerPointProcessConnectorStyle.SegmentArrows;
            }

            if (options.DesignIntent.Mood == PowerPointDesignMood.Editorial ||
                options.DesignIntent.VisualStyle == PowerPointVisualStyle.Soft) {
                return PowerPointProcessConnectorStyle.StepDots;
            }

            return PowerPointProcessConnectorStyle.ContinuousRail;
        }

        private static void AddSectionGeometricCover(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointDesignerSlideOptions options, string title, string? subtitle, double slideWidthCm,
            double slideHeightCm) {
            slide.BackgroundColor = theme.AccentDarkColor;
            AddDiagonalPlanes(slide, theme, slideWidthCm, slideHeightCm, dark: true);
            AddChrome(slide, theme, slideWidthCm, slideHeightCm, dark: true, options);

            PowerPointTitleAccentStyle titleAccent = ResolveTitleAccentStyle(options,
                PowerPointSectionLayoutVariant.GeometricCover);
            double titleLeft = 1.85;
            double titleTop = slideHeightCm * 0.47;
            double titleWidth = slideWidthCm * 0.58;
            AddText(slide, title, titleLeft, titleTop, titleWidth, 1.35, 40,
                theme.AccentContrastColor, theme.HeadingFontName, bold: true);
            AddSectionTitleAccent(slide, theme, titleAccent, titleLeft, titleTop, titleWidth, 1.35, dark: true);

            if (!string.IsNullOrWhiteSpace(subtitle)) {
                AddText(slide, subtitle!, 1.9, slideHeightCm * 0.59, slideWidthCm * 0.52, 0.8, 15,
                    theme.AccentLightColor, theme.BodyFontName);
            }

            if (ShouldShowDirectionMotif(options)) {
                AddDirectionMotif(slide, options, 1.95, slideHeightCm * 0.67, 11, 0.46, theme.WarningColor);
            }
        }

        private static void AddSectionEditorialRail(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointDesignerSlideOptions options, string title, string? subtitle, double slideWidthCm,
            double slideHeightCm) {
            slide.BackgroundColor = theme.BackgroundColor;
            AddSubtleLightBackground(slide, theme, slideWidthCm, slideHeightCm);
            AddChrome(slide, theme, slideWidthCm, slideHeightCm, dark: false, options);

            PowerPointAutoShape rail = slide.AddRectangleCm(1.45, 1.85, 0.18, slideHeightCm - 3.8,
                "Section Editorial Rail");
            rail.FillColor = theme.AccentColor;
            rail.OutlineColor = theme.AccentColor;

            PowerPointAutoShape block = slide.AddRectangleCm(1.9, 3.65, slideWidthCm * 0.42, 0.18,
                "Section Editorial Rule");
            block.FillColor = theme.WarningColor;
            block.OutlineColor = theme.WarningColor;

            PowerPointTitleAccentStyle titleAccent = ResolveTitleAccentStyle(options,
                PowerPointSectionLayoutVariant.EditorialRail);
            double titleLeft = 1.9;
            double titleTop = 2.15;
            double titleWidth = slideWidthCm * 0.55;
            AddText(slide, title, titleLeft, titleTop, titleWidth, 1.2, 38,
                theme.PrimaryTextColor, theme.HeadingFontName, bold: true);
            AddSectionTitleAccent(slide, theme, titleAccent, titleLeft, titleTop, titleWidth, 1.2, dark: false);

            if (!string.IsNullOrWhiteSpace(subtitle)) {
                AddText(slide, subtitle!, 1.95, 4.15, slideWidthCm * 0.47, 0.8, 14,
                    theme.SecondaryTextColor, theme.BodyFontName);
            }

            PowerPointAutoShape accentPanel = slide.AddShapeCm(A.ShapeTypeValues.Parallelogram,
                slideWidthCm * 0.64, 0, slideWidthCm * 0.26, slideHeightCm, "Section Editorial Accent Plane");
            accentPanel.FillColor = theme.AccentLightColor;
            accentPanel.FillTransparency = 30;
            accentPanel.OutlineColor = theme.AccentLightColor;

            if (ShouldShowDirectionMotif(options)) {
                AddDirectionMotif(slide, options, slideWidthCm - 5.25, 2.05, 10, 0.36, theme.AccentColor,
                    flip: true);
            }
        }

        private static void AddSectionPoster(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointDesignerSlideOptions options, string title, string? subtitle, double slideWidthCm,
            double slideHeightCm) {
            slide.BackgroundColor = theme.AccentDarkColor;
            AddChrome(slide, theme, slideWidthCm, slideHeightCm, dark: true, options);

            PowerPointAutoShape wash = slide.AddRectangleCm(0, 0, slideWidthCm, slideHeightCm,
                "Section Poster Wash");
            wash.FillColor = theme.AccentColor;
            wash.FillTransparency = 50;
            wash.OutlineColor = theme.AccentColor;
            wash.SendToBack();

            PowerPointAutoShape frame = slide.AddRectangleCm(1.45, 1.55, slideWidthCm - 2.9, slideHeightCm - 3.2,
                "Section Poster Frame");
            frame.FillColor = theme.AccentDarkColor;
            frame.FillTransparency = 100;
            frame.OutlineColor = theme.AccentLightColor;
            frame.OutlineWidthPoints = 0.7;

            PowerPointTitleAccentStyle titleAccent = ResolveTitleAccentStyle(options,
                PowerPointSectionLayoutVariant.Poster);
            PowerPointTextBox titleBox = AddText(slide, title, 2.4, slideHeightCm * 0.42, slideWidthCm - 4.8, 1.4,
                42, theme.AccentContrastColor, theme.HeadingFontName, bold: true);
            CenterText(titleBox);
            AddSectionTitleAccent(slide, theme, titleAccent, 2.4, slideHeightCm * 0.42, slideWidthCm - 4.8, 1.4,
                dark: true, centered: true);

            if (!string.IsNullOrWhiteSpace(subtitle)) {
                PowerPointTextBox subtitleBox = AddText(slide, subtitle!, 4.1, slideHeightCm * 0.58,
                    slideWidthCm - 8.2, 0.65, 14, theme.AccentLightColor, theme.BodyFontName);
                CenterText(subtitleBox);
            }

            if (ShouldShowDirectionMotif(options)) {
                AddDirectionMotif(slide, options, slideWidthCm * 0.39, slideHeightCm * 0.68, 12, 0.4,
                    theme.WarningColor);
            }
        }

        private static void AddCaseStudyColumns(PowerPointSlide slide, PowerPointDesignTheme theme, string clientTitle,
            IReadOnlyList<PowerPointCaseStudySection> sections, double slideWidthCm) {
            double left = 1.45;
            double top = 1.75;
            double gutter = 0.85;
            double width = slideWidthCm - 2.9;
            PowerPointLayoutBox[] columns = PowerPointLayoutBox
                .FromCentimeters(left, top, width, 6.15)
                .SplitColumnsCm(sections.Count, gutter);

            for (int i = 0; i < sections.Count; i++) {
                PowerPointLayoutBox box = columns[i];
                PowerPointCaseStudySection section = sections[i];
                AddText(slide, section.Heading, box.LeftCm, box.TopCm, box.WidthCm, 0.55, i == 0 ? 19 : 11,
                    theme.PrimaryTextColor, theme.HeadingFontName, bold: true);

                if (i == 0) {
                    PowerPointAutoShape rule = slide.AddRectangleCm(box.LeftCm, box.TopCm + 0.78, box.WidthCm, 0.025,
                        "Case Study Client Rule");
                    rule.FillColor = theme.PanelBorderColor;
                    rule.OutlineColor = theme.PanelBorderColor;
                }

                string body = i == 0 ? clientTitle + Environment.NewLine + section.Body : section.Body;
                PowerPointTextBox bodyBox = AddText(slide, body, box.LeftCm, box.TopCm + 1.05, box.WidthCm, 4.5,
                    i == 0 ? 13 : 10, i == 0 ? theme.PrimaryTextColor : theme.SecondaryTextColor,
                    theme.BodyFontName, bold: i == 0);
                bodyBox.TextAutoFit = PowerPointTextAutoFit.Normal;
                bodyBox.TextAutoFitOptions = new PowerPointTextAutoFitOptions(fontScalePercent: 80, lineSpaceReductionPercent: 20);
            }
        }

        private static void AddCaseStudyEditorialSplit(PowerPointSlide slide, PowerPointDesignTheme theme,
            string clientTitle, IReadOnlyList<PowerPointCaseStudySection> sections, PowerPointCaseStudySlideOptions options,
            IReadOnlyList<PowerPointMetric> metrics, double slideWidthCm, double slideHeightCm) {
            AddText(slide, clientTitle, 1.45, 1.6, slideWidthCm * 0.58, 1.35, 22,
                theme.PrimaryTextColor, theme.HeadingFontName, bold: true);

            PowerPointAutoShape titleRule = slide.AddRectangleCm(1.47, 3.18, slideWidthCm * 0.24, 0.08,
                "Case Study Editorial Rule");
            titleRule.FillColor = theme.AccentColor;
            titleRule.OutlineColor = theme.AccentColor;

            int textCount = Math.Min(sections.Count, 4);
            PowerPointLayoutBox[,] boxes = PowerPointLayoutBox
                .FromCentimeters(1.45, 3.75, slideWidthCm * 0.54, 5.05)
                .SplitGridCm(textCount > 2 ? 2 : 1, textCount > 1 ? 2 : 1, 0.55, 0.55);

            for (int i = 0; i < textCount; i++) {
                int row = i / (textCount > 1 ? 2 : 1);
                int column = i % (textCount > 1 ? 2 : 1);
                PowerPointLayoutBox box = boxes[row, column];
                PowerPointCaseStudySection section = sections[i];

                PowerPointAutoShape panel = slide.AddRectangleCm(box.LeftCm, box.TopCm, box.WidthCm, box.HeightCm,
                    "Case Study Editorial Section " + (i + 1));
                panel.FillColor = theme.PanelColor;
                panel.OutlineColor = theme.PanelBorderColor;
                panel.OutlineWidthPoints = 0.45;
                panel.SetShadow("000000", blurPoints: 3, distancePoints: 0.8, angleDegrees: 90, transparencyPercent: 90);

                PowerPointAutoShape accent = slide.AddRectangleCm(box.LeftCm, box.TopCm, 0.12, box.HeightCm,
                    "Case Study Editorial Accent " + (i + 1));
                accent.FillColor = GetAccent(theme, i);
                accent.OutlineColor = GetAccent(theme, i);

                AddText(slide, section.Heading, box.LeftCm + 0.45, box.TopCm + 0.38, box.WidthCm - 0.75, 0.45, 11,
                    theme.PrimaryTextColor, theme.HeadingFontName, bold: true);
                PowerPointTextBox body = AddText(slide, section.Body, box.LeftCm + 0.45, box.TopCm + 1.0,
                    box.WidthCm - 0.75, box.HeightCm - 1.25, 9, theme.SecondaryTextColor, theme.BodyFontName);
                body.TextAutoFitOptions = new PowerPointTextAutoFitOptions(fontScalePercent: 82, lineSpaceReductionPercent: 18);
            }

            double visualLeft = slideWidthCm * 0.68;
            double visualTop = 2.15;
            double visualWidth = slideWidthCm - visualLeft - 1.45;
            double visualHeight = 4.75;
            AddVisualFrame(slide, theme, options.VisualImagePath, visualLeft, visualTop, visualWidth, visualHeight,
                options.VisualFrameVariant, options.DesignIntent);

            if (metrics.Count > 0) {
                PowerPointAutoShape metricBand = slide.AddRectangleCm(visualLeft, visualTop + visualHeight + 0.55,
                    visualWidth, 1.8, "Case Study Editorial Metric Band");
                metricBand.FillColor = theme.AccentColor;
                metricBand.OutlineColor = theme.AccentColor;
                metricBand.SetShadow("000000", blurPoints: 3, distancePoints: 0.8, angleDegrees: 90, transparencyPercent: 88);
                AddMetrics(slide, theme, metrics, visualLeft + 0.35, visualTop + visualHeight + 0.8,
                    visualWidth - 0.7, 1.35);
            }

            AddTags(slide, theme, options.Tags, 1.45, slideHeightCm - 2.05, slideWidthCm * 0.52, 0.55);
        }

        private static void AddCaseStudyBand(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointCaseStudySlideOptions options, IReadOnlyList<PowerPointMetric> metrics,
            double slideWidthCm, double slideHeightCm) {
            double bandTop = slideHeightCm * 0.55;
            double bandHeight = slideHeightCm * 0.36;
            PowerPointAutoShape band = slide.AddRectangleCm(1.2, bandTop, slideWidthCm - 2.4, bandHeight,
                "Case Study Visual Band");
            band.FillColor = theme.AccentColor;
            band.OutlineColor = theme.AccentColor;
            band.SetSoftEdges(1.1);

            PowerPointAutoShape wash = slide.AddShapeCm(A.ShapeTypeValues.Parallelogram, 8.0, bandTop + 0.25,
                7.2, bandHeight - 0.5, "Case Study Band Wash");
            wash.FillColor = theme.Accent2Color;
            wash.FillTransparency = 45;
            wash.OutlineColor = theme.Accent2Color;

            AddText(slide, options.BrandText ?? string.Empty, 2.0, bandTop + 0.85, 4.4, 0.7, 20,
                theme.AccentContrastColor, theme.HeadingFontName, bold: true);

            if (!string.IsNullOrWhiteSpace(options.BandLabel)) {
                AddText(slide, options.BandLabel!, 2.0, bandTop + bandHeight - 1.35, 6.8, 0.5, 17,
                    theme.AccentLightColor, theme.HeadingFontName, bold: true);
            }

            if (!string.IsNullOrWhiteSpace(options.PersonImagePath)) {
                AddPictureIfExists(slide, options.PersonImagePath!, 9.2, bandTop - 1.0, 5.1, bandHeight + 0.4, crop: true);
            }

            AddMetrics(slide, theme, metrics, slideWidthCm * 0.46, bandTop + 1.75, slideWidthCm * 0.22, 1.65);
            AddVisualFrame(slide, theme, options.VisualImagePath, slideWidthCm - 8.6, bandTop + 0.9, 6.8, bandHeight - 1.5,
                options.VisualFrameVariant, options.DesignIntent);
            AddTags(slide, theme, options.Tags, 9.4, bandTop + bandHeight - 1.15, slideWidthCm - 12.6, 0.7);
        }

        internal static void AddMetrics(PowerPointSlide slide, PowerPointDesignTheme theme, IReadOnlyList<PowerPointMetric> metrics,
            double leftCm, double topCm, double widthCm, double heightCm) {
            if (metrics.Count == 0) {
                return;
            }

            int count = Math.Min(metrics.Count, 3);
            PowerPointLayoutBox[] boxes = PowerPointLayoutBox
                .FromCentimeters(leftCm, topCm, widthCm, heightCm)
                .SplitColumnsCm(count, 0.45);
            double valueHeight = Math.Min(0.88, heightCm * 0.52);
            double labelTopOffset = valueHeight + Math.Min(0.14, heightCm * 0.08);
            double labelHeight = Math.Max(0.32, heightCm - labelTopOffset);
            int valueFontSize = heightCm < 1.6 ? 24 : 29;
            int labelFontSize = heightCm < 1.6 ? 8 : 9;
            for (int i = 0; i < count; i++) {
                PowerPointMetric metric = metrics[i];
                PowerPointLayoutBox box = boxes[i];
                int resolvedValueFontSize = ResolveMetricValueFontSize(metric.Value, box.WidthCm, valueFontSize);
                PowerPointTextBox value = AddText(slide, metric.Value, box.LeftCm, box.TopCm, box.WidthCm, valueHeight,
                    resolvedValueFontSize,
                    theme.AccentContrastColor, theme.HeadingFontName, bold: true);
                CenterText(value);
                int resolvedLabelFontSize = ResolveMetricLabelFontSize(metric.Label, box.WidthCm, labelFontSize);
                PowerPointTextBox label = AddText(slide, metric.Label, box.LeftCm, box.TopCm + labelTopOffset,
                    box.WidthCm, labelHeight, resolvedLabelFontSize,
                    theme.AccentContrastColor, theme.BodyFontName, bold: true);
                CenterText(label);
            }
        }

        internal static void AddVisualFrame(PowerPointSlide slide, PowerPointDesignTheme theme, string? imagePath,
            double leftCm, double topCm, double widthCm, double heightCm, PowerPointDesignIntent? intent = null) {
            AddVisualFrame(slide, theme, imagePath, leftCm, topCm, widthCm, heightCm,
                PowerPointVisualFrameVariant.Auto, intent);
        }

        internal static void AddVisualFrame(PowerPointSlide slide, PowerPointDesignTheme theme, string? imagePath,
            double leftCm, double topCm, double widthCm, double heightCm, PowerPointVisualFrameVariant variant,
            PowerPointDesignIntent? intent = null) {
            PowerPointAutoShape frame = slide.AddRectangleCm(leftCm, topCm, widthCm, heightCm, "Case Study Visual Frame");
            frame.FillColor = theme.AccentDarkColor;
            frame.OutlineColor = theme.AccentDarkColor;
            frame.OutlineWidthPoints = 0;
            frame.SetShadow("000000", blurPoints: 5, distancePoints: 1.5, angleDegrees: 90, transparencyPercent: 82);

            PowerPointVisualFrameVariant resolvedVariant = ResolveVisualPlaceholderVariant(variant, intent);
            if (!string.IsNullOrWhiteSpace(imagePath) && File.Exists(imagePath)) {
                AddVisualPicture(slide, theme, imagePath!, leftCm + 0.08, topCm + 0.08,
                    widthCm - 0.16, heightCm - 0.16, resolvedVariant);
                return;
            }

            AddVisualPlaceholder(slide, theme, leftCm + 0.08, topCm + 0.08, widthCm - 0.16, heightCm - 0.16,
                resolvedVariant, intent);
        }

        private static void AddVisualPicture(PowerPointSlide slide, PowerPointDesignTheme theme, string imagePath,
            double leftCm, double topCm, double widthCm, double heightCm, PowerPointVisualFrameVariant variant) {
            if (variant == PowerPointVisualFrameVariant.DeviceMockup) {
                AddVisualDeviceChrome(slide, theme, leftCm, topCm, widthCm, heightCm);
                GetVisualDeviceContentBounds(leftCm, topCm, widthCm, heightCm, out double contentLeft,
                    out double contentTop, out double contentWidth, out double contentHeight);
                AddPictureIfExists(slide, imagePath, contentLeft, contentTop, contentWidth, contentHeight,
                    crop: true);
                return;
            }

            if (variant == PowerPointVisualFrameVariant.ProofBoard) {
                AddVisualProofMat(slide, theme, leftCm, topCm, widthCm, heightCm);
                AddPictureIfExists(slide, imagePath, leftCm + widthCm * 0.10, topCm + heightCm * 0.14,
                    widthCm * 0.78, heightCm * 0.68, crop: true);
                return;
            }

            AddPictureIfExists(slide, imagePath, leftCm, topCm, widthCm, heightCm, crop: true);
        }

        private static void AddVisualPlaceholder(PowerPointSlide slide, PowerPointDesignTheme theme,
            double leftCm, double topCm, double widthCm, double heightCm, PowerPointVisualFrameVariant variant,
            PowerPointDesignIntent? intent) {
            PowerPointAutoShape surface = slide.AddRectangleCm(leftCm, topCm, widthCm, heightCm,
                "Case Study Visual Surface");
            surface.FillColor = theme.AccentDarkColor;
            surface.OutlineColor = theme.AccentDarkColor;

            PowerPointVisualFrameVariant resolvedVariant = ResolveVisualPlaceholderVariant(variant, intent);
            if (resolvedVariant == PowerPointVisualFrameVariant.Collage) {
                AddVisualCollagePlaceholder(slide, theme, leftCm, topCm, widthCm, heightCm);
                return;
            }
            if (resolvedVariant == PowerPointVisualFrameVariant.Diagram) {
                AddVisualDiagramPlaceholder(slide, theme, leftCm, topCm, widthCm, heightCm);
                return;
            }
            if (resolvedVariant == PowerPointVisualFrameVariant.DeviceMockup) {
                AddVisualDeviceMockupPlaceholder(slide, theme, leftCm, topCm, widthCm, heightCm);
                return;
            }
            if (resolvedVariant == PowerPointVisualFrameVariant.ProofBoard) {
                AddVisualProofBoardPlaceholder(slide, theme, leftCm, topCm, widthCm, heightCm);
                return;
            }

            AddVisualDashboardPlaceholder(slide, theme, leftCm, topCm, widthCm, heightCm);
        }

        private static void AddVisualDashboardPlaceholder(PowerPointSlide slide, PowerPointDesignTheme theme,
            double leftCm, double topCm, double widthCm, double heightCm) {
            PowerPointAutoShape glow = slide.AddShapeCm(A.ShapeTypeValues.Parallelogram, leftCm + widthCm * 0.08,
                topCm, widthCm * 0.42, heightCm, "Case Study Visual Wash");
            glow.FillColor = theme.Accent2Color;
            glow.FillTransparency = 68;
            glow.OutlineColor = theme.Accent2Color;

            double panelTop = topCm + heightCm * 0.24;
            double panelLeft = leftCm + widthCm * 0.17;
            double panelWidth = widthCm * 0.66;
            double panelHeight = heightCm * 0.42;
            PowerPointAutoShape panel = slide.AddRectangleCm(panelLeft, panelTop, panelWidth, panelHeight,
                "Case Study Visual Content Panel");
            panel.FillColor = theme.AccentColor;
            panel.FillTransparency = 35;
            panel.OutlineColor = theme.AccentColor;
            panel.OutlineWidthPoints = 0;

            PowerPointAutoShape imageBlock = slide.AddRectangleCm(panelLeft + panelWidth * 0.08, panelTop + panelHeight * 0.18,
                panelWidth * 0.34, panelHeight * 0.56, "Case Study Visual Image Block");
            imageBlock.FillColor = theme.AccentLightColor;
            imageBlock.FillTransparency = 10;
            imageBlock.OutlineColor = theme.AccentLightColor;

            for (int i = 0; i < 3; i++) {
                double barWidth = panelWidth * (i == 0 ? 0.36 : 0.28);
                PowerPointAutoShape bar = slide.AddRectangleCm(panelLeft + panelWidth * 0.49,
                    panelTop + panelHeight * (0.23 + i * 0.18), barWidth, 0.045,
                    "Case Study Visual Line " + (i + 1));
                bar.FillColor = theme.AccentLightColor;
                bar.FillTransparency = i == 0 ? 5 : 35;
                bar.OutlineColor = theme.AccentLightColor;
            }

            PowerPointAutoShape baseLine = slide.AddLineCm(leftCm + widthCm * 0.18, topCm + heightCm * 0.84,
                leftCm + widthCm * 0.82, topCm + heightCm * 0.84, "Case Study Visual Base Line");
            baseLine.OutlineColor = theme.AccentLightColor;
            baseLine.OutlineWidthPoints = 0.9;
        }

        private static void AddVisualCollagePlaceholder(PowerPointSlide slide, PowerPointDesignTheme theme,
            double leftCm, double topCm, double widthCm, double heightCm) {
            PowerPointAutoShape wash = slide.AddShapeCm(A.ShapeTypeValues.Parallelogram, leftCm + widthCm * 0.58,
                topCm, widthCm * 0.24, heightCm, "Visual Collage Wash");
            wash.FillColor = theme.Accent2Color;
            wash.FillTransparency = 62;
            wash.OutlineColor = theme.Accent2Color;

            AddVisualTile(slide, theme, leftCm + widthCm * 0.10, topCm + heightCm * 0.17,
                widthCm * 0.38, heightCm * 0.42, "Visual Collage Tile 1", theme.AccentLightColor, 62);
            AddVisualTile(slide, theme, leftCm + widthCm * 0.43, topCm + heightCm * 0.09,
                widthCm * 0.43, heightCm * 0.28, "Visual Collage Tile 2", theme.AccentColor, 70);
            AddVisualTile(slide, theme, leftCm + widthCm * 0.36, topCm + heightCm * 0.55,
                widthCm * 0.48, heightCm * 0.25, "Visual Collage Tile 3", theme.Accent3Color, 72);

            for (int i = 0; i < 3; i++) {
                double dot = 0.18 - i * 0.025;
                PowerPointAutoShape node = slide.AddEllipseCm(leftCm + widthCm * (0.24 + i * 0.19),
                    topCm + heightCm * (0.77 - i * 0.08), dot, dot, "Visual Collage Marker " + (i + 1));
                node.FillColor = GetAccent(theme, i);
                node.OutlineColor = theme.AccentLightColor;
                node.OutlineWidthPoints = 0.55;
            }
        }

        private static void AddVisualDiagramPlaceholder(PowerPointSlide slide, PowerPointDesignTheme theme,
            double leftCm, double topCm, double widthCm, double heightCm) {
            PowerPointAutoShape rail = slide.AddRectangleCm(leftCm + widthCm * 0.12, topCm + heightCm * 0.49,
                widthCm * 0.76, 0.045, "Visual Diagram Rail");
            rail.FillColor = theme.AccentLightColor;
            rail.FillTransparency = 15;
            rail.OutlineColor = theme.AccentLightColor;

            for (int i = 0; i < 4; i++) {
                double cx = leftCm + widthCm * (0.16 + i * 0.22);
                double cy = topCm + heightCm * (i % 2 == 0 ? 0.35 : 0.61);
                PowerPointAutoShape node = slide.AddEllipseCm(cx, cy, 0.58, 0.58, "Visual Diagram Node " + (i + 1));
                node.FillColor = GetAccent(theme, i);
                node.FillTransparency = i == 0 ? 0 : 22;
                node.OutlineColor = theme.AccentLightColor;
                node.OutlineWidthPoints = 0.55;

                PowerPointAutoShape label = slide.AddRectangleCm(cx + 0.78, cy + 0.22,
                    widthCm * (i == 0 ? 0.22 : 0.16), 0.04, "Visual Diagram Label " + (i + 1));
                label.FillColor = theme.AccentLightColor;
                label.FillTransparency = i == 0 ? 0 : 35;
                label.OutlineColor = theme.AccentLightColor;
            }

            PowerPointAutoShape plate = slide.AddShapeCm(A.ShapeTypeValues.Parallelogram, leftCm + widthCm * 0.20,
                topCm + heightCm * 0.14, widthCm * 0.50, heightCm * 0.72, "Visual Diagram Plate");
            plate.FillColor = theme.AccentColor;
            plate.FillTransparency = 82;
            plate.OutlineColor = theme.AccentLightColor;
            plate.OutlineWidthPoints = 0.45;
        }

        private static void AddVisualDeviceMockupPlaceholder(PowerPointSlide slide, PowerPointDesignTheme theme,
            double leftCm, double topCm, double widthCm, double heightCm) {
            AddVisualDeviceChrome(slide, theme, leftCm, topCm, widthCm, heightCm);

            GetVisualDeviceContentBounds(leftCm, topCm, widthCm, heightCm, out double contentLeft,
                out double contentTop, out double contentWidth, out double contentHeight);

            PowerPointAutoShape hero = slide.AddRectangleCm(contentLeft, contentTop, contentWidth, contentHeight,
                "Visual Device Hero Area");
            hero.FillColor = theme.AccentColor;
            hero.FillTransparency = 34;
            hero.OutlineColor = theme.AccentColor;
            hero.OutlineWidthPoints = 0;

            for (int i = 0; i < 3; i++) {
                PowerPointAutoShape bar = slide.AddRectangleCm(contentLeft + contentWidth * 0.09,
                    contentTop + contentHeight * (0.22 + i * 0.18),
                    contentWidth * (i == 0 ? 0.38 : 0.25), 0.06, "Visual Device Content Line " + (i + 1));
                bar.FillColor = theme.AccentLightColor;
                bar.FillTransparency = i == 0 ? 12 : 42;
                bar.OutlineColor = theme.AccentLightColor;
                bar.OutlineWidthPoints = 0;
            }

            for (int i = 0; i < 3; i++) {
                PowerPointAutoShape tile = slide.AddRectangleCm(contentLeft + contentWidth * (0.58 + i * 0.11),
                    contentTop + contentHeight * 0.22, contentWidth * 0.08, contentHeight * 0.46,
                    "Visual Device Metric Tile " + (i + 1));
                tile.FillColor = GetAccent(theme, i);
                tile.FillTransparency = 18 + i * 10;
                tile.OutlineColor = theme.AccentLightColor;
                tile.OutlineWidthPoints = 0.35;
            }
        }

        private static void AddVisualDeviceChrome(PowerPointSlide slide, PowerPointDesignTheme theme,
            double leftCm, double topCm, double widthCm, double heightCm) {
            PowerPointAutoShape screen = slide.AddRectangleCm(leftCm + widthCm * 0.06, topCm + heightCm * 0.08,
                widthCm * 0.88, heightCm * 0.76, "Visual Device Screen");
            screen.FillColor = theme.AccentDarkColor;
            screen.FillTransparency = 18;
            screen.OutlineColor = theme.AccentLightColor;
            screen.OutlineWidthPoints = 0.45;

            PowerPointAutoShape topBar = slide.AddRectangleCm(leftCm + widthCm * 0.06, topCm + heightCm * 0.08,
                widthCm * 0.88, 0.38, "Visual Device Chrome Bar");
            topBar.FillColor = theme.AccentLightColor;
            topBar.FillTransparency = 8;
            topBar.OutlineColor = theme.AccentLightColor;
            topBar.OutlineWidthPoints = 0;

            for (int i = 0; i < 3; i++) {
                PowerPointAutoShape dot = slide.AddEllipseCm(leftCm + widthCm * 0.10 + i * 0.18,
                    topCm + heightCm * 0.08 + 0.12, 0.09, 0.09, "Visual Device Chrome Dot " + (i + 1));
                dot.FillColor = GetAccent(theme, i);
                dot.FillTransparency = i == 0 ? 0 : 18;
                dot.OutlineColor = dot.FillColor;
                dot.OutlineWidthPoints = 0;
            }

            PowerPointAutoShape baseLine = slide.AddLineCm(leftCm + widthCm * 0.18, topCm + heightCm * 0.89,
                leftCm + widthCm * 0.82, topCm + heightCm * 0.89, "Visual Device Base");
            baseLine.OutlineColor = theme.AccentLightColor;
            baseLine.OutlineWidthPoints = 1.1;
        }

        private static void GetVisualDeviceContentBounds(double leftCm, double topCm, double widthCm,
            double heightCm, out double contentLeft, out double contentTop, out double contentWidth,
            out double contentHeight) {
            double screenLeft = leftCm + widthCm * 0.06;
            double screenTop = topCm + heightCm * 0.08;
            double screenWidth = widthCm * 0.88;
            double screenHeight = heightCm * 0.76;

            contentLeft = screenLeft + widthCm * 0.05;
            contentTop = screenTop + 0.46;
            contentWidth = screenWidth - widthCm * 0.10;
            contentHeight = Math.Max(0.8, screenHeight - 0.54);
        }

        private static void AddVisualProofBoardPlaceholder(PowerPointSlide slide, PowerPointDesignTheme theme,
            double leftCm, double topCm, double widthCm, double heightCm) {
            AddVisualProofMat(slide, theme, leftCm, topCm, widthCm, heightCm);

            AddVisualProofPanel(slide, theme, leftCm + widthCm * 0.10, topCm + heightCm * 0.15,
                widthCm * 0.36, heightCm * 0.55, "Visual Proof Primary Panel", theme.PanelColor, 0);
            AddVisualProofPanel(slide, theme, leftCm + widthCm * 0.53, topCm + heightCm * 0.18,
                widthCm * 0.34, heightCm * 0.24, "Visual Proof Detail Panel 1", theme.AccentLightColor, 10);
            AddVisualProofPanel(slide, theme, leftCm + widthCm * 0.58, topCm + heightCm * 0.48,
                widthCm * 0.28, heightCm * 0.22, "Visual Proof Detail Panel 2", theme.AccentLightColor, 24);

            for (int i = 0; i < 3; i++) {
                PowerPointAutoShape rule = slide.AddRectangleCm(leftCm + widthCm * 0.15,
                    topCm + heightCm * (0.78 + i * 0.055), widthCm * (0.46 - i * 0.06), 0.045,
                    "Visual Proof Caption Line " + (i + 1));
                rule.FillColor = theme.AccentLightColor;
                rule.FillTransparency = 18 + i * 18;
                rule.OutlineColor = theme.AccentLightColor;
                rule.OutlineWidthPoints = 0;
            }
        }

        private static void AddVisualProofMat(PowerPointSlide slide, PowerPointDesignTheme theme,
            double leftCm, double topCm, double widthCm, double heightCm) {
            PowerPointAutoShape mat = slide.AddRectangleCm(leftCm + widthCm * 0.05, topCm + heightCm * 0.08,
                widthCm * 0.90, heightCm * 0.78, "Visual Proof Mat");
            mat.FillColor = theme.AccentContrastColor;
            mat.FillTransparency = 8;
            mat.OutlineColor = theme.AccentLightColor;
            mat.OutlineWidthPoints = 0.4;

            PowerPointAutoShape accent = slide.AddRectangleCm(leftCm + widthCm * 0.05, topCm + heightCm * 0.08,
                widthCm * 0.90, 0.08, "Visual Proof Mat Accent");
            accent.FillColor = theme.Accent2Color;
            accent.OutlineColor = theme.Accent2Color;
            accent.OutlineWidthPoints = 0;
        }

        private static void AddVisualProofPanel(PowerPointSlide slide, PowerPointDesignTheme theme,
            double leftCm, double topCm, double widthCm, double heightCm, string name, string fillColor,
            int fillTransparency) {
            PowerPointAutoShape panel = slide.AddRectangleCm(leftCm, topCm, widthCm, heightCm, name);
            panel.FillColor = fillColor;
            panel.FillTransparency = fillTransparency;
            panel.OutlineColor = theme.PanelBorderColor;
            panel.OutlineWidthPoints = 0.35;
        }

        private static void AddVisualTile(PowerPointSlide slide, PowerPointDesignTheme theme,
            double leftCm, double topCm, double widthCm, double heightCm, string name, string fillColor,
            int fillTransparency) {
            PowerPointAutoShape tile = slide.AddShapeCm(A.ShapeTypeValues.Parallelogram, leftCm, topCm, widthCm, heightCm, name);
            tile.FillColor = fillColor;
            tile.FillTransparency = fillTransparency;
            tile.OutlineColor = theme.AccentLightColor;
            tile.OutlineWidthPoints = 0.45;
        }

        private static PowerPointVisualFrameVariant ResolveVisualPlaceholderVariant(PowerPointVisualFrameVariant variant,
            PowerPointDesignIntent? intent) {
            if (variant != PowerPointVisualFrameVariant.Auto) {
                return variant;
            }
            if (intent == null) {
                return PowerPointVisualFrameVariant.Dashboard;
            }
            if (string.IsNullOrWhiteSpace(intent.Seed) &&
                intent.Mood == PowerPointDesignMood.Corporate &&
                intent.Density == PowerPointSlideDensity.Balanced &&
                intent.VisualStyle == PowerPointVisualStyle.Geometric) {
                return PowerPointVisualFrameVariant.Dashboard;
            }
            if (intent.Mood == PowerPointDesignMood.Energetic) {
                return PowerPointVisualFrameVariant.DeviceMockup;
            }
            if (intent.Mood == PowerPointDesignMood.Editorial) {
                return PowerPointVisualFrameVariant.ProofBoard;
            }
            if (intent.VisualStyle == PowerPointVisualStyle.Soft) {
                return PowerPointVisualFrameVariant.Collage;
            }
            if (intent.VisualStyle == PowerPointVisualStyle.Minimal) {
                return PowerPointVisualFrameVariant.Diagram;
            }

            return intent.Pick(5, "visual-placeholder") switch {
                0 => PowerPointVisualFrameVariant.Dashboard,
                1 => PowerPointVisualFrameVariant.Collage,
                2 => PowerPointVisualFrameVariant.Diagram,
                3 => PowerPointVisualFrameVariant.DeviceMockup,
                _ => PowerPointVisualFrameVariant.ProofBoard
            };
        }

        private static void AddTags(PowerPointSlide slide, PowerPointDesignTheme theme, IList<string> tags,
            double leftCm, double topCm, double widthCm, double heightCm) {
            if (tags.Count == 0) {
                return;
            }

            int count = Math.Min(tags.Count, 7);
            PowerPointLayoutBox[] boxes = PowerPointLayoutBox
                .FromCentimeters(leftCm, topCm, widthCm, heightCm)
                .SplitColumnsCm(count, 0.28);
            for (int i = 0; i < count; i++) {
                PowerPointLayoutBox box = boxes[i];
                PowerPointAutoShape pill = slide.AddRectangleCm(box.LeftCm, box.TopCm, box.WidthCm, box.HeightCm,
                    "Case Study Tag " + (i + 1));
                pill.FillColor = i == count - 1 ? theme.Accent2Color : theme.AccentColor;
                pill.FillTransparency = i == count - 1 ? 0 : 40;
                pill.OutlineColor = theme.AccentLightColor;
                pill.OutlineWidthPoints = 0.6;

                PowerPointTextBox label = AddText(slide, tags[i], box.LeftCm + 0.08, box.TopCm + 0.16,
                    box.WidthCm - 0.16, box.HeightCm - 0.25, 8, theme.AccentContrastColor, theme.BodyFontName,
                    bold: i == count - 1);
                CenterText(label);
            }
        }

        internal static void AddCardGrid(PowerPointSlide slide, PowerPointDesignTheme theme,
            IReadOnlyList<PowerPointCardContent> cards, PowerPointCardGridSlideOptions options,
            PowerPointCardGridLayoutVariant variant, double slideWidthCm, double slideHeightCm) {
            double top = options.DesignIntent.Density == PowerPointSlideDensity.Relaxed ? 4.35 : 4.05;
            double height = string.IsNullOrWhiteSpace(options.SupportingText)
                ? slideHeightCm - 6.0
                : slideHeightCm - 8.7;
            PowerPointLayoutBox bounds = PowerPointLayoutBox.FromCentimeters(1.5, top, slideWidthCm - 3.0, height);
            AddCardGrid(slide, theme, cards, options, variant, bounds);

            if (!string.IsNullOrWhiteSpace(options.SupportingText)) {
                PowerPointAutoShape band = slide.AddRectangleCm(1.55, slideHeightCm - 3.25, slideWidthCm - 3.1, 1.8,
                    "Designer Supporting Band");
                band.FillColor = theme.PanelColor;
                band.OutlineColor = theme.PanelBorderColor;
                band.SetShadow("000000", blurPoints: 4, distancePoints: 1, angleDegrees: 90, transparencyPercent: 88);
                AddText(slide, options.SupportingText!, 2.15, slideHeightCm - 2.8, slideWidthCm - 4.3, 0.9, 13,
                    theme.SecondaryTextColor, theme.BodyFontName);
            }
        }

        internal static void AddCardGrid(PowerPointSlide slide, PowerPointDesignTheme theme,
            IReadOnlyList<PowerPointCardContent> cards, PowerPointCardGridSlideOptions options,
            PowerPointCardGridLayoutVariant variant, PowerPointLayoutBox bounds) {
            int maxColumns = Math.Max(1, options.MaxColumns);
            int columns = Math.Min(maxColumns, cards.Count);
            int rows = (int)Math.Ceiling(cards.Count / (double)columns);
            double columnGap = variant == PowerPointCardGridLayoutVariant.SoftTiles ? 0.42 : 0.65;
            double rowGap = variant == PowerPointCardGridLayoutVariant.SoftTiles ? 0.42 : 0.55;
            PowerPointLayoutBox[,] grid = PowerPointLayoutBox
                .FromCentimeters(bounds.LeftCm, bounds.TopCm, bounds.WidthCm, bounds.HeightCm)
                .SplitGridCm(rows, columns, rowGap, columnGap);
            PowerPointCardSurfaceStyle surfaceStyle = ResolveCardSurfaceStyle(options, variant);

            for (int i = 0; i < cards.Count; i++) {
                int row = i / columns;
                int column = i % columns;
                AddDesignerCard(slide, theme, cards[i], grid[row, column], i, variant, surfaceStyle);
            }
        }

        private static void AddDesignerCard(PowerPointSlide slide, PowerPointDesignTheme theme, PowerPointCardContent card,
            PowerPointLayoutBox box, int index, PowerPointCardGridLayoutVariant variant,
            PowerPointCardSurfaceStyle surfaceStyle) {
            string accent = card.AccentColor ?? GetAccent(theme, index);
            PowerPointAutoShape panel = slide.AddRectangleCm(box.LeftCm, box.TopCm, box.WidthCm, box.HeightCm,
                "Designer Card " + (index + 1));
            ApplyDesignerCardSurface(panel, theme, accent, variant, surfaceStyle);

            if (variant == PowerPointCardGridLayoutVariant.SoftTiles) {
                PowerPointAutoShape accentStrip = slide.AddRectangleCm(box.LeftCm, box.TopCm, 0.13, box.HeightCm,
                    "Designer Card Accent " + (index + 1));
                accentStrip.FillColor = accent;
                accentStrip.OutlineColor = accent;
            } else {
                PowerPointAutoShape accentBar = slide.AddRectangleCm(box.LeftCm, box.TopCm, box.WidthCm, 0.18,
                    "Designer Card Accent " + (index + 1));
                accentBar.FillColor = accent;
                accentBar.OutlineColor = accent;
            }

            double titleLeft = variant == PowerPointCardGridLayoutVariant.SoftTiles ? box.LeftCm + 0.6 : box.LeftCm + 0.45;
            double titleWidth = Math.Max(0.5, box.WidthCm - 0.9);
            int titleFontSize = ResolveCardTitleFontSize(card.Title, titleWidth);
            double titleHeight = ResolveCardTitleHeight(card.Title, titleWidth, titleFontSize, box.HeightCm);
            AddText(slide, card.Title, titleLeft, box.TopCm + 0.65, titleWidth, titleHeight,
                titleFontSize, theme.PrimaryTextColor, theme.HeadingFontName, bold: true);

            double bodyTop = box.TopCm + 0.65 + titleHeight + 0.28;
            double bodyHeight = Math.Max(0.42, box.HeightCm - (bodyTop - box.TopCm) - 0.35);
            int bodyFontSize = ResolveCardBodyFontSize(card.Items, titleWidth, bodyHeight);
            PowerPointTextBox body = slide.AddTextBox("", PowerPointLayoutBox.FromCentimeters(
                titleLeft + 0.1, bodyTop, box.WidthCm - 1.05, bodyHeight));
            body.SetTextMarginsCm(0, 0, 0, 0);
            body.TextAutoFit = PowerPointTextAutoFit.Normal;

            if (card.Items.Count == 0) {
                body.SetParagraphs(new[] { " " });
                return;
            }

            int bulletSpaceAfter = bodyHeight < 1.15 || card.Items.Count > 3 ? 2 : 4;
            body.SetBullets(card.Items.Select(item => " " + item), configure: paragraph => {
                paragraph.SetFontName(theme.BodyFontName)
                    .SetFontSize(bodyFontSize)
                    .SetColor(theme.SecondaryTextColor)
                    .SetHangingPoints(bodyFontSize <= 8 ? 12 : 16)
                    .SetSpaceAfterPoints(bulletSpaceAfter)
                    .SetBulletSizePercent(70);
            });
        }

        private static void ApplyDesignerCardSurface(PowerPointAutoShape panel, PowerPointDesignTheme theme,
            string accent, PowerPointCardGridLayoutVariant variant, PowerPointCardSurfaceStyle surfaceStyle) {
            panel.FillColor = variant == PowerPointCardGridLayoutVariant.SoftTiles
                ? theme.SurfaceColor
                : theme.PanelColor;
            panel.FillTransparency = 0;
            panel.OutlineColor = theme.PanelBorderColor;
            panel.OutlineWidthPoints = variant == PowerPointCardGridLayoutVariant.SoftTiles ? 0.35 : 0.8;

            switch (surfaceStyle) {
                case PowerPointCardSurfaceStyle.Flat:
                    panel.FillColor = theme.SurfaceColor;
                    panel.OutlineWidthPoints = 0.25;
                    break;
                case PowerPointCardSurfaceStyle.Hairline:
                    panel.FillColor = theme.SurfaceColor;
                    panel.OutlineColor = theme.PanelBorderColor;
                    panel.OutlineWidthPoints = 0.35;
                    break;
                case PowerPointCardSurfaceStyle.AccentWash:
                    panel.FillColor = accent;
                    panel.FillTransparency = 88;
                    panel.OutlineColor = accent;
                    panel.OutlineWidthPoints = 0.25;
                    break;
                default:
                    panel.SetShadow("000000", blurPoints: variant == PowerPointCardGridLayoutVariant.SoftTiles ? 3 : 5,
                        distancePoints: variant == PowerPointCardGridLayoutVariant.SoftTiles ? 0.8 : 1.5,
                        angleDegrees: 90, transparencyPercent: 88);
                    break;
            }
        }

        private static int ResolveCardTitleFontSize(string title, double widthCm) {
            int length = string.IsNullOrWhiteSpace(title) ? 0 : title.Trim().Length;
            if (widthCm < 3.4 || length > 46) {
                return 12;
            }
            if (widthCm < 4.2 || length > 32) {
                return 13;
            }
            return 15;
        }

        private static double ResolveCardTitleHeight(string title, double widthCm, int fontSize, double cardHeightCm) {
            int lines = EstimateWrappedLines(title, widthCm, fontSize);
            double desiredHeight = 0.58 + Math.Max(0, lines - 1) * 0.34;
            double maxHeight = Math.Min(1.35, Math.Max(0.62, cardHeightCm * 0.36));
            return Math.Min(maxHeight, Math.Max(0.62, desiredHeight));
        }

        private static int ResolveCardBodyFontSize(IReadOnlyList<string> items, double widthCm, double bodyHeightCm) {
            if (items.Count == 0) {
                return 10;
            }

            int longest = items.Max(item => string.IsNullOrWhiteSpace(item) ? 0 : item.Trim().Length);
            int estimatedLines = items.Sum(item => EstimateWrappedLines(item, widthCm, 10));
            if (bodyHeightCm < 1.05 || items.Count > 5 || estimatedLines > 7 || longest > 60) {
                return 8;
            }
            if (bodyHeightCm < 1.35 || items.Count > 3 || estimatedLines > 5 || longest > 42) {
                return 9;
            }
            return 10;
        }

        private static int ResolveMetricValueFontSize(string value, double widthCm, int preferredFontSize) {
            int length = string.IsNullOrWhiteSpace(value) ? 0 : value.Trim().Length;
            if (length <= 3) {
                return preferredFontSize;
            }

            int estimate = (int)Math.Floor(widthCm * 7.5 / Math.Max(1, length));
            return Math.Max(12, Math.Min(preferredFontSize, estimate));
        }

        private static int ResolveMetricLabelFontSize(string label, double widthCm, int preferredFontSize) {
            int length = string.IsNullOrWhiteSpace(label) ? 0 : label.Trim().Length;
            if (length <= 12) {
                return preferredFontSize;
            }

            int estimate = (int)Math.Floor(widthCm * 5.5 / Math.Max(1, length));
            return Math.Max(7, Math.Min(preferredFontSize, estimate));
        }

        private static int EstimateWrappedLines(string? text, double widthCm, int fontSize) {
            string textValue = text == null ? string.Empty : text.Trim();
            if (textValue.Length == 0) {
                return 1;
            }

            double charsPerLine = Math.Max(8, widthCm * (fontSize <= 10 ? 5.1 : 4.3));
            return Math.Max(1, (int)Math.Ceiling(textValue.Length / charsPerLine));
        }

        internal static void AddProcessTimeline(PowerPointSlide slide, PowerPointDesignTheme theme,
            IReadOnlyList<PowerPointProcessStep> steps, PowerPointProcessSlideOptions options,
            double slideWidthCm, double slideHeightCm) {
            PowerPointProcessLayoutVariant variant = ResolveProcessVariant(options, steps);
            if (variant == PowerPointProcessLayoutVariant.NumberedColumns) {
                AddProcessColumns(slide, theme, steps, options, slideWidthCm, slideHeightCm);
                return;
            }

            double left = options.DesignIntent.Density == PowerPointSlideDensity.Relaxed ? 2.35 : 2.1;
            double top = slideHeightCm * 0.47;
            double width = slideWidthCm - 4.2;
            double height = 4.7;
            AddProcessRailTimeline(slide, theme, steps, options,
                PowerPointLayoutBox.FromCentimeters(left, top, width, height));
        }

        internal static void AddProcessTimeline(PowerPointSlide slide, PowerPointDesignTheme theme,
            IReadOnlyList<PowerPointProcessStep> steps, PowerPointProcessSlideOptions options,
            PowerPointLayoutBox bounds) {
            PowerPointProcessLayoutVariant variant = ResolveProcessVariant(options, steps);
            if (variant == PowerPointProcessLayoutVariant.NumberedColumns) {
                AddProcessColumns(slide, theme, steps, options, bounds);
                return;
            }

            AddProcessRailTimeline(slide, theme, steps, options, bounds);
        }

        private static void AddProcessRailTimeline(PowerPointSlide slide, PowerPointDesignTheme theme,
            IReadOnlyList<PowerPointProcessStep> steps, PowerPointProcessSlideOptions options,
            PowerPointLayoutBox bounds) {
            int count = steps.Count;
            PowerPointLayoutBox[] boxes = PowerPointLayoutBox
                .FromCentimeters(bounds.LeftCm, bounds.TopCm, bounds.WidthCm, bounds.HeightCm)
                .SplitColumnsCm(count, count > 5 ? 0.45 : 0.75);

            double nodeSize = count > 5 ? 0.95 : 1.16;
            double railY = bounds.TopCm + nodeSize / 2;
            double railStart = boxes[0].LeftCm + nodeSize / 2;
            double railEnd = boxes[count - 1].LeftCm + nodeSize / 2;
            PowerPointProcessConnectorStyle connectorStyle = ResolveProcessConnectorStyle(options, steps);
            AddProcessConnectors(slide, theme, boxes, nodeSize, railY, railStart, railEnd, connectorStyle);

            for (int i = 0; i < count; i++) {
                PowerPointLayoutBox box = boxes[i];
                PowerPointProcessStep step = steps[i];
                string number = !string.IsNullOrWhiteSpace(step.Number) ? step.Number! : (i + 1) + ".";
                AddProcessNode(slide, theme, i, box.LeftCm, box.TopCm, nodeSize, number);
                AddText(slide, step.Title, box.LeftCm, box.TopCm + 1.55, box.WidthCm, 0.7, 13,
                    theme.AccentContrastColor, theme.HeadingFontName, bold: true);
                AddText(slide, step.Body, box.LeftCm, box.TopCm + 2.45, box.WidthCm, 1.7, 10,
                    theme.AccentLightColor, theme.BodyFontName);
            }
        }

        private static void AddProcessColumns(PowerPointSlide slide, PowerPointDesignTheme theme,
            IReadOnlyList<PowerPointProcessStep> steps, PowerPointProcessSlideOptions options,
            double slideWidthCm, double slideHeightCm) {
            double left = 1.85;
            double top = slideHeightCm * 0.45;
            double width = slideWidthCm - 3.7;
            double height = 4.85;
            AddProcessColumns(slide, theme, steps, options, PowerPointLayoutBox.FromCentimeters(left, top, width, height));
        }

        private static void AddProcessColumns(PowerPointSlide slide, PowerPointDesignTheme theme,
            IReadOnlyList<PowerPointProcessStep> steps, PowerPointProcessSlideOptions options,
            PowerPointLayoutBox bounds) {
            int count = steps.Count;
            double gutter = options.DesignIntent.Density == PowerPointSlideDensity.Compact ? 0.35 : 0.6;
            PowerPointLayoutBox[] boxes = PowerPointLayoutBox
                .FromCentimeters(bounds.LeftCm, bounds.TopCm, bounds.WidthCm, bounds.HeightCm)
                .SplitColumnsCm(count, gutter);

            for (int i = 0; i < count; i++) {
                PowerPointLayoutBox box = boxes[i];
                PowerPointProcessStep step = steps[i];
                string number = !string.IsNullOrWhiteSpace(step.Number) ? step.Number! : (i + 1).ToString("00");

                PowerPointAutoShape panel = slide.AddRectangleCm(box.LeftCm, box.TopCm, box.WidthCm, box.HeightCm,
                    "Process Column " + (i + 1));
                panel.FillColor = theme.AccentColor;
                panel.FillTransparency = 72;
                panel.OutlineColor = theme.AccentLightColor;
                panel.OutlineWidthPoints = 0.35;

                PowerPointAutoShape rule = slide.AddRectangleCm(box.LeftCm, box.TopCm, box.WidthCm * 0.34, 0.08,
                    "Process Column Rule " + (i + 1));
                rule.FillColor = GetAccent(theme, i);
                rule.OutlineColor = GetAccent(theme, i);

                AddText(slide, number.TrimEnd('.'), box.LeftCm + 0.26, box.TopCm + 0.45, box.WidthCm - 0.52, 0.85,
                    count > 5 ? 20 : 25, theme.AccentContrastColor, theme.HeadingFontName, bold: true);
                AddText(slide, step.Title, box.LeftCm + 0.26, box.TopCm + 1.65, box.WidthCm - 0.52, 0.72, 13,
                    theme.AccentContrastColor, theme.HeadingFontName, bold: true);
                AddText(slide, step.Body, box.LeftCm + 0.26, box.TopCm + 2.55, box.WidthCm - 0.52, 1.55, 10,
                    theme.AccentLightColor, theme.BodyFontName);
            }
        }

        private static void AddSubtleLightBackground(PowerPointSlide slide, PowerPointDesignTheme theme,
            double slideWidthCm, double slideHeightCm) {
            PowerPointAutoShape diagonal = slide.AddShapeCm(A.ShapeTypeValues.Parallelogram, slideWidthCm * 0.28, 0,
                slideWidthCm * 0.22, slideHeightCm, "Designer Light Diagonal");
            diagonal.FillColor = theme.SurfaceColor;
            diagonal.FillTransparency = 35;
            diagonal.OutlineColor = theme.SurfaceColor;
            diagonal.SendToBack();
        }

        private static void AddDiagonalPlanes(PowerPointSlide slide, PowerPointDesignTheme theme, double slideWidthCm,
            double slideHeightCm, bool dark) {
            string baseColor = dark ? theme.AccentColor : theme.SurfaceColor;
            string second = dark ? theme.AccentDarkColor : theme.AccentLightColor;

            PowerPointAutoShape left = slide.AddShapeCm(A.ShapeTypeValues.Parallelogram, -1.0, 0,
                slideWidthCm * 0.48, slideHeightCm, "Designer Plane Left");
            left.FillColor = baseColor;
            left.FillTransparency = dark ? 18 : 60;
            left.OutlineColor = baseColor;

            PowerPointAutoShape middle = slide.AddShapeCm(A.ShapeTypeValues.Parallelogram, slideWidthCm * 0.46, 0,
                slideWidthCm * 0.27, slideHeightCm, "Designer Plane Middle");
            middle.FillColor = second;
            middle.FillTransparency = dark ? 35 : 72;
            middle.OutlineColor = second;
        }

        private static void AddSectionTitleAccent(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointTitleAccentStyle style, double titleLeftCm, double titleTopCm, double titleWidthCm,
            double titleHeightCm, bool dark, bool centered = false) {
            if (style == PowerPointTitleAccentStyle.None) {
                return;
            }

            string accent = dark ? theme.WarningColor : theme.AccentColor;
            string secondary = dark ? theme.AccentLightColor : theme.WarningColor;
            switch (style) {
                case PowerPointTitleAccentStyle.SideRule:
                    PowerPointAutoShape sideRule = slide.AddRectangleCm(titleLeftCm - 0.28, titleTopCm + 0.16,
                        0.08, Math.Max(0.82, titleHeightCm * 0.7), "Section Title Accent Side Rule");
                    sideRule.FillColor = accent;
                    sideRule.OutlineColor = accent;
                    sideRule.OutlineWidthPoints = 0;
                    break;
                case PowerPointTitleAccentStyle.KickerRule:
                    double kickerWidth = Math.Min(3.2, titleWidthCm * 0.28);
                    double kickerLeft = centered ? titleLeftCm + (titleWidthCm - kickerWidth) / 2d : titleLeftCm;
                    PowerPointAutoShape kicker = slide.AddLineCm(kickerLeft, Math.Max(0.65, titleTopCm - 0.22),
                        kickerLeft + kickerWidth, Math.Max(0.65, titleTopCm - 0.22),
                        "Section Title Accent Kicker Rule");
                    kicker.OutlineColor = secondary;
                    kicker.OutlineWidthPoints = 1.4;
                    break;
                case PowerPointTitleAccentStyle.Underline:
                    double underlineWidth = Math.Min(4.2, titleWidthCm * 0.36);
                    double underlineLeft = centered ? titleLeftCm + (titleWidthCm - underlineWidth) / 2d : titleLeftCm;
                    double underlineTop = titleTopCm + Math.Min(0.86, Math.Max(0.52, titleHeightCm * 0.62));
                    PowerPointAutoShape underline = slide.AddRectangleCm(underlineLeft, underlineTop, underlineWidth,
                        0.08, "Section Title Accent Underline");
                    underline.FillColor = accent;
                    underline.OutlineColor = accent;
                    underline.OutlineWidthPoints = 0;
                    break;
            }
        }

        private static void AddChrome(PowerPointSlide slide, PowerPointDesignTheme theme, double slideWidthCm,
            double slideHeightCm, bool dark, PowerPointDesignerSlideOptions options) {
            string text = dark ? theme.AccentLightColor : theme.MutedTextColor;
            string footer = dark ? theme.AccentContrastColor : theme.AccentDarkColor;

            if (!string.IsNullOrWhiteSpace(options.Eyebrow)) {
                AddText(slide, options.Eyebrow!, 1.8, 1.05, 8.0, 0.35, 8, text, theme.BodyFontName);
            }

            if (!string.IsNullOrWhiteSpace(options.FooterLeft)) {
                AddText(slide, options.FooterLeft!, 1.75, slideHeightCm - 1.35, 6.0, 0.45, 16, footer,
                    theme.HeadingFontName, bold: true);
            }

            if (!string.IsNullOrWhiteSpace(options.FooterRight)) {
                PowerPointTextBox right = AddText(slide, options.FooterRight!, slideWidthCm - 5.4,
                    slideHeightCm - 1.35, 4.1, 0.45, 12, footer, theme.HeadingFontName, bold: true);
                RightAlignText(right);
            }

            if (ShouldShowDirectionMotif(options) && !dark) {
                AddDirectionMotif(slide, options, slideWidthCm - 4.9, 1.48, 12, 0.35, theme.AccentColor,
                    flip: true);
            }
        }

        private static bool ShouldShowDirectionMotif(PowerPointDesignerSlideOptions options) {
            return options.ShowDirectionMotif &&
                   ResolveDirectionMotifStyle(options) != PowerPointDirectionMotifStyle.None;
        }

        private static void AddProcessRail(PowerPointSlide slide, PowerPointDesignTheme theme,
            double startXCm, double endXCm, double yCm) {
            PowerPointAutoShape rail = slide.AddLineCm(startXCm, yCm, endXCm, yCm, "Process Rail");
            rail.OutlineColor = theme.AccentLightColor;
            rail.OutlineWidthPoints = 1.1;
        }

        private static void AddProcessConnectors(PowerPointSlide slide, PowerPointDesignTheme theme,
            IReadOnlyList<PowerPointLayoutBox> boxes, double nodeSize, double yCm, double railStartCm,
            double railEndCm, PowerPointProcessConnectorStyle style) {
            if (style == PowerPointProcessConnectorStyle.None) {
                return;
            }

            if (style == PowerPointProcessConnectorStyle.ContinuousRail) {
                AddProcessRail(slide, theme, railStartCm, railEndCm, yCm);
                return;
            }

            for (int i = 0; i < boxes.Count - 1; i++) {
                double start = boxes[i].LeftCm + nodeSize + 0.16;
                double end = boxes[i + 1].LeftCm - 0.16;
                if (end <= start) {
                    continue;
                }

                if (style == PowerPointProcessConnectorStyle.StepDots) {
                    AddProcessConnectorDots(slide, theme, i, start, end, yCm);
                } else {
                    AddProcessConnectorArrow(slide, theme, i, start, end, yCm);
                }
            }
        }

        private static void AddProcessConnectorArrow(PowerPointSlide slide, PowerPointDesignTheme theme, int index,
            double startXCm, double endXCm, double yCm) {
            PowerPointAutoShape connector = slide.AddLineCm(startXCm, yCm, endXCm, yCm,
                "Process Connector " + (index + 1));
            connector.OutlineColor = GetAccent(theme, index);
            connector.OutlineWidthPoints = 1.2;
            connector.SetLineEnds(null, A.LineEndValues.Triangle, A.LineEndWidthValues.Small,
                A.LineEndLengthValues.Small);
        }

        private static void AddProcessConnectorDots(PowerPointSlide slide, PowerPointDesignTheme theme, int index,
            double startXCm, double endXCm, double yCm) {
            const int dotCount = 4;
            double spacing = (endXCm - startXCm) / (dotCount + 1);
            for (int dot = 0; dot < dotCount; dot++) {
                double center = startXCm + spacing * (dot + 1);
                PowerPointAutoShape marker = slide.AddEllipseCm(center - 0.055, yCm - 0.055, 0.11, 0.11,
                    "Process Connector Dot " + (index + 1) + "-" + (dot + 1));
                marker.FillColor = GetAccent(theme, index);
                marker.FillTransparency = 12 + dot * 8;
                marker.OutlineColor = marker.FillColor;
                marker.OutlineWidthPoints = 0;
            }
        }

        private static void AddProcessNode(PowerPointSlide slide, PowerPointDesignTheme theme, int index,
            double leftCm, double topCm, double sizeCm, string number) {
            PowerPointAutoShape halo = slide.AddEllipseCm(leftCm - 0.08, topCm - 0.08,
                sizeCm + 0.16, sizeCm + 0.16, "Process Node Halo " + (index + 1));
            halo.FillColor = theme.AccentLightColor;
            halo.FillTransparency = 78;
            halo.OutlineColor = theme.AccentLightColor;
            halo.OutlineWidthPoints = 0.2;

            PowerPointAutoShape node = slide.AddEllipseCm(leftCm, topCm, sizeCm, sizeCm,
                "Process Node " + (index + 1));
            node.FillColor = theme.AccentDarkColor;
            node.FillTransparency = 8;
            node.OutlineColor = theme.AccentLightColor;
            node.OutlineWidthPoints = 1.2;

            PowerPointTextBox numberBox = AddText(slide, number.TrimEnd('.'), leftCm, topCm - 0.01, sizeCm, sizeCm,
                sizeCm < 1 ? 16 : 20, theme.AccentContrastColor, theme.HeadingFontName, bold: true);
            CenterText(numberBox);
        }

        private static PowerPointDirectionMotifStyle ResolveDirectionMotifStyle(
            PowerPointDesignerSlideOptions options) {
            if (options.DirectionMotifStyle != PowerPointDirectionMotifStyle.Auto) {
                return options.DirectionMotifStyle;
            }

            PowerPointDesignIntent intent = options.DesignIntent;
            if (intent.VisualStyle == PowerPointVisualStyle.Minimal) {
                return PowerPointDirectionMotifStyle.None;
            }
            if (string.IsNullOrWhiteSpace(intent.Seed)) {
                return PowerPointDirectionMotifStyle.Triangles;
            }
            if (intent.Mood == PowerPointDesignMood.Energetic) {
                return PowerPointDirectionMotifStyle.Chevrons;
            }
            if (intent.Mood == PowerPointDesignMood.Editorial) {
                return PowerPointDirectionMotifStyle.Bars;
            }
            if (intent.VisualStyle == PowerPointVisualStyle.Soft) {
                return PowerPointDirectionMotifStyle.Dots;
            }

            return intent.Pick(4, "direction-motif") switch {
                0 => PowerPointDirectionMotifStyle.Triangles,
                1 => PowerPointDirectionMotifStyle.Chevrons,
                2 => PowerPointDirectionMotifStyle.Dots,
                _ => PowerPointDirectionMotifStyle.Bars
            };
        }

        private static void AddDirectionMotif(PowerPointSlide slide, PowerPointDesignerSlideOptions options,
            double leftCm, double topCm, int count, double spacingCm, string color, bool flip = false) {
            PowerPointDirectionMotifStyle style = ResolveDirectionMotifStyle(options);
            if (style == PowerPointDirectionMotifStyle.None) {
                return;
            }

            for (int i = 0; i < count; i++) {
                double left = leftCm + i * spacingCm;
                int transparency = Math.Min(45, i * 3);
                switch (style) {
                    case PowerPointDirectionMotifStyle.Chevrons:
                        AddDirectionChevron(slide, left, topCm, i, color, transparency, flip);
                        break;
                    case PowerPointDirectionMotifStyle.Dots:
                        AddDirectionDot(slide, left, topCm, i, color, transparency);
                        break;
                    case PowerPointDirectionMotifStyle.Bars:
                        AddDirectionBar(slide, left, topCm, i, color, transparency);
                        break;
                    default:
                        AddDirectionTriangle(slide, left, topCm, i, color, transparency, flip);
                        break;
                }
            }
        }

        private static void AddDirectionTriangle(PowerPointSlide slide, double leftCm, double topCm, int index,
            string color, int transparency, bool flip) {
            PowerPointAutoShape arrow = slide.AddShapeCm(A.ShapeTypeValues.Triangle,
                leftCm, topCm, 0.22, 0.24, "Designer Direction " + (index + 1));
            arrow.FillColor = color;
            arrow.FillTransparency = transparency;
            arrow.OutlineColor = color;
            arrow.Rotation = flip ? 270 : 90;
        }

        private static void AddDirectionDot(PowerPointSlide slide, double leftCm, double topCm, int index,
            string color, int transparency) {
            PowerPointAutoShape dot = slide.AddEllipseCm(leftCm, topCm + 0.04, 0.16, 0.16,
                "Designer Direction " + (index + 1));
            dot.FillColor = color;
            dot.FillTransparency = transparency;
            dot.OutlineColor = color;
            dot.OutlineWidthPoints = 0;
        }

        private static void AddDirectionBar(PowerPointSlide slide, double leftCm, double topCm, int index,
            string color, int transparency) {
            PowerPointAutoShape bar = slide.AddRectangleCm(leftCm, topCm + 0.08, 0.24, 0.07,
                "Designer Direction " + (index + 1));
            bar.FillColor = color;
            bar.FillTransparency = transparency;
            bar.OutlineColor = color;
            bar.OutlineWidthPoints = 0;
        }

        private static void AddDirectionChevron(PowerPointSlide slide, double leftCm, double topCm, int index,
            string color, int transparency, bool flip) {
            double tip = flip ? leftCm : leftCm + 0.22;
            double back = flip ? leftCm + 0.22 : leftCm;
            double middleY = topCm + 0.12;

            PowerPointAutoShape upper = slide.AddLineCm(back, topCm + 0.02, tip, middleY,
                "Designer Direction " + (index + 1));
            upper.OutlineColor = color;
            upper.OutlineWidthPoints = Math.Max(0.55, 1.0 - transparency / 100d);

            PowerPointAutoShape lower = slide.AddLineCm(back, topCm + 0.22, tip, middleY,
                "Designer Direction Chevron " + (index + 1) + "B");
            lower.OutlineColor = color;
            lower.OutlineWidthPoints = Math.Max(0.55, 1.0 - transparency / 100d);
        }

        internal static PowerPointTextBox AddText(PowerPointSlide slide, string text, double leftCm, double topCm,
            double widthCm, double heightCm, int fontSize, string color, string fontName, bool bold = false) {
            PowerPointTextBox box = slide.AddTextBoxCm(text, leftCm, topCm, widthCm, heightCm);
            box.SetTextMarginsCm(0, 0, 0, 0);
            box.FontName = fontName;
            box.FontSize = fontSize;
            box.Color = color;
            box.Bold = bold;
            box.TextAutoFit = PowerPointTextAutoFit.Normal;
            return box;
        }

        private static PowerPointPicture? AddPictureIfExists(PowerPointSlide slide, string imagePath,
            double leftCm, double topCm, double widthCm, double heightCm, bool crop) {
            if (!File.Exists(imagePath)) {
                return null;
            }

            PowerPointPicture picture = slide.AddPictureCm(imagePath, leftCm, topCm, widthCm, heightCm);
            if (crop && TryGetImageDimensions(imagePath, out double imageWidth, out double imageHeight)) {
                picture.FitToBox(imageWidth, imageHeight, crop: true);
            }
            return picture;
        }

        private static bool TryGetImageDimensions(string imagePath, out double width, out double height) {
            width = 0;
            height = 0;

            using FileStream stream = File.OpenRead(imagePath);
            if (TryGetPngDimensions(stream, out width, out height)) {
                return true;
            }

            stream.Position = 0;
            return TryGetJpegDimensions(stream, out width, out height);
        }

        private static bool TryGetPngDimensions(Stream stream, out double width, out double height) {
            width = 0;
            height = 0;

            byte[] header = new byte[24];
            if (stream.Read(header, 0, header.Length) != header.Length) {
                return false;
            }

            byte[] signature = { 137, 80, 78, 71, 13, 10, 26, 10 };
            for (int i = 0; i < signature.Length; i++) {
                if (header[i] != signature[i]) {
                    return false;
                }
            }

            width = ReadBigEndianInt32(header, 16);
            height = ReadBigEndianInt32(header, 20);
            return width > 0 && height > 0;
        }

        private static bool TryGetJpegDimensions(Stream stream, out double width, out double height) {
            width = 0;
            height = 0;

            int first = stream.ReadByte();
            int second = stream.ReadByte();
            if (first != 0xFF || second != 0xD8) {
                return false;
            }

            while (stream.Position < stream.Length) {
                int markerPrefix;
                do {
                    markerPrefix = stream.ReadByte();
                    if (markerPrefix < 0) {
                        return false;
                    }
                } while (markerPrefix != 0xFF);

                int marker;
                do {
                    marker = stream.ReadByte();
                    if (marker < 0) {
                        return false;
                    }
                } while (marker == 0xFF);

                if (marker is 0xD8 or 0xD9) {
                    continue;
                }

                int segmentLength = ReadBigEndianUInt16(stream);
                if (segmentLength < 2 || stream.Position + segmentLength - 2 > stream.Length) {
                    return false;
                }

                if (IsJpegStartOfFrame(marker)) {
                    stream.ReadByte();
                    height = ReadBigEndianUInt16(stream);
                    width = ReadBigEndianUInt16(stream);
                    return width > 0 && height > 0;
                }

                stream.Position += segmentLength - 2;
            }

            return false;
        }

        private static bool IsJpegStartOfFrame(int marker) {
            return marker is 0xC0 or 0xC1 or 0xC2 or 0xC3 or 0xC5 or 0xC6 or 0xC7 or
                0xC9 or 0xCA or 0xCB or 0xCD or 0xCE or 0xCF;
        }

        private static int ReadBigEndianInt32(byte[] bytes, int offset) {
            return (bytes[offset] << 24) | (bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3];
        }

        private static int ReadBigEndianUInt16(Stream stream) {
            int high = stream.ReadByte();
            int low = stream.ReadByte();
            if (high < 0 || low < 0) {
                return -1;
            }

            return (high << 8) | low;
        }

        private static void CenterText(PowerPointTextBox textBox) {
            foreach (PowerPointParagraph paragraph in textBox.Paragraphs) {
                paragraph.Alignment = A.TextAlignmentTypeValues.Center;
            }
            textBox.TextVerticalAlignment = A.TextAnchoringTypeValues.Center;
        }

        private static void RightAlignText(PowerPointTextBox textBox) {
            foreach (PowerPointParagraph paragraph in textBox.Paragraphs) {
                paragraph.Alignment = A.TextAlignmentTypeValues.Right;
            }
        }

        private static string GetAccent(PowerPointDesignTheme theme, int index) {
            string[] colors = {
                theme.AccentColor,
                theme.Accent3Color,
                theme.Accent2Color,
                theme.AccentDarkColor,
                theme.WarningColor
            };
            return colors[index % colors.Length];
        }
    }
}
