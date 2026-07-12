using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointDesignExtensions {
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
    }
}
