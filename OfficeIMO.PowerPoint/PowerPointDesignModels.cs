using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Broad visual direction for designer compositions.
    /// </summary>
    public enum PowerPointDesignMood {
        /// <summary>Clean business presentation style.</summary>
        Corporate,
        /// <summary>More editorial spacing and emphasis.</summary>
        Editorial,
        /// <summary>More vivid accent treatment.</summary>
        Energetic,
        /// <summary>Reduced decoration and quieter surfaces.</summary>
        Minimal
    }

    /// <summary>
    ///     Preferred amount of content breathing room on generated slides.
    /// </summary>
    public enum PowerPointSlideDensity {
        /// <summary>Fits more content into each slide region.</summary>
        Compact,
        /// <summary>Default balance between content and whitespace.</summary>
        Balanced,
        /// <summary>Adds more breathing room around content.</summary>
        Relaxed
    }

    /// <summary>
    ///     Visual language used by designer primitives.
    /// </summary>
    public enum PowerPointVisualStyle {
        /// <summary>Uses angular planes and structural accents.</summary>
        Geometric,
        /// <summary>Uses softer panels and lighter accents.</summary>
        Soft,
        /// <summary>Uses the fewest decorative elements.</summary>
        Minimal
    }

    /// <summary>
    ///     Controls how far generated alternatives should move from the selected recipe or preferred direction.
    /// </summary>
    public enum PowerPointDesignVariety {
        /// <summary>Stay close to the highest-ranked direction and vary mostly by palette and seed.</summary>
        Focused,
        /// <summary>Use the selected recipe or custom direction set as supplied.</summary>
        Balanced,
        /// <summary>Extend recipe alternatives with broader built-in directions for more creative distance.</summary>
        Exploratory
    }

    /// <summary>
    ///     Supporting palette strategy used to move generated alternatives away from a single house style.
    /// </summary>
    public enum PowerPointPaletteStyle {
        /// <summary>Choose a deterministic supporting palette from the design seed.</summary>
        Auto,
        /// <summary>Use nearby hues for a calm brand-led palette.</summary>
        Analogous,
        /// <summary>Use a contrasting hue for stronger visual separation.</summary>
        Complementary,
        /// <summary>Use two contrasting hues around the brand accent complement.</summary>
        SplitComplementary,
        /// <summary>Use mostly tints and shades of the brand accent.</summary>
        Monochrome,
        /// <summary>Pair the brand accent with warmer neutral surfaces and markers.</summary>
        WarmNeutral,
        /// <summary>Pair the brand accent with cooler neutral surfaces and markers.</summary>
        CoolNeutral
    }

    /// <summary>
    ///     Controls how Auto slide variants balance content fit against visual variety.
    /// </summary>
    public enum PowerPointAutoLayoutStrategy {
        /// <summary>Favor layouts that fit the supplied content, then use the design seed for ties.</summary>
        ContentFirst,
        /// <summary>Favor deterministic design-seed variety unless content would become dense or less readable.</summary>
        DesignFirst,
        /// <summary>Favor compact, scannable variants for denser business decks.</summary>
        Compact,
        /// <summary>Favor variants with stronger visual panels, proof areas, or hero framing when content allows it.</summary>
        VisualFirst
    }

    /// <summary>
    ///     Section/title slide layout variants. Auto uses the design intent seed to pick a stable variant.
    /// </summary>
    public enum PowerPointSectionLayoutVariant {
        /// <summary>Choose a deterministic variant from the design intent.</summary>
        Auto,
        /// <summary>Full-bleed dark title slide with geometric planes.</summary>
        GeometricCover,
        /// <summary>Editorial light title slide with a strong accent rail.</summary>
        EditorialRail,
        /// <summary>Poster-style dark title slide with a large centered title area.</summary>
        Poster
    }

    /// <summary>
    ///     Case-study slide layout variants. Auto uses the design intent seed to pick a stable variant.
    /// </summary>
    public enum PowerPointCaseStudyLayoutVariant {
        /// <summary>Choose a deterministic variant from the design intent.</summary>
        Auto,
        /// <summary>Summary columns with a strong visual band.</summary>
        VisualBand,
        /// <summary>Editorial split with narrative cards and a right-side visual panel.</summary>
        EditorialSplit,
        /// <summary>Large visual frame paired with a compact narrative stack.</summary>
        VisualHero
    }

    /// <summary>
    ///     Process slide layout variants. Auto uses the design intent seed to pick a stable variant.
    /// </summary>
    public enum PowerPointProcessLayoutVariant {
        /// <summary>Choose a deterministic variant from the design intent.</summary>
        Auto,
        /// <summary>Horizontal rail with numbered nodes.</summary>
        Rail,
        /// <summary>Separate numbered columns without a rail.</summary>
        NumberedColumns
    }

    /// <summary>
    ///     Card grid layout variants. Auto uses the design intent seed to pick a stable variant.
    /// </summary>
    public enum PowerPointCardGridLayoutVariant {
        /// <summary>Choose a deterministic variant from the design intent.</summary>
        Auto,
        /// <summary>Cards with a horizontal accent bar.</summary>
        AccentTop,
        /// <summary>Softer cards with a vertical accent strip.</summary>
        SoftTiles
    }

    /// <summary>
    ///     Metric strip surface variants for raw designer compositions.
    /// </summary>
    public enum PowerPointMetricStripVariant {
        /// <summary>Choose a deterministic surface from the active design intent.</summary>
        Auto,
        /// <summary>Single accent band behind the metrics.</summary>
        SolidBand,
        /// <summary>Separate metric tiles with individual accent surfaces.</summary>
        SeparatedTiles,
        /// <summary>Quiet text metrics with accent underlines and no filled panel.</summary>
        Underlined
    }

    /// <summary>
    ///     Visual frame placeholder variants for raw designer compositions.
    /// </summary>
    public enum PowerPointVisualFrameVariant {
        /// <summary>Choose a deterministic visual placeholder from the active design intent.</summary>
        Auto,
        /// <summary>Dashboard-like placeholder with panels and content lines.</summary>
        Dashboard,
        /// <summary>Layered collage placeholder with overlapping proof tiles.</summary>
        Collage,
        /// <summary>Diagram placeholder with editable nodes and connectors.</summary>
        Diagram
    }

    /// <summary>
    ///     Logo/certification wall layout variants. Auto uses the design intent seed to pick a stable variant.
    /// </summary>
    public enum PowerPointLogoWallLayoutVariant {
        /// <summary>Choose a deterministic variant from the design intent.</summary>
        Auto,
        /// <summary>Balanced grid of partner, certification, or product marks.</summary>
        LogoMosaic,
        /// <summary>Logo grid paired with a larger certificate or proof frame.</summary>
        CertificateFeature
    }

    /// <summary>
    ///     Coverage/location slide layout variants. Auto uses the design intent seed to pick a stable variant.
    /// </summary>
    public enum PowerPointCoverageLayoutVariant {
        /// <summary>Choose a deterministic variant from the design intent.</summary>
        Auto,
        /// <summary>Large map-like board with pins and a compact location strip.</summary>
        PinBoard,
        /// <summary>Location list paired with a focused map-like panel.</summary>
        ListMap
    }

    /// <summary>
    ///     Capability/content slide layout variants. Auto uses the design intent seed to pick a stable variant.
    /// </summary>
    public enum PowerPointCapabilityLayoutVariant {
        /// <summary>Choose a deterministic variant from the design intent.</summary>
        Auto,
        /// <summary>Narrative sections on the left, visual support on the right.</summary>
        TextVisual,
        /// <summary>Visual support on the left, narrative sections on the right.</summary>
        VisualText,
        /// <summary>Full-width stacked section panels with optional metrics below.</summary>
        Stacked
    }

    /// <summary>
    ///     Visual support type for capability/content slides.
    /// </summary>
    public enum PowerPointCapabilityVisualKind {
        /// <summary>A polished image frame or editable placeholder.</summary>
        VisualFrame,
        /// <summary>An editable coverage map-like panel with normalized pins.</summary>
        CoverageMap,
        /// <summary>An editable logo/proof wall.</summary>
        LogoWall
    }

    /// <summary>
    ///     Describes the intended feel of generated designer slides without requiring manual placement.
    /// </summary>
    public sealed class PowerPointDesignIntent {
        /// <summary>
        ///     Creates a neutral, reusable design intent.
        /// </summary>
        public PowerPointDesignIntent() {
        }

        /// <summary>
        ///     Creates a design intent from a broad deck mood.
        /// </summary>
        public static PowerPointDesignIntent FromMood(PowerPointDesignMood mood, string? seed = null) {
            PowerPointDesignIntent intent = new() {
                Seed = seed,
                Mood = mood
            };

            switch (mood) {
                case PowerPointDesignMood.Editorial:
                    intent.Density = PowerPointSlideDensity.Relaxed;
                    intent.VisualStyle = PowerPointVisualStyle.Soft;
                    break;
                case PowerPointDesignMood.Energetic:
                    intent.Density = PowerPointSlideDensity.Balanced;
                    intent.VisualStyle = PowerPointVisualStyle.Geometric;
                    break;
                case PowerPointDesignMood.Minimal:
                    intent.Density = PowerPointSlideDensity.Relaxed;
                    intent.VisualStyle = PowerPointVisualStyle.Minimal;
                    break;
                default:
                    intent.Density = PowerPointSlideDensity.Balanced;
                    intent.VisualStyle = PowerPointVisualStyle.Geometric;
                    break;
            }

            return intent;
        }

        /// <summary>
        ///     Creates a copy of this intent.
        /// </summary>
        public PowerPointDesignIntent Clone() {
            return new PowerPointDesignIntent {
                Seed = Seed,
                Mood = Mood,
                Density = Density,
                VisualStyle = VisualStyle,
                LayoutStrategy = LayoutStrategy
            };
        }

        /// <summary>
        ///     Optional stable seed used to choose deterministic variants.
        /// </summary>
        public string? Seed { get; set; }

        /// <summary>
        ///     Broad visual mood.
        /// </summary>
        public PowerPointDesignMood Mood { get; set; } = PowerPointDesignMood.Corporate;

        /// <summary>
        ///     Desired content density.
        /// </summary>
        public PowerPointSlideDensity Density { get; set; } = PowerPointSlideDensity.Balanced;

        /// <summary>
        ///     Preferred primitive visual style.
        /// </summary>
        public PowerPointVisualStyle VisualStyle { get; set; } = PowerPointVisualStyle.Geometric;

        /// <summary>
        ///     Strategy used by Auto slide variants.
        /// </summary>
        public PowerPointAutoLayoutStrategy LayoutStrategy { get; set; } = PowerPointAutoLayoutStrategy.ContentFirst;

        internal int Pick(int choices, string salt) {
            if (choices <= 0) {
                throw new ArgumentOutOfRangeException(nameof(choices));
            }

            string value = LayoutStrategy == PowerPointAutoLayoutStrategy.ContentFirst
                ? string.Join("|", Seed ?? string.Empty, Mood, Density, VisualStyle, salt)
                : string.Join("|", Seed ?? string.Empty, Mood, Density, VisualStyle, LayoutStrategy, salt);
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

    /// <summary>
    ///     Common chrome options for high-level designer slides.
    /// </summary>
    public class PowerPointDesignerSlideOptions {
        /// <summary>
        ///     Small label placed near the top of the slide.
        /// </summary>
        public string? Eyebrow { get; set; }

        /// <summary>
        ///     Left footer text, often a logo wordmark or product name.
        /// </summary>
        public string? FooterLeft { get; set; } = "OfficeIMO";

        /// <summary>
        ///     Right footer text.
        /// </summary>
        public string? FooterRight { get; set; }

        /// <summary>
        ///     Adds a row of editable triangle markers for movement and visual rhythm.
        /// </summary>
        public bool ShowDirectionMotif { get; set; } = true;

        /// <summary>
        ///     Design intent used for deterministic visual variation.
        /// </summary>
        public PowerPointDesignIntent DesignIntent { get; set; } = new();

        /// <summary>
        ///     Section slide layout variant. Used by section/title slide helpers.
        /// </summary>
        public PowerPointSectionLayoutVariant SectionVariant { get; set; } = PowerPointSectionLayoutVariant.Auto;
    }

    /// <summary>
    ///     Options for a case-study summary slide.
    /// </summary>
    public sealed class PowerPointCaseStudySlideOptions : PowerPointDesignerSlideOptions {
        /// <summary>
        ///     Case-study layout variant. Auto uses the design intent seed.
        /// </summary>
        public PowerPointCaseStudyLayoutVariant Variant { get; set; } = PowerPointCaseStudyLayoutVariant.Auto;

        /// <summary>
        ///     Optional supporting image displayed in the bottom visual area.
        /// </summary>
        public string? VisualImagePath { get; set; }

        /// <summary>
        ///     Optional person or product cutout displayed in the bottom brand band.
        /// </summary>
        public string? PersonImagePath { get; set; }

        /// <summary>
        ///     Brand text displayed inside the bottom band when no logo asset is provided.
        /// </summary>
        public string? BrandText { get; set; } = "OfficeIMO";

        /// <summary>
        ///     Label shown in the bottom band.
        /// </summary>
        public string? BandLabel { get; set; } = "Project portfolio";

        /// <summary>
        ///     Optional tags displayed as pills along the bottom band.
        /// </summary>
        public IList<string> Tags { get; } = new List<string>();
    }

    /// <summary>
    ///     Options for a card-grid slide.
    /// </summary>
    public sealed class PowerPointCardGridSlideOptions : PowerPointDesignerSlideOptions {
        /// <summary>
        ///     Maximum cards per row before the layout wraps.
        /// </summary>
        public int MaxColumns { get; set; } = 4;

        /// <summary>
        ///     Card layout variant. Auto uses the design intent seed.
        /// </summary>
        public PowerPointCardGridLayoutVariant Variant { get; set; } = PowerPointCardGridLayoutVariant.Auto;

        /// <summary>
        ///     Optional supporting text block displayed below the cards.
        /// </summary>
        public string? SupportingText { get; set; }
    }

    /// <summary>
    ///     Options for process/timeline slides.
    /// </summary>
    public sealed class PowerPointProcessSlideOptions : PowerPointDesignerSlideOptions {
        /// <summary>
        ///     Adds translucent diagonal planes behind the timeline.
        /// </summary>
        public bool ShowDiagonalPlanes { get; set; } = true;

        /// <summary>
        ///     Process layout variant. Auto uses the design intent seed.
        /// </summary>
        public PowerPointProcessLayoutVariant Variant { get; set; } = PowerPointProcessLayoutVariant.Auto;
    }

    /// <summary>
    ///     Options for logo, partner, or certification wall slides.
    /// </summary>
    public sealed class PowerPointLogoWallSlideOptions : PowerPointDesignerSlideOptions {
        /// <summary>
        ///     Logo wall layout variant. Auto uses the design intent seed.
        /// </summary>
        public PowerPointLogoWallLayoutVariant Variant { get; set; } = PowerPointLogoWallLayoutVariant.Auto;

        /// <summary>
        ///     Maximum logo tiles per row before the layout wraps.
        /// </summary>
        public int MaxColumns { get; set; } = 5;

        /// <summary>
        ///     Optional supporting text displayed below or beside the logo wall.
        /// </summary>
        public string? SupportingText { get; set; }

        /// <summary>
        ///     Optional image displayed as a featured proof/certificate visual.
        /// </summary>
        public string? FeaturedImagePath { get; set; }

        /// <summary>
        ///     Optional caption shown near the featured proof/certificate visual.
        /// </summary>
        public string? FeatureTitle { get; set; }
    }

    /// <summary>
    ///     Options for coverage and location slides.
    /// </summary>
    public sealed class PowerPointCoverageSlideOptions : PowerPointDesignerSlideOptions {
        /// <summary>
        ///     Coverage layout variant. Auto uses the design intent seed.
        /// </summary>
        public PowerPointCoverageLayoutVariant Variant { get; set; } = PowerPointCoverageLayoutVariant.Auto;

        /// <summary>
        ///     Optional supporting text displayed with the location list or location strip.
        /// </summary>
        public string? SupportingText { get; set; }

        /// <summary>
        ///     Optional label placed on the editable map-like surface.
        /// </summary>
        public string? MapLabel { get; set; }
    }

    /// <summary>
    ///     Options for capability/content slides.
    /// </summary>
    public sealed class PowerPointCapabilitySlideOptions : PowerPointDesignerSlideOptions {
        /// <summary>
        ///     Capability layout variant. Auto uses the design intent seed.
        /// </summary>
        public PowerPointCapabilityLayoutVariant Variant { get; set; } = PowerPointCapabilityLayoutVariant.Auto;

        /// <summary>
        ///     Visual support type displayed beside or below the narrative sections.
        /// </summary>
        public PowerPointCapabilityVisualKind VisualKind { get; set; } = PowerPointCapabilityVisualKind.VisualFrame;

        /// <summary>
        ///     Optional image path for the visual frame.
        /// </summary>
        public string? VisualImagePath { get; set; }

        /// <summary>
        ///     Optional label displayed with the visual support area.
        /// </summary>
        public string? VisualLabel { get; set; }

        /// <summary>
        ///     Optional logo/proof items when VisualKind is LogoWall.
        /// </summary>
        public IList<PowerPointLogoItem> Logos { get; } = new List<PowerPointLogoItem>();

        /// <summary>
        ///     Optional coverage locations when VisualKind is CoverageMap.
        /// </summary>
        public IList<PowerPointCoverageLocation> Locations { get; } = new List<PowerPointCoverageLocation>();

        /// <summary>
        ///     Optional metrics displayed as a supporting strip.
        /// </summary>
        public IList<PowerPointMetric> Metrics { get; } = new List<PowerPointMetric>();
    }

    /// <summary>
    ///     A text section in a case-study slide.
    /// </summary>
    public sealed class PowerPointCaseStudySection {
        /// <summary>
        ///     Creates a case-study text section.
        /// </summary>
        public PowerPointCaseStudySection(string heading, string body) {
            Heading = heading ?? throw new ArgumentNullException(nameof(heading));
            Body = body ?? throw new ArgumentNullException(nameof(body));
        }

        /// <summary>
        ///     Section heading.
        /// </summary>
        public string Heading { get; }

        /// <summary>
        ///     Section body text.
        /// </summary>
        public string Body { get; }
    }

    /// <summary>
    ///     A metric displayed prominently on a designer slide.
    /// </summary>
    public sealed class PowerPointMetric {
        /// <summary>
        ///     Creates a metric with a value and label.
        /// </summary>
        public PowerPointMetric(string value, string label) {
            Value = value ?? throw new ArgumentNullException(nameof(value));
            Label = label ?? throw new ArgumentNullException(nameof(label));
        }

        /// <summary>
        ///     Prominent metric value.
        /// </summary>
        public string Value { get; }

        /// <summary>
        ///     Metric label or caption.
        /// </summary>
        public string Label { get; }
    }

    /// <summary>
    ///     A card with a title and optional bullet items.
    /// </summary>
    public sealed class PowerPointCardContent {
        /// <summary>
        ///     Creates a designer card.
        /// </summary>
        public PowerPointCardContent(string title, IEnumerable<string>? items = null, string? accentColor = null) {
            Title = title ?? throw new ArgumentNullException(nameof(title));
            Items = (items ?? Enumerable.Empty<string>()).Where(item => item != null).ToList();
            AccentColor = accentColor;
        }

        /// <summary>
        ///     Card title.
        /// </summary>
        public string Title { get; }

        /// <summary>
        ///     Bullet items displayed in the card.
        /// </summary>
        public IReadOnlyList<string> Items { get; }

        /// <summary>
        ///     Optional accent color override.
        /// </summary>
        public string? AccentColor { get; }
    }

    /// <summary>
    ///     A single step in a process/timeline slide.
    /// </summary>
    public sealed class PowerPointProcessStep {
        /// <summary>
        ///     Creates a process step.
        /// </summary>
        public PowerPointProcessStep(string title, string body, string? number = null) {
            Title = title ?? throw new ArgumentNullException(nameof(title));
            Body = body ?? throw new ArgumentNullException(nameof(body));
            Number = number;
        }

        /// <summary>
        ///     Optional displayed step number. When omitted, the composition assigns one.
        /// </summary>
        public string? Number { get; }

        /// <summary>
        ///     Step title.
        /// </summary>
        public string Title { get; }

        /// <summary>
        ///     Step body text.
        /// </summary>
        public string Body { get; }
    }

    /// <summary>
    ///     A logo, partner, product, or certification mark for a logo wall.
    /// </summary>
    public sealed class PowerPointLogoItem {
        /// <summary>
        ///     Creates a logo wall item. When ImagePath is omitted, the name is rendered as editable text.
        /// </summary>
        public PowerPointLogoItem(string name, string? subtitle = null, string? imagePath = null,
            string? accentColor = null) {
            Name = name ?? throw new ArgumentNullException(nameof(name));
            Subtitle = subtitle;
            ImagePath = imagePath;
            AccentColor = accentColor;
        }

        /// <summary>
        ///     Logo or certification name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        ///     Optional short descriptor.
        /// </summary>
        public string? Subtitle { get; }

        /// <summary>
        ///     Optional image path for an actual logo/certification mark.
        /// </summary>
        public string? ImagePath { get; }

        /// <summary>
        ///     Optional accent color override for the tile.
        /// </summary>
        public string? AccentColor { get; }
    }

    /// <summary>
    ///     A location marker for coverage and map-like slides.
    /// </summary>
    public sealed class PowerPointCoverageLocation {
        /// <summary>
        ///     Creates a coverage location. X and Y are normalized positions from 0 to 1 inside the map panel.
        /// </summary>
        public PowerPointCoverageLocation(string name, double x, double y, string? detail = null) {
            Name = name ?? throw new ArgumentNullException(nameof(name));
            X = x;
            Y = y;
            Detail = detail;
        }

        /// <summary>
        ///     Location name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        ///     Normalized horizontal position from 0 to 1 inside the map panel.
        /// </summary>
        public double X { get; }

        /// <summary>
        ///     Normalized vertical position from 0 to 1 inside the map panel.
        /// </summary>
        public double Y { get; }

        /// <summary>
        ///     Optional location detail.
        /// </summary>
        public string? Detail { get; }
    }

    /// <summary>
    ///     A narrative section for capability/content slides.
    /// </summary>
    public sealed class PowerPointCapabilitySection {
        /// <summary>
        ///     Creates a capability section with optional body text and bullet items.
        /// </summary>
        public PowerPointCapabilitySection(string heading, string? body = null, IEnumerable<string>? items = null,
            string? accentColor = null) {
            Heading = heading ?? throw new ArgumentNullException(nameof(heading));
            Body = body;
            Items = (items ?? Enumerable.Empty<string>()).Where(item => item != null).ToList();
            AccentColor = accentColor;
        }

        /// <summary>
        ///     Section heading.
        /// </summary>
        public string Heading { get; }

        /// <summary>
        ///     Optional body text.
        /// </summary>
        public string? Body { get; }

        /// <summary>
        ///     Optional bullet items.
        /// </summary>
        public IReadOnlyList<string> Items { get; }

        /// <summary>
        ///     Optional accent color override.
        /// </summary>
        public string? AccentColor { get; }
    }
}
