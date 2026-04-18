using System;
using System.Collections.Generic;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Named creative direction used to generate distinct deck personalities from one brand accent.
    /// </summary>
    public sealed class PowerPointDesignDirection {
        /// <summary>
        ///     Clean consulting-style deck direction with geometric structure.
        /// </summary>
        public static PowerPointDesignDirection Structured { get; } = new(
            "Structured",
            PowerPointDesignMood.Corporate,
            PowerPointSlideDensity.Balanced,
            PowerPointVisualStyle.Geometric,
            "Poppins",
            "Lato",
            showDirectionMotif: true);

        /// <summary>
        ///     Spacious editorial direction with softer surfaces and calmer typography.
        /// </summary>
        public static PowerPointDesignDirection Editorial { get; } = new(
            "Editorial",
            PowerPointDesignMood.Editorial,
            PowerPointSlideDensity.Relaxed,
            PowerPointVisualStyle.Soft,
            "Aptos Display",
            "Aptos",
            showDirectionMotif: true);

        /// <summary>
        ///     Quiet direction for understated decks with very little decoration.
        /// </summary>
        public static PowerPointDesignDirection Quiet { get; } = new(
            "Quiet",
            PowerPointDesignMood.Minimal,
            PowerPointSlideDensity.Relaxed,
            PowerPointVisualStyle.Minimal,
            "Aptos Display",
            "Aptos",
            showDirectionMotif: false);

        /// <summary>
        ///     Higher-energy direction with stronger accents and compact content rhythm.
        /// </summary>
        public static PowerPointDesignDirection Signal { get; } = new(
            "Signal",
            PowerPointDesignMood.Energetic,
            PowerPointSlideDensity.Compact,
            PowerPointVisualStyle.Geometric,
            "Poppins",
            "Aptos",
            showDirectionMotif: true);

        /// <summary>
        ///     Simple executive direction using system-safe typography and balanced spacing.
        /// </summary>
        public static PowerPointDesignDirection Executive { get; } = new(
            "Executive",
            PowerPointDesignMood.Corporate,
            PowerPointSlideDensity.Balanced,
            PowerPointVisualStyle.Soft,
            "Segoe UI Semibold",
            "Segoe UI",
            showDirectionMotif: false);

        /// <summary>
        ///     Built-in directions used by deck alternatives.
        /// </summary>
        public static IReadOnlyList<PowerPointDesignDirection> BuiltIn { get; } = new[] {
            Structured,
            Editorial,
            Quiet,
            Signal,
            Executive
        };

        /// <summary>
        ///     Creates a reusable direction definition.
        /// </summary>
        public PowerPointDesignDirection(string name, PowerPointDesignMood mood,
            PowerPointSlideDensity density, PowerPointVisualStyle visualStyle,
            string headingFontName, string bodyFontName, bool showDirectionMotif = true) {
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("Direction name cannot be null or empty.", nameof(name));
            }
            if (string.IsNullOrWhiteSpace(headingFontName)) {
                throw new ArgumentException("Heading font cannot be null or empty.", nameof(headingFontName));
            }
            if (string.IsNullOrWhiteSpace(bodyFontName)) {
                throw new ArgumentException("Body font cannot be null or empty.", nameof(bodyFontName));
            }

            Name = name;
            Mood = mood;
            Density = density;
            VisualStyle = visualStyle;
            HeadingFontName = headingFontName;
            BodyFontName = bodyFontName;
            ShowDirectionMotif = showDirectionMotif;
        }

        /// <summary>
        ///     Direction display name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        ///     Broad visual mood.
        /// </summary>
        public PowerPointDesignMood Mood { get; }

        /// <summary>
        ///     Preferred content density.
        /// </summary>
        public PowerPointSlideDensity Density { get; }

        /// <summary>
        ///     Preferred primitive visual style.
        /// </summary>
        public PowerPointVisualStyle VisualStyle { get; }

        /// <summary>
        ///     Heading font for generated text.
        /// </summary>
        public string HeadingFontName { get; }

        /// <summary>
        ///     Body font for generated text.
        /// </summary>
        public string BodyFontName { get; }

        /// <summary>
        ///     Whether this direction uses repeated direction markers by default.
        /// </summary>
        public bool ShowDirectionMotif { get; }
    }
}
