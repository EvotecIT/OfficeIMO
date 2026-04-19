using System;
using System.Collections.Generic;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Curated set of creative directions for a repeatable presentation scenario.
    /// </summary>
    public sealed class PowerPointDesignRecipe {
        /// <summary>
        ///     Consulting and service portfolio recipe with proof, story, and appendix personalities.
        /// </summary>
        public static PowerPointDesignRecipe ConsultingPortfolio { get; } = new(
            "Consulting Portfolio",
            new[] {
                new PowerPointDesignDirection("Board Story", PowerPointDesignMood.Corporate,
                    PowerPointSlideDensity.Relaxed, PowerPointVisualStyle.Soft, "Georgia", "Aptos",
                    showDirectionMotif: false),
                new PowerPointDesignDirection("Field Proof", PowerPointDesignMood.Energetic,
                    PowerPointSlideDensity.Balanced, PowerPointVisualStyle.Geometric, "Poppins", "Aptos",
                    showDirectionMotif: true),
                new PowerPointDesignDirection("Quiet Appendix", PowerPointDesignMood.Minimal,
                    PowerPointSlideDensity.Relaxed, PowerPointVisualStyle.Minimal, "Aptos Display", "Aptos",
                    showDirectionMotif: false)
            },
            "Project portfolio",
            "Service and case-study decks that need story, evidence, and supporting detail.");

        /// <summary>
        ///     Executive brief recipe for decision decks and board-ready summaries.
        /// </summary>
        public static PowerPointDesignRecipe ExecutiveBrief { get; } = new(
            "Executive Brief",
            new[] {
                new PowerPointDesignDirection("Decision Pack", PowerPointDesignMood.Corporate,
                    PowerPointSlideDensity.Balanced, PowerPointVisualStyle.Soft, "Segoe UI Semibold", "Segoe UI",
                    showDirectionMotif: false),
                new PowerPointDesignDirection("Investment Memo", PowerPointDesignMood.Editorial,
                    PowerPointSlideDensity.Relaxed, PowerPointVisualStyle.Soft, "Georgia", "Aptos",
                    showDirectionMotif: false),
                new PowerPointDesignDirection("Signal Summary", PowerPointDesignMood.Corporate,
                    PowerPointSlideDensity.Compact, PowerPointVisualStyle.Geometric, "Poppins", "Segoe UI",
                    showDirectionMotif: true)
            },
            "Executive summary",
            "Concise decks where hierarchy and restraint matter more than decoration.");

        /// <summary>
        ///     Technical proposal recipe for architecture, rollout, and operations decks.
        /// </summary>
        public static PowerPointDesignRecipe TechnicalProposal { get; } = new(
            "Technical Proposal",
            new[] {
                new PowerPointDesignDirection("Architecture Map", PowerPointDesignMood.Corporate,
                    PowerPointSlideDensity.Balanced, PowerPointVisualStyle.Geometric, "Aptos Display", "Aptos",
                    showDirectionMotif: true),
                new PowerPointDesignDirection("Runbook", PowerPointDesignMood.Minimal,
                    PowerPointSlideDensity.Compact, PowerPointVisualStyle.Minimal, "Segoe UI Semibold", "Segoe UI",
                    showDirectionMotif: false),
                new PowerPointDesignDirection("Delivery Signal", PowerPointDesignMood.Energetic,
                    PowerPointSlideDensity.Compact, PowerPointVisualStyle.Geometric, "Poppins", "Lato",
                    showDirectionMotif: true)
            },
            "Technical proposal",
            "Structured technical decks with room for process, scope, proof, and operational detail.");

        /// <summary>
        ///     Built-in recipes suitable for generating varied deck alternatives.
        /// </summary>
        public static IReadOnlyList<PowerPointDesignRecipe> BuiltIn { get; } = new[] {
            ConsultingPortfolio,
            ExecutiveBrief,
            TechnicalProposal
        };

        /// <summary>
        ///     Creates a reusable design recipe from one or more directions.
        /// </summary>
        public PowerPointDesignRecipe(string name, IEnumerable<PowerPointDesignDirection> directions,
            string? defaultEyebrow = null, string? description = null) {
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("Recipe name cannot be null or empty.", nameof(name));
            }
            if (directions == null) {
                throw new ArgumentNullException(nameof(directions));
            }

            List<PowerPointDesignDirection> directionList = new();
            foreach (PowerPointDesignDirection direction in directions) {
                if (direction == null) {
                    throw new ArgumentException("Design recipe directions cannot contain null entries.", nameof(directions));
                }
                directionList.Add(direction);
            }
            if (directionList.Count == 0) {
                throw new ArgumentException("At least one design direction is required.", nameof(directions));
            }

            Name = name;
            Directions = directionList.AsReadOnly();
            DefaultEyebrow = defaultEyebrow;
            Description = description;
        }

        /// <summary>
        ///     Recipe display name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        ///     Optional short explanation of where the recipe fits.
        /// </summary>
        public string? Description { get; }

        /// <summary>
        ///     Optional eyebrow applied when the caller does not supply one.
        /// </summary>
        public string? DefaultEyebrow { get; }

        /// <summary>
        ///     Curated creative directions used by this recipe.
        /// </summary>
        public IReadOnlyList<PowerPointDesignDirection> Directions { get; }

        internal PowerPointDesignDirection DirectionAt(int index) {
            return Directions[index % Directions.Count];
        }
    }
}
