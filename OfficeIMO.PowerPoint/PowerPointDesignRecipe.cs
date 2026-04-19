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
            "Service and case-study decks that need story, evidence, and supporting detail.",
            new[] { "consulting", "portfolio", "case study", "service", "proof" });

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
            "Concise decks where hierarchy and restraint matter more than decoration.",
            new[] { "executive", "board", "brief", "summary", "decision" });

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
            "Structured technical decks with room for process, scope, proof, and operational detail.",
            new[] { "technical", "proposal", "architecture", "rollout", "operations" });

        /// <summary>
        ///     Transformation roadmap recipe for change programs, milestones, and phased plans.
        /// </summary>
        public static PowerPointDesignRecipe TransformationRoadmap { get; } = new(
            "Transformation Roadmap",
            new[] {
                new PowerPointDesignDirection("North Star", PowerPointDesignMood.Editorial,
                    PowerPointSlideDensity.Relaxed, PowerPointVisualStyle.Soft, "Georgia", "Aptos",
                    showDirectionMotif: false),
                new PowerPointDesignDirection("Momentum Map", PowerPointDesignMood.Energetic,
                    PowerPointSlideDensity.Balanced, PowerPointVisualStyle.Geometric, "Poppins", "Lato",
                    showDirectionMotif: true),
                new PowerPointDesignDirection("Operating Plan", PowerPointDesignMood.Corporate,
                    PowerPointSlideDensity.Compact, PowerPointVisualStyle.Minimal, "Segoe UI Semibold", "Segoe UI",
                    showDirectionMotif: false)
            },
            "Roadmap",
            "Change and transformation decks that need phases, decisions, and implementation rhythm.",
            new[] { "roadmap", "transformation", "change", "program", "journey", "milestone" });

        /// <summary>
        ///     Built-in recipes suitable for generating varied deck alternatives.
        /// </summary>
        public static IReadOnlyList<PowerPointDesignRecipe> BuiltIn { get; } = new[] {
            ConsultingPortfolio,
            ExecutiveBrief,
            TechnicalProposal,
            TransformationRoadmap
        };

        /// <summary>
        ///     Creates a reusable design recipe from one or more directions.
        /// </summary>
        public PowerPointDesignRecipe(string name, IEnumerable<PowerPointDesignDirection> directions,
            string? defaultEyebrow = null, string? description = null, IEnumerable<string>? keywords = null) {
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
            Keywords = NormalizeKeywords(keywords);
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
        ///     Optional purpose keywords used when selecting built-in recipes from plain text.
        /// </summary>
        public IReadOnlyList<string> Keywords { get; }

        /// <summary>
        ///     Curated creative directions used by this recipe.
        /// </summary>
        public IReadOnlyList<PowerPointDesignDirection> Directions { get; }

        /// <summary>
        ///     Creates deterministic deck design alternatives from this recipe.
        /// </summary>
        public IReadOnlyList<PowerPointDeckDesign> CreateAlternativesFromBrand(string accentColor, string seed,
            int count = 0, string? name = null, string? eyebrow = null, string? footerLeft = null,
            string? footerRight = null, string? headingFontName = null, string? bodyFontName = null) {
            return PowerPointDeckDesign.CreateAlternativesFromBrand(accentColor, seed, this, count, name, eyebrow,
                footerLeft, footerRight, headingFontName, bodyFontName);
        }

        /// <summary>
        ///     Finds a built-in recipe that matches a plain-language purpose such as "executive brief".
        /// </summary>
        public static PowerPointDesignRecipe? FindBuiltIn(string purpose) {
            if (string.IsNullOrWhiteSpace(purpose)) {
                return null;
            }

            foreach (PowerPointDesignRecipe recipe in BuiltIn) {
                if (recipe.Matches(purpose)) {
                    return recipe;
                }
            }

            return null;
        }

        /// <summary>
        ///     Determines whether this recipe matches a plain-language purpose.
        /// </summary>
        public bool Matches(string purpose) {
            if (string.IsNullOrWhiteSpace(purpose)) {
                return false;
            }

            if (Contains(Name, purpose) || Contains(purpose, Name)) {
                return true;
            }

            if (!string.IsNullOrWhiteSpace(Description) &&
                (Contains(Description!, purpose) || Contains(purpose, Description!))) {
                return true;
            }

            foreach (string keyword in Keywords) {
                if (Contains(purpose, keyword) || Contains(keyword, purpose)) {
                    return true;
                }
            }

            foreach (PowerPointDesignDirection direction in Directions) {
                if (Contains(purpose, direction.Name) || Contains(direction.Name, purpose)) {
                    return true;
                }
            }

            return false;
        }

        internal PowerPointDesignDirection DirectionAt(int index) {
            return Directions[index % Directions.Count];
        }

        private static IReadOnlyList<string> NormalizeKeywords(IEnumerable<string>? keywords) {
            List<string> normalized = new();
            if (keywords == null) {
                return normalized.AsReadOnly();
            }

            foreach (string keyword in keywords) {
                if (!string.IsNullOrWhiteSpace(keyword)) {
                    normalized.Add(keyword.Trim());
                }
            }

            return normalized.AsReadOnly();
        }

        private static bool Contains(string source, string value) {
            return source.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
