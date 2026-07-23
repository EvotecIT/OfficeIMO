using System;

namespace OfficeIMO.PowerPoint {
    /// <summary>Built-in accessibility policy profiles.</summary>
    public enum PowerPointAccessibilityPolicyProfile {
        /// <summary>Practical generated-deck policy suitable for ordinary CI gates.</summary>
        Default,
        /// <summary>Stricter policy requiring document metadata and complete explicit semantics.</summary>
        Strict
    }

    /// <summary>Options controlling PowerPoint accessibility inspection and CI enforcement.</summary>
    public sealed class PowerPointAccessibilityOptions {
        private double _minimumTextContrastRatio = 4.5D;
        private double _minimumLargeTextContrastRatio = 3D;
        private int _maximumShapeCount = 10000;
        private int _maximumGroupDepth = 128;

        /// <summary>Maximum number of shapes inspected on one slide.</summary>
        public int MaximumShapeCount {
            get => _maximumShapeCount;
            set => _maximumShapeCount = ValidatePositive(value, nameof(MaximumShapeCount));
        }

        /// <summary>Maximum supported nesting depth for grouped shapes.</summary>
        public int MaximumGroupDepth {
            get => _maximumGroupDepth;
            set => _maximumGroupDepth = ValidatePositive(value, nameof(MaximumGroupDepth));
        }

        /// <summary>Selected built-in profile. Use <see cref="ForProfile"/> to apply its defaults.</summary>
        public PowerPointAccessibilityPolicyProfile Profile { get; set; } = PowerPointAccessibilityPolicyProfile.Default;

        /// <summary>Whether the package must define a document title.</summary>
        public bool RequireDocumentTitle { get; set; }

        /// <summary>Whether every visible slide must expose a recognizable title.</summary>
        public bool RequireSlideTitles { get; set; } = true;

        /// <summary>Whether hidden slides should be included in accessibility inspection. Defaults to false.</summary>
        public bool IncludeHiddenSlides { get; set; }

        /// <summary>Whether informative visual shapes must have concise titles in addition to descriptions.</summary>
        public bool RequireShapeTitles { get; set; }

        /// <summary>Whether informative pictures, charts, tables, media, and SmartArt require descriptions.</summary>
        public bool RequireAlternativeText { get; set; } = true;

        /// <summary>Whether text-bearing shapes require an explicit language tag.</summary>
        public bool RequireLanguage { get; set; } = true;

        /// <summary>Whether tables require first-row header semantics.</summary>
        public bool RequireTableHeaders { get; set; } = true;

        /// <summary>Whether visible text with resolvable foreground/background colors is checked for contrast.</summary>
        public bool CheckContrast { get; set; } = true;

        /// <summary>Whether hyperlinks require meaningful visible labels.</summary>
        public bool CheckMeaningfulLinks { get; set; } = true;

        /// <summary>Whether multi-series charts without a data summary are treated as color-only meaning risks.</summary>
        public bool CheckColorOnlyMeaning { get; set; } = true;

        /// <summary>Minimum normal-text contrast ratio.</summary>
        public double MinimumTextContrastRatio {
            get => _minimumTextContrastRatio;
            set => _minimumTextContrastRatio = ValidateContrast(value, nameof(MinimumTextContrastRatio));
        }

        /// <summary>Minimum large-text contrast ratio.</summary>
        public double MinimumLargeTextContrastRatio {
            get => _minimumLargeTextContrastRatio;
            set => _minimumLargeTextContrastRatio = ValidateContrast(value, nameof(MinimumLargeTextContrastRatio));
        }

        /// <summary>Large regular text threshold in points.</summary>
        public double LargeTextThresholdPoints { get; set; } = 18D;

        /// <summary>Large bold text threshold in points.</summary>
        public double LargeBoldTextThresholdPoints { get; set; } = 14D;

        /// <summary>Creates options initialized for a built-in profile.</summary>
        public static PowerPointAccessibilityOptions ForProfile(PowerPointAccessibilityPolicyProfile profile) {
            return profile == PowerPointAccessibilityPolicyProfile.Strict
                ? new PowerPointAccessibilityOptions {
                    Profile = profile,
                    RequireDocumentTitle = true,
                    RequireSlideTitles = true,
                    RequireShapeTitles = true,
                    RequireAlternativeText = true,
                    RequireLanguage = true,
                    RequireTableHeaders = true,
                    CheckContrast = true,
                    CheckMeaningfulLinks = true,
                    CheckColorOnlyMeaning = true
                }
                : new PowerPointAccessibilityOptions { Profile = profile };
        }

        internal PowerPointAccessibilityOptions CloneValidated() {
            if (double.IsNaN(LargeTextThresholdPoints) || double.IsInfinity(LargeTextThresholdPoints) ||
                LargeTextThresholdPoints <= 0D) throw new ArgumentOutOfRangeException(nameof(LargeTextThresholdPoints));
            if (double.IsNaN(LargeBoldTextThresholdPoints) || double.IsInfinity(LargeBoldTextThresholdPoints) ||
                LargeBoldTextThresholdPoints <= 0D) throw new ArgumentOutOfRangeException(nameof(LargeBoldTextThresholdPoints));
            return new PowerPointAccessibilityOptions {
                Profile = Profile,
                RequireDocumentTitle = RequireDocumentTitle,
                RequireSlideTitles = RequireSlideTitles,
                IncludeHiddenSlides = IncludeHiddenSlides,
                RequireShapeTitles = RequireShapeTitles,
                RequireAlternativeText = RequireAlternativeText,
                RequireLanguage = RequireLanguage,
                RequireTableHeaders = RequireTableHeaders,
                CheckContrast = CheckContrast,
                CheckMeaningfulLinks = CheckMeaningfulLinks,
                CheckColorOnlyMeaning = CheckColorOnlyMeaning,
                MinimumTextContrastRatio = MinimumTextContrastRatio,
                MinimumLargeTextContrastRatio = MinimumLargeTextContrastRatio,
                LargeTextThresholdPoints = LargeTextThresholdPoints,
                LargeBoldTextThresholdPoints = LargeBoldTextThresholdPoints,
                MaximumShapeCount = MaximumShapeCount,
                MaximumGroupDepth = MaximumGroupDepth
            };
        }

        private static int ValidatePositive(int value, string name) {
            if (value <= 0) throw new ArgumentOutOfRangeException(name, "Value must be positive.");
            return value;
        }

        private static double ValidateContrast(double value, string name) {
            if (double.IsNaN(value) || double.IsInfinity(value) || value < 1D || value > 21D) {
                throw new ArgumentOutOfRangeException(name, "Contrast ratio must be between 1 and 21.");
            }
            return value;
        }
    }
}
