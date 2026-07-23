using System;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Configures deterministic layout checks performed before a presentation is saved or published.
    /// </summary>
    public sealed class PowerPointDeckPreflightOptions {
        private double _minimumReadableFontSizePoints = 10D;
        private double _minimumCollisionOverlapRatio = 0.2D;
        private double _collisionTolerancePoints = 1D;
        private double _defaultFontSizePoints = 18D;
        private double _maximumDecorativeBleedPoints = 36D;
        private int _maximumShapeCount = 10000;
        private int _maximumGroupDepth = 128;

        /// <summary>Maximum number of shapes inspected across the presentation.</summary>
        public int MaximumShapeCount {
            get => _maximumShapeCount;
            set => _maximumShapeCount = RequirePositive(value, nameof(MaximumShapeCount));
        }

        /// <summary>Maximum supported nesting depth for grouped shapes.</summary>
        public int MaximumGroupDepth {
            get => _maximumGroupDepth;
            set => _maximumGroupDepth = RequirePositive(value, nameof(MaximumGroupDepth));
        }

        /// <summary>Checks whether explicit slide shapes extend beyond the slide canvas.</summary>
        public bool DetectOffSlideShapes { get; set; } = true;

        /// <summary>Checks whether text is clipped after deterministic measurement.</summary>
        public bool DetectTextOverflow { get; set; } = true;

        /// <summary>Checks whether normal auto-fit would reduce text below the readable threshold.</summary>
        public bool DetectUnreadableFontReduction { get; set; } = true;

        /// <summary>Checks for significant intersections between peer shapes.</summary>
        public bool DetectShapeCollisions { get; set; } = true;

        /// <summary>Checks picture relationships and image parts.</summary>
        public bool DetectMissingVisualAssets { get; set; } = true;

        /// <summary>Includes diagnostics from the shared slide visual snapshot.</summary>
        public bool IncludeVisualSnapshotDiagnostics { get; set; } = true;

        /// <summary>
        ///     Ignores a collision when one shape fully contains another. This avoids reporting intentional
        ///     text-on-panel and image-in-frame compositions as collisions.
        /// </summary>
        public bool IgnoreContainedShapeCollisions { get; set; } = true;

        /// <summary>
        ///     Allows non-text auto shapes to extend slightly beyond the canvas for intentional full-bleed design.
        /// </summary>
        public bool AllowDecorativeShapeBleed { get; set; } = true;

        /// <summary>Maximum intentional decorative bleed beyond any slide edge, in points.</summary>
        public double MaximumDecorativeBleedPoints {
            get => _maximumDecorativeBleedPoints;
            set {
                if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D) {
                    throw new ArgumentOutOfRangeException(nameof(MaximumDecorativeBleedPoints),
                        "Decorative bleed must be a finite non-negative number.");
                }
                _maximumDecorativeBleedPoints = value;
            }
        }

        /// <summary>Minimum acceptable resolved font size in points.</summary>
        public double MinimumReadableFontSizePoints {
            get => _minimumReadableFontSizePoints;
            set => _minimumReadableFontSizePoints = RequirePositive(value, nameof(MinimumReadableFontSizePoints));
        }

        /// <summary>Fallback font size used when a text run does not carry an explicit size.</summary>
        public double DefaultFontSizePoints {
            get => _defaultFontSizePoints;
            set => _defaultFontSizePoints = RequirePositive(value, nameof(DefaultFontSizePoints));
        }

        /// <summary>
        ///     Minimum intersection area divided by the smaller shape area before a peer overlap is reported.
        /// </summary>
        public double MinimumCollisionOverlapRatio {
            get => _minimumCollisionOverlapRatio;
            set {
                if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D || value > 1D) {
                    throw new ArgumentOutOfRangeException(nameof(MinimumCollisionOverlapRatio),
                        "Collision overlap ratio must be between zero and one.");
                }

                _minimumCollisionOverlapRatio = value;
            }
        }

        /// <summary>Small point-unit tolerance applied to bounds and containment comparisons.</summary>
        public double CollisionTolerancePoints {
            get => _collisionTolerancePoints;
            set {
                if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D) {
                    throw new ArgumentOutOfRangeException(nameof(CollisionTolerancePoints),
                        "Collision tolerance must be a finite non-negative number.");
                }

                _collisionTolerancePoints = value;
            }
        }

        /// <summary>Lowest finding severity that causes <c>SaveWithPreflight</c> to fail.</summary>
        public PowerPointDeckPreflightSeverity FailureSeverity { get; set; } = PowerPointDeckPreflightSeverity.Error;

        /// <summary>Creates a detached copy suitable for per-operation customization.</summary>
        public PowerPointDeckPreflightOptions Clone() => new PowerPointDeckPreflightOptions {
            DetectOffSlideShapes = DetectOffSlideShapes,
            DetectTextOverflow = DetectTextOverflow,
            DetectUnreadableFontReduction = DetectUnreadableFontReduction,
            DetectShapeCollisions = DetectShapeCollisions,
            DetectMissingVisualAssets = DetectMissingVisualAssets,
            IncludeVisualSnapshotDiagnostics = IncludeVisualSnapshotDiagnostics,
            IgnoreContainedShapeCollisions = IgnoreContainedShapeCollisions,
            AllowDecorativeShapeBleed = AllowDecorativeShapeBleed,
            MaximumDecorativeBleedPoints = MaximumDecorativeBleedPoints,
            MinimumReadableFontSizePoints = MinimumReadableFontSizePoints,
            DefaultFontSizePoints = DefaultFontSizePoints,
            MinimumCollisionOverlapRatio = MinimumCollisionOverlapRatio,
            CollisionTolerancePoints = CollisionTolerancePoints,
            MaximumShapeCount = MaximumShapeCount,
            MaximumGroupDepth = MaximumGroupDepth,
            FailureSeverity = FailureSeverity
        };

        private static int RequirePositive(int value, string propertyName) {
            if (value <= 0) {
                throw new ArgumentOutOfRangeException(propertyName, "Value must be positive.");
            }

            return value;
        }

        private static double RequirePositive(double value, string propertyName) {
            if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0D) {
                throw new ArgumentOutOfRangeException(propertyName, "Value must be a finite positive number.");
            }

            return value;
        }
    }
}
