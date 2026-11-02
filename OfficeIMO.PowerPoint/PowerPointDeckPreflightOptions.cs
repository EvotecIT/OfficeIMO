using System;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Configures deterministic layout checks performed before a presentation is saved or published.
    /// </summary>
    public sealed class PowerPointDeckPreflightOptions {
        private OfficeLength _minimumReadableFontSize = OfficeLength.Points(10);
        private double _minimumCollisionOverlapRatio = 0.2D;
        private OfficeLength _collisionTolerance = OfficeLength.Points(1);
        private OfficeLength _defaultFontSize = OfficeLength.Points(18);
        private OfficeLength _maximumDecorativeBleed = OfficeLength.Points(36);
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

        /// <summary>Maximum intentional decorative bleed beyond any slide edge.</summary>
        public OfficeLength MaximumDecorativeBleed {
            get => _maximumDecorativeBleed;
            set {
                if (value < OfficeLength.Zero) {
                    throw new ArgumentOutOfRangeException(nameof(MaximumDecorativeBleed),
                        "Decorative bleed must be a non-negative length.");
                }
                _maximumDecorativeBleed = value;
            }
        }

        /// <summary>Minimum acceptable resolved font size.</summary>
        public OfficeLength MinimumReadableFontSize {
            get => _minimumReadableFontSize;
            set => _minimumReadableFontSize = RequirePositive(value, nameof(MinimumReadableFontSize));
        }

        /// <summary>Fallback font size used when a text run does not carry an explicit size.</summary>
        public OfficeLength DefaultFontSize {
            get => _defaultFontSize;
            set => _defaultFontSize = RequirePositive(value, nameof(DefaultFontSize));
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

        /// <summary>Small physical tolerance applied to bounds and containment comparisons.</summary>
        public OfficeLength CollisionTolerance {
            get => _collisionTolerance;
            set {
                if (value < OfficeLength.Zero) {
                    throw new ArgumentOutOfRangeException(nameof(CollisionTolerance),
                        "Collision tolerance must be a non-negative length.");
                }

                _collisionTolerance = value;
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
            MaximumDecorativeBleed = MaximumDecorativeBleed,
            MinimumReadableFontSize = MinimumReadableFontSize,
            DefaultFontSize = DefaultFontSize,
            MinimumCollisionOverlapRatio = MinimumCollisionOverlapRatio,
            CollisionTolerance = CollisionTolerance,
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

        private static OfficeLength RequirePositive(OfficeLength value, string propertyName) {
            if (value <= OfficeLength.Zero) {
                throw new ArgumentOutOfRangeException(propertyName, "Value must be a positive length.");
            }

            return value;
        }
    }
}
