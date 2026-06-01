using System;

namespace OfficeIMO.Visio.Stencils {
    /// <summary>
    /// Describes a native connection point discovered from a stencil master.
    /// Coordinates are expressed in the stencil master's local coordinate space.
    /// </summary>
    public sealed class VisioStencilConnectionPoint {
        /// <summary>
        /// Initializes a new stencil connection point definition.
        /// </summary>
        public VisioStencilConnectionPoint(double x, double y, double dirX, double dirY, int? sectionIndex = null, double? sourceWidth = null, double? sourceHeight = null) {
            if (!IsFinite(x)) throw new ArgumentOutOfRangeException(nameof(x), "Connection point X must be finite.");
            if (!IsFinite(y)) throw new ArgumentOutOfRangeException(nameof(y), "Connection point Y must be finite.");
            if (!IsFinite(dirX)) throw new ArgumentOutOfRangeException(nameof(dirX), "Connection point DirX must be finite.");
            if (!IsFinite(dirY)) throw new ArgumentOutOfRangeException(nameof(dirY), "Connection point DirY must be finite.");
            if (sectionIndex.HasValue && sectionIndex.Value < 0) throw new ArgumentOutOfRangeException(nameof(sectionIndex), "Connection point section index must be zero or greater.");
            if (sourceWidth.HasValue && (!IsFinite(sourceWidth.Value) || sourceWidth.Value <= 0)) throw new ArgumentOutOfRangeException(nameof(sourceWidth), "Connection point source width must be positive and finite.");
            if (sourceHeight.HasValue && (!IsFinite(sourceHeight.Value) || sourceHeight.Value <= 0)) throw new ArgumentOutOfRangeException(nameof(sourceHeight), "Connection point source height must be positive and finite.");

            X = x;
            Y = y;
            DirX = dirX;
            DirY = dirY;
            SectionIndex = sectionIndex;
            SourceWidth = sourceWidth;
            SourceHeight = sourceHeight;
        }

        /// <summary>Local X coordinate.</summary>
        public double X { get; }

        /// <summary>Local Y coordinate.</summary>
        public double Y { get; }

        /// <summary>Directional X component.</summary>
        public double DirX { get; }

        /// <summary>Directional Y component.</summary>
        public double DirY { get; }

        /// <summary>Original Visio Connection section row index, when present.</summary>
        public int? SectionIndex { get; }

        /// <summary>Native master width used as the source coordinate space, when known.</summary>
        public double? SourceWidth { get; }

        /// <summary>Native master height used as the source coordinate space, when known.</summary>
        public double? SourceHeight { get; }

        private static bool IsFinite(double value) {
            return !double.IsNaN(value) && !double.IsInfinity(value);
        }
    }
}
