using System;
using OfficeIMO.Visio.Diagrams;

namespace OfficeIMO.Visio {
    /// <summary>
    /// A discovered swimlane activity and its current lane/phase placement.
    /// </summary>
    public sealed class VisioSwimlaneActivityPlacement {
        internal VisioSwimlaneActivityPlacement(VisioShape shape, string? laneId, string? phaseId, VisioSwimlaneActivityKind? activityKind) {
            Shape = shape ?? throw new ArgumentNullException(nameof(shape));
            LaneId = laneId;
            PhaseId = phaseId;
            ActivityKind = activityKind;
        }

        /// <summary>Activity shape.</summary>
        public VisioShape Shape { get; }

        /// <summary>Current lane identifier, when known.</summary>
        public string? LaneId { get; }

        /// <summary>Current phase identifier, when known.</summary>
        public string? PhaseId { get; }

        /// <summary>Semantic activity kind, when known.</summary>
        public VisioSwimlaneActivityKind? ActivityKind { get; }
    }
}
