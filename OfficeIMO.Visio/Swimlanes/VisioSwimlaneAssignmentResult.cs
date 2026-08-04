using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Visio {
    /// <summary>Result of assigning swimlane activity metadata from geometry.</summary>
    public sealed class VisioSwimlaneAssignmentResult {
        internal VisioSwimlaneAssignmentResult(
            IReadOnlyList<VisioSwimlaneActivityPlacement> assigned,
            IReadOnlyList<string> unassignedShapeIds) {
            Assigned = new ReadOnlyCollection<VisioSwimlaneActivityPlacement>(
                new List<VisioSwimlaneActivityPlacement>(assigned));
            UnassignedShapeIds = new ReadOnlyCollection<string>(
                new List<string>(unassignedShapeIds));
        }

        /// <summary>Activities assigned to both a lane and phase.</summary>
        public IReadOnlyList<VisioSwimlaneActivityPlacement> Assigned { get; }

        /// <summary>Activity identifiers whose centers do not fall in a unique lane/phase cell.</summary>
        public IReadOnlyList<string> UnassignedShapeIds { get; }

        /// <summary>Whether every discovered activity received an unambiguous assignment.</summary>
        public bool Complete => UnassignedShapeIds.Count == 0;
    }
}
