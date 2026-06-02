using System.Collections.Generic;

namespace OfficeIMO.Visio {
/// <summary>
    /// Snapshot of one connector waypoint.
    /// </summary>
    public sealed class VisioInspectionWaypointSnapshot {
        internal VisioInspectionWaypointSnapshot(double x, double y) {
            X = x;
            Y = y;
        }

        /// <summary>Waypoint X coordinate.</summary>
        public double X { get; }

        /// <summary>Waypoint Y coordinate.</summary>
        public double Y { get; }
    }

/// <summary>
    /// Snapshot of a Shape Data row.
    /// </summary>
    public sealed class VisioInspectionShapeDataSnapshot {
        internal VisioInspectionShapeDataSnapshot(string name, string? label, string? value, string? type, string? format, string? prompt) {
            Name = name;
            Label = label;
            Value = value;
            Type = type;
            Format = format;
            Prompt = prompt;
        }

        /// <summary>Shape Data row name.</summary>
        public string Name { get; }

        /// <summary>Shape Data display label.</summary>
        public string? Label { get; }

        /// <summary>Shape Data value.</summary>
        public string? Value { get; }

        /// <summary>Shape Data type.</summary>
        public string? Type { get; }

        /// <summary>Shape Data format string.</summary>
        public string? Format { get; }

        /// <summary>Shape Data prompt.</summary>
        public string? Prompt { get; }
    }

/// <summary>
    /// Snapshot of a User cell row.
    /// </summary>
    public sealed class VisioInspectionUserCellSnapshot {
        internal VisioInspectionUserCellSnapshot(string name, string? value, string? formula, string? prompt) {
            Name = name;
            Value = value;
            Formula = formula;
            Prompt = prompt;
        }

        /// <summary>User cell row name.</summary>
        public string Name { get; }

        /// <summary>User cell value.</summary>
        public string? Value { get; }

        /// <summary>User cell formula.</summary>
        public string? Formula { get; }

        /// <summary>User cell prompt.</summary>
        public string? Prompt { get; }
    }

/// <summary>
    /// Snapshot of one Visio shape connection point.
    /// </summary>
    public sealed class VisioInspectionConnectionPointSnapshot {
        internal VisioInspectionConnectionPointSnapshot(int index, int? sectionIndex, double x, double y, double dirX, double dirY) {
            Index = index;
            SectionIndex = sectionIndex;
            X = x;
            Y = y;
            DirX = dirX;
            DirY = dirY;
        }

        /// <summary>Zero-based position in the shape connection point collection.</summary>
        public int Index { get; }

        /// <summary>Original Visio Connection section row index, when loaded or assigned.</summary>
        public int? SectionIndex { get; }

        /// <summary>X coordinate relative to the shape.</summary>
        public double X { get; }

        /// <summary>Y coordinate relative to the shape.</summary>
        public double Y { get; }

        /// <summary>Directional X component.</summary>
        public double DirX { get; }

        /// <summary>Directional Y component.</summary>
        public double DirY { get; }
    }
}
