using System;

namespace OfficeIMO.Visio {
    /// <summary>
    /// A discovered swimlane phase column.
    /// </summary>
    public sealed class VisioSwimlanePhase {
        internal VisioSwimlanePhase(string id, VisioShape header, int order) {
            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException("Phase id cannot be empty.", nameof(id));
            }

            Id = id;
            Header = header ?? throw new ArgumentNullException(nameof(header));
            Order = order;
        }

        /// <summary>Phase identifier used by swimlane activity placement metadata.</summary>
        public string Id { get; }

        /// <summary>Phase header shape defining the phase column center.</summary>
        public VisioShape Header { get; }

        /// <summary>Deterministic left-to-right phase order.</summary>
        public int Order { get; }

        /// <summary>Display name from the phase header.</summary>
        public string? Name => Header.Text;

        /// <summary>Current phase header bounds.</summary>
        public VisioShapeBounds Bounds => Header.GetShapeBounds();
    }
}
