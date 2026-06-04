using System;

namespace OfficeIMO.Visio {
    /// <summary>
    /// A discovered swimlane phase column.
    /// </summary>
    public sealed class VisioSwimlanePhase {
        internal VisioSwimlanePhase(string id, VisioShape header) {
            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException("Phase id cannot be empty.", nameof(id));
            }

            Id = id;
            Header = header ?? throw new ArgumentNullException(nameof(header));
        }

        /// <summary>Phase identifier used by swimlane activity placement metadata.</summary>
        public string Id { get; }

        /// <summary>Phase header shape defining the phase column center.</summary>
        public VisioShape Header { get; }
    }
}
