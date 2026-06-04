using System;

namespace OfficeIMO.Visio {
    /// <summary>
    /// A discovered swimlane lane with its body and optional header shape.
    /// </summary>
    public sealed class VisioSwimlaneLane {
        internal VisioSwimlaneLane(string id, VisioShape body, VisioShape? header) {
            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException("Lane id cannot be empty.", nameof(id));
            }

            Id = id;
            Body = body ?? throw new ArgumentNullException(nameof(body));
            Header = header;
        }

        /// <summary>Lane identifier used by swimlane activity placement metadata.</summary>
        public string Id { get; }

        /// <summary>Lane body shape covering the process grid area.</summary>
        public VisioShape Body { get; }

        /// <summary>Optional lane header shape.</summary>
        public VisioShape? Header { get; }
    }
}
