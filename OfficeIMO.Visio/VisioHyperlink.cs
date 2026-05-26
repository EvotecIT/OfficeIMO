using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents a Visio ShapeSheet Hyperlink row on a shape or connector.
    /// </summary>
    public sealed class VisioHyperlink {
        internal static readonly string[] CellOrder = {
            "Description",
            "Address",
            "SubAddress",
            "ExtraInfo",
            "Frame",
            "NewWindow",
            "Default",
            "Invisible",
            "SortKey"
        };

        /// <summary>
        /// Initializes a new hyperlink row.
        /// </summary>
        /// <param name="address">External hyperlink address.</param>
        /// <param name="description">Display description shown by Visio.</param>
        /// <param name="subAddress">Optional internal sub-address.</param>
        public VisioHyperlink(string? address = null, string? description = null, string? subAddress = null) {
            Address = address;
            Description = description;
            SubAddress = subAddress;
        }

        /// <summary>
        /// Row name stored in the Hyperlink section. When omitted, OfficeIMO writes Row_1, Row_2, and so on.
        /// </summary>
        public string? RowName { get; set; }

        /// <summary>
        /// Description displayed for the hyperlink.
        /// </summary>
        public string? Description { get; set; }

        /// <summary>
        /// External hyperlink address.
        /// </summary>
        public string? Address { get; set; }

        /// <summary>
        /// Optional target inside the addressed document.
        /// </summary>
        public string? SubAddress { get; set; }

        /// <summary>
        /// Optional query-string style extra information.
        /// </summary>
        public string? ExtraInfo { get; set; }

        /// <summary>
        /// Optional target frame.
        /// </summary>
        public string? Frame { get; set; }

        /// <summary>
        /// Opens the hyperlink in a new window when supported by Visio.
        /// </summary>
        public bool NewWindow { get; set; }

        /// <summary>
        /// Marks this hyperlink as the default hyperlink for the shape.
        /// </summary>
        public bool Default { get; set; }

        /// <summary>
        /// Hides the hyperlink from normal Visio hyperlink UI.
        /// </summary>
        public bool Invisible { get; set; }

        /// <summary>
        /// Optional sort key used by Visio.
        /// </summary>
        public string? SortKey { get; set; }

        internal int? RowIndex { get; set; }

        internal IList<XAttribute> PreservedRowAttributes { get; } = new List<XAttribute>();

        internal IDictionary<string, XElement> PreservedKnownCells { get; } = new Dictionary<string, XElement>(StringComparer.OrdinalIgnoreCase);

        internal IList<XElement> PreservedCells { get; } = new List<XElement>();
    }
}
