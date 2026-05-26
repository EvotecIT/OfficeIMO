using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents a Visio page layer stored in the page ShapeSheet.
    /// </summary>
    public sealed class VisioLayer {
        /// <summary>
        /// Creates a layer with the provided display name.
        /// </summary>
        /// <param name="name">Layer name shown in Visio.</param>
        /// <param name="nameU">Universal layer name. Defaults to <paramref name="name"/>.</param>
        public VisioLayer(string name, string? nameU = null) {
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("Layer name cannot be null or whitespace.", nameof(name));
            }

            Name = name;
            NameU = string.IsNullOrWhiteSpace(nameU) ? name : nameU!;
        }

        /// <summary>
        /// Layer name shown in Visio.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Universal layer name used for stable matching.
        /// </summary>
        public string NameU { get; set; }

        /// <summary>
        /// Visio layer color index.
        /// </summary>
        public int Color { get; set; } = 255;

        /// <summary>
        /// Visio layer status value.
        /// </summary>
        public int Status { get; set; }

        /// <summary>
        /// Whether layer members are visible.
        /// </summary>
        public bool Visible { get; set; } = true;

        /// <summary>
        /// Whether layer members are printed.
        /// </summary>
        public bool Print { get; set; } = true;

        /// <summary>
        /// Whether the layer is active in Visio.
        /// </summary>
        public bool Active { get; set; }

        /// <summary>
        /// Whether layer members are locked.
        /// </summary>
        public bool Lock { get; set; }

        /// <summary>
        /// Whether snapping to layer members is enabled.
        /// </summary>
        public bool Snap { get; set; } = true;

        /// <summary>
        /// Whether glue to layer members is enabled.
        /// </summary>
        public bool Glue { get; set; } = true;

        /// <summary>
        /// Visio color transparency value.
        /// </summary>
        public int ColorTransparency { get; set; }

        internal int? SourceIndex { get; set; }

        internal IList<XAttribute> PreservedRowAttributes { get; } = new List<XAttribute>();

        internal IList<XElement> PreservedCells { get; } = new List<XElement>();
    }
}
