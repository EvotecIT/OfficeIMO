using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents a row in a Visio ShapeSheet User section.
    /// </summary>
    public sealed class VisioUserCell {
        /// <summary>
        /// Initializes a new user-defined cell row.
        /// </summary>
        /// <param name="name">Row name.</param>
        /// <param name="value">Value cell contents.</param>
        public VisioUserCell(string name, string? value = null) {
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("User cell name cannot be empty.", nameof(name));
            }

            Name = name;
            Value = value;
        }

        /// <summary>
        /// Row name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Value cell contents.
        /// </summary>
        public string? Value { get; set; }

        /// <summary>
        /// Optional unit for the Value cell.
        /// </summary>
        public string? Unit { get; set; }

        /// <summary>
        /// Optional ShapeSheet formula for the Value cell.
        /// </summary>
        public string? Formula { get; set; }

        /// <summary>
        /// Optional prompt cell contents.
        /// </summary>
        public string? Prompt { get; set; }

        /// <summary>
        /// Optional ShapeSheet formula for the Prompt cell.
        /// </summary>
        public string? PromptFormula { get; set; }

        internal int? RowIndex { get; set; }

        internal IList<XAttribute> PreservedRowAttributes { get; } = new List<XAttribute>();

        internal IList<XAttribute> PreservedValueAttributes { get; } = new List<XAttribute>();

        internal IList<XAttribute> PreservedPromptAttributes { get; } = new List<XAttribute>();

        internal IList<XElement> PreservedCells { get; } = new List<XElement>();
    }
}
