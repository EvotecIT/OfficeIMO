using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents a Visio ShapeSheet Hyperlink row on a shape or connector.
    /// Explicit cell value assignments replace imported formulas, null conditions and errors, even when unchanged.
    /// </summary>
    public sealed class VisioHyperlink {
        private string? _description, _address, _subAddress, _extraInfo, _frame, _sortKey;
        private bool _newWindow, _default, _invisible;
        private readonly HashSet<string> _assignedValues = new(StringComparer.Ordinal);

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
        /// Row name stored in the Hyperlink section. When omitted on a new row, OfficeIMO chooses
        /// an unused Row_1, Row_2, and so on, respecting explicit and inherited master row names.
        /// Explicit names must be unique within the local section; matching a master row overrides it.
        /// </summary>
        public string? RowName { get; set; }

        /// <summary>
        /// Description displayed for the hyperlink.
        /// </summary>
        public string? Description { get => _description; set => Assign(ref _description, value, "Description"); }

        /// <summary>
        /// External hyperlink address.
        /// </summary>
        public string? Address { get => _address; set => Assign(ref _address, value, "Address"); }

        /// <summary>
        /// Optional target inside the addressed document.
        /// </summary>
        public string? SubAddress { get => _subAddress; set => Assign(ref _subAddress, value, "SubAddress"); }

        /// <summary>
        /// Optional query-string style extra information.
        /// </summary>
        public string? ExtraInfo { get => _extraInfo; set => Assign(ref _extraInfo, value, "ExtraInfo"); }

        /// <summary>
        /// Optional target frame.
        /// </summary>
        public string? Frame { get => _frame; set => Assign(ref _frame, value, "Frame"); }

        /// <summary>
        /// Opens the hyperlink in a new window when supported by Visio.
        /// </summary>
        public bool NewWindow { get => _newWindow; set => Assign(ref _newWindow, value, "NewWindow"); }

        /// <summary>
        /// Marks this hyperlink as the default hyperlink for the shape.
        /// </summary>
        public bool Default { get => _default; set => Assign(ref _default, value, "Default"); }

        /// <summary>
        /// Hides the hyperlink from normal Visio hyperlink UI.
        /// </summary>
        public bool Invisible { get => _invisible; set => Assign(ref _invisible, value, "Invisible"); }

        /// <summary>
        /// Optional sort key used by Visio.
        /// </summary>
        public string? SortKey { get => _sortKey; set => Assign(ref _sortKey, value, "SortKey"); }

        internal IEnumerable<string> AssignedValues => _assignedValues;
        internal void ResetValueAssignments() => _assignedValues.Clear();
        internal void CopyValueAssignmentsFrom(VisioHyperlink source) {
            _assignedValues.Clear();
            _assignedValues.UnionWith(source._assignedValues);
        }

        private void Assign<T>(ref T field, T value, string cellName) {
            field = value;
            _assignedValues.Add(cellName);
            if (!PreservedKnownCells.TryGetValue(cellName, out XElement? cell)) return;
            cell.Attribute("F")?.Remove();
            cell.Attribute("E")?.Remove();
            cell.Attribute("Err")?.Remove();
        }

        internal int? RowIndex { get; set; }

        internal IList<XAttribute> PreservedRowAttributes { get; } = new List<XAttribute>();

        internal IDictionary<string, XElement> PreservedKnownCells { get; } = new Dictionary<string, XElement>(StringComparer.OrdinalIgnoreCase);

        internal IList<XElement> PreservedCells { get; } = new List<XElement>();
    }
}
