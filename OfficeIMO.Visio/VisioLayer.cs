using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents a Visio page layer stored in the page ShapeSheet.
    /// Explicit value assignments replace imported formulas and producer error/null state, even when unchanged.
    /// </summary>
    public sealed class VisioLayer {
        private string _name = string.Empty, _nameU = string.Empty;
        private int _color = 255, _status, _colorTransparency;
        private bool _visible = true, _print = true, _active, _lock, _snap = true, _glue = true;
        private readonly HashSet<string> _assignedValues = new(StringComparer.Ordinal);
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
        public string Name { get => _name; set => Assign(ref _name, value, "Name"); }

        /// <summary>
        /// Universal layer name used for stable matching.
        /// </summary>
        public string NameU { get => _nameU; set => Assign(ref _nameU, value, "NameUniv"); }

        /// <summary>
        /// Visio layer color index.
        /// </summary>
        public int Color { get => _color; set => Assign(ref _color, value, "Color"); }

        /// <summary>
        /// Visio layer status value.
        /// </summary>
        public int Status { get => _status; set => Assign(ref _status, value, "Status"); }

        /// <summary>
        /// Whether layer members are visible.
        /// </summary>
        public bool Visible { get => _visible; set => Assign(ref _visible, value, "Visible"); }

        /// <summary>
        /// Whether layer members are printed.
        /// </summary>
        public bool Print { get => _print; set => Assign(ref _print, value, "Print"); }

        /// <summary>
        /// Whether the layer is active in Visio.
        /// </summary>
        public bool Active { get => _active; set => Assign(ref _active, value, "Active"); }

        /// <summary>
        /// Whether layer members are locked.
        /// </summary>
        public bool Lock { get => _lock; set => Assign(ref _lock, value, "Lock"); }

        /// <summary>
        /// Whether snapping to layer members is enabled.
        /// </summary>
        public bool Snap { get => _snap; set => Assign(ref _snap, value, "Snap"); }

        /// <summary>
        /// Whether glue to layer members is enabled.
        /// </summary>
        public bool Glue { get => _glue; set => Assign(ref _glue, value, "Glue"); }

        /// <summary>
        /// Visio color transparency value.
        /// </summary>
        public int ColorTransparency { get => _colorTransparency; set => Assign(ref _colorTransparency, value, "ColorTrans"); }

        internal IEnumerable<string> AssignedValues => _assignedValues;
        internal void ResetValueAssignments() => _assignedValues.Clear();
        internal void CopyValueAssignmentsFrom(VisioLayer source) {
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

        internal int? SourceIndex { get; set; }

        internal IList<XAttribute> PreservedRowAttributes { get; } = new List<XAttribute>();

        internal IDictionary<string, XElement> PreservedKnownCells { get; } = new Dictionary<string, XElement>(StringComparer.OrdinalIgnoreCase);

        internal IList<XElement> PreservedCells { get; } = new List<XElement>();
    }
}
