using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents one row in the Visio ShapeSheet Shape Data (Prop) section.
    /// Explicit assignments to string values replace imported native null conditions, even when unchanged.
    /// </summary>
    public sealed class VisioShapeDataRow {
        private string? _value, _label, _prompt, _format, _sortKey, _calendar, _langId;
        private readonly HashSet<string> _assignedValues = new(StringComparer.Ordinal);
        /// <summary>
        /// Initializes a new Shape Data row.
        /// </summary>
        /// <param name="name">ShapeSheet row name.</param>
        /// <param name="value">Initial value.</param>
        public VisioShapeDataRow(string name, string? value = null) {
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("Shape data name cannot be empty.", nameof(name));
            }

            Name = name;
            Value = value;
        }

        /// <summary>ShapeSheet row name.</summary>
        public string Name { get; }

        /// <summary>Shape data value.</summary>
        public string? Value { get => _value; set => Assign(ref _value, value, "Value"); }

        /// <summary>Optional value unit.</summary>
        public string? ValueUnit { get; set; }

        /// <summary>Optional value formula.</summary>
        public string? ValueFormula { get; set; }

        /// <summary>Label shown in the Visio Shape Data window.</summary>
        public string? Label { get => _label; set => Assign(ref _label, value, "Label"); }

        /// <summary>Optional label formula.</summary>
        public string? LabelFormula { get; set; }

        /// <summary>Prompt shown as help text in the Visio Shape Data window.</summary>
        public string? Prompt { get => _prompt; set => Assign(ref _prompt, value, "Prompt"); }

        /// <summary>Optional prompt formula.</summary>
        public string? PromptFormula { get; set; }

        /// <summary>Shape data type.</summary>
        public VisioShapeDataType? Type { get; set; }

        /// <summary>Optional type formula.</summary>
        public string? TypeFormula { get; set; }

        /// <summary>Format picture or list items, depending on the shape data type.</summary>
        public string? Format { get => _format; set => Assign(ref _format, value, "Format"); }

        /// <summary>Optional format formula.</summary>
        public string? FormatFormula { get; set; }

        /// <summary>Sort key used by Visio's Shape Data window.</summary>
        public string? SortKey { get => _sortKey; set => Assign(ref _sortKey, value, "SortKey"); }

        /// <summary>Optional sort key formula.</summary>
        public string? SortKeyFormula { get; set; }

        /// <summary>Whether the data row is hidden in the Shape Data window.</summary>
        public bool? Invisible { get; set; }

        /// <summary>Optional invisible formula.</summary>
        public string? InvisibleFormula { get; set; }

        /// <summary>Whether Visio asks for this value when the shape is created, copied, or duplicated.</summary>
        public bool? Verify { get; set; }

        /// <summary>Optional verify formula.</summary>
        public string? VerifyFormula { get; set; }

        /// <summary>Whether this row is linked to an external data recordset.</summary>
        public bool? DataLinked { get; set; }

        /// <summary>Optional data-linked formula.</summary>
        public string? DataLinkedFormula { get; set; }

        /// <summary>Calendar type used for date values.</summary>
        public string? Calendar { get => _calendar; set => Assign(ref _calendar, value, "Calendar"); }

        /// <summary>Optional calendar formula.</summary>
        public string? CalendarFormula { get; set; }

        /// <summary>Language identifier used for the value.</summary>
        public string? LangId { get => _langId; set => Assign(ref _langId, value, "LangID"); }

        internal IEnumerable<string> AssignedValues => _assignedValues;
        internal void ResetValueAssignments() => _assignedValues.Clear();
        internal void CopyValueAssignmentsFrom(VisioShapeDataRow source) {
            _assignedValues.Clear();
            _assignedValues.UnionWith(source._assignedValues);
        }

        private void Assign(ref string? field, string? value, string cellName) {
            field = value;
            _assignedValues.Add(cellName);
        }

        /// <summary>Optional language identifier formula.</summary>
        public string? LangIdFormula { get; set; }

        internal string? LoadedValue { get; set; }

        internal string? MirroredDataValue { get; set; }

        internal int? RowIndex { get; set; }

        internal void ClearValueFormula() {
            ValueFormula = null;
            if (!PreservedKnownCells.TryGetValue("Value", out XElement? cell)) return;
            cell.Attribute("F")?.Remove();
            cell.Attribute("E")?.Remove();
            cell.Attribute("Err")?.Remove();
        }

        internal IList<XAttribute> PreservedRowAttributes { get; } = new List<XAttribute>();

        internal IDictionary<string, XElement> PreservedKnownCells { get; } = new Dictionary<string, XElement>(StringComparer.OrdinalIgnoreCase);

        internal IList<string> PreservedCellOrder { get; } = new List<string>();

        internal IList<XElement> PreservedCells { get; } = new List<XElement>();

        internal static readonly string[] CellOrder = {
            "Value",
            "Label",
            "Prompt",
            "Type",
            "Format",
            "SortKey",
            "Invisible",
            "Verify",
            "DataLinked",
            "Calendar",
            "LangID"
        };
    }
}
