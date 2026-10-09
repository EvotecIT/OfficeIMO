using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents a row in a Visio ShapeSheet User section.
    /// </summary>
    public sealed class VisioUserCell {
        private string? _value;
        private string? _prompt;
        private string? _formula;
        private string? _promptFormula;
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
        /// Value cell contents. An explicit assignment replaces an imported native null condition, even when unchanged.
        /// </summary>
        public string? Value {
            get => _value;
            set { _value = value; ValueAssigned = true; }
        }

        /// <summary>
        /// Optional unit for the Value cell.
        /// </summary>
        public string? Unit { get; set; }

        /// <summary>
        /// Optional ShapeSheet formula for the Value cell. Changing it clears the preserved producer error.
        /// </summary>
        public string? Formula {
            get => _formula;
            set { if (_formula != value) ClearFormulaError(PreservedValueAttributes); _formula = value; }
        }

        /// <summary>
        /// Optional prompt cell contents. An explicit assignment replaces an imported native null condition, even when unchanged.
        /// </summary>
        public string? Prompt {
            get => _prompt;
            set { _prompt = value; PromptAssigned = true; }
        }

        // An explicit same-value write replaces a legacy null condition. Loading and
        // copying initialize these flags separately from the public assignment API.
        internal bool ValueAssigned { get; private set; }
        internal bool PromptAssigned { get; private set; }
        internal void ResetValueAssignments() { ValueAssigned = false; PromptAssigned = false; }
        internal void CopyValueAssignmentsFrom(VisioUserCell source) {
            ValueAssigned = source.ValueAssigned; PromptAssigned = source.PromptAssigned;
        }

        /// <summary>
        /// Optional ShapeSheet formula for the Prompt cell. Changing it clears the preserved producer error.
        /// </summary>
        public string? PromptFormula {
            get => _promptFormula;
            set { if (_promptFormula != value) ClearFormulaError(PreservedPromptAttributes); _promptFormula = value; }
        }

        internal int? RowIndex { get; set; }

        internal IList<XAttribute> PreservedRowAttributes { get; } = new List<XAttribute>();

        internal IList<XAttribute> PreservedValueAttributes { get; } = new List<XAttribute>();

        internal IList<XAttribute> PreservedPromptAttributes { get; } = new List<XAttribute>();

        internal IList<XElement> PreservedCells { get; } = new List<XElement>();

        // Graph copying rebinds an equivalent formula to new native identities. Its cached
        // error belongs to that formula and must not be treated as a caller's formula edit.
        internal void RemapFormulas(IReadOnlyDictionary<string, string> ids) {
            _formula = VisioShapeFormulaReferences.Rewrite(_formula, ids);
            _promptFormula = VisioShapeFormulaReferences.Rewrite(_promptFormula, ids);
        }

        private static void ClearFormulaError(IList<XAttribute> attributes) {
            for (int index = attributes.Count - 1; index >= 0; index--)
                if (attributes[index].Name == "E" || attributes[index].Name == "Err") attributes.RemoveAt(index);
        }
    }
}
