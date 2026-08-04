using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    /// <summary>One typed cell in a Visio ShapeSheet row.</summary>
    public sealed class VisioShapeSheetCell {
        internal VisioShapeSheetCell(XElement element) {
            Element = new XElement(element);
        }

        /// <summary>Creates a ShapeSheet cell.</summary>
        public VisioShapeSheetCell(string name, string? value = null,
            string? formula = null, string? unit = null) {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Cell name cannot be empty.", nameof(name));
            Element = new XElement(VisioShapeSheetSection.VisioNamespace + "Cell",
                new XAttribute("N", name));
            Value = value;
            Formula = formula;
            Unit = unit;
        }

        internal XElement Element { get; }

        /// <summary>Gets the ShapeSheet cell name.</summary>
        public string Name => (string?)Element.Attribute("N") ?? string.Empty;

        /// <summary>Gets or sets the cached cell value.</summary>
        public string? Value {
            get => (string?)Element.Attribute("V");
            set => Element.SetAttributeValue("V", value);
        }

        /// <summary>Gets or sets the native ShapeSheet formula.</summary>
        public string? Formula {
            get => (string?)Element.Attribute("F");
            set => Element.SetAttributeValue("F", value);
        }

        /// <summary>Gets or sets the native unit token.</summary>
        public string? Unit {
            get => (string?)Element.Attribute("U");
            set => Element.SetAttributeValue("U", value);
        }

        /// <summary>Gets the producer error token, when present.</summary>
        public string? Error => (string?)Element.Attribute("E");
    }

    /// <summary>One named or indexed row in a Visio ShapeSheet section.</summary>
    public sealed class VisioShapeSheetRow {
        private readonly List<VisioShapeSheetCell> _cells;

        internal VisioShapeSheetRow(XElement element) {
            Element = new XElement(element);
            _cells = Element.Elements(VisioShapeSheetSection.VisioNamespace + "Cell")
                .Select(cell => new VisioShapeSheetCell(cell)).ToList();
        }

        /// <summary>Creates a named ShapeSheet row.</summary>
        public VisioShapeSheetRow(string name) {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Row name cannot be empty.", nameof(name));
            Element = new XElement(VisioShapeSheetSection.VisioNamespace + "Row",
                new XAttribute("N", name));
            _cells = new List<VisioShapeSheetCell>();
        }

        internal XElement Element { get; }

        /// <summary>Gets or sets the universal row name.</summary>
        public string? Name {
            get => (string?)Element.Attribute("N");
            set => Element.SetAttributeValue("N", value);
        }

        /// <summary>Gets or sets the zero-based native row index.</summary>
        public int? Index {
            get => int.TryParse((string?)Element.Attribute("IX"), out int value) ? value : (int?)null;
            set => Element.SetAttributeValue("IX", value);
        }

        /// <summary>Gets the row's typed cells.</summary>
        public IReadOnlyList<VisioShapeSheetCell> Cells =>
            new ReadOnlyCollection<VisioShapeSheetCell>(_cells);

        /// <summary>Finds a cell by native name.</summary>
        public VisioShapeSheetCell? FindCell(string name) =>
            _cells.FirstOrDefault(cell => string.Equals(cell.Name, name,
                StringComparison.OrdinalIgnoreCase));

        /// <summary>Sets or creates a cell while retaining unmodeled row content.</summary>
        public VisioShapeSheetCell SetCell(string name, string? value = null,
            string? formula = null, string? unit = null) {
            VisioShapeSheetCell? cell = FindCell(name);
            if (cell == null) {
                cell = new VisioShapeSheetCell(name);
                _cells.Add(cell);
            }
            cell.Value = value;
            cell.Formula = formula;
            cell.Unit = unit;
            return cell;
        }

        internal XElement ToXElement() {
            XElement clone = new(Element);
            XElement[] existing = clone.Elements(
                VisioShapeSheetSection.VisioNamespace + "Cell").ToArray();
            int retained = Math.Min(existing.Length, _cells.Count);
            for (int index = 0; index < retained; index++)
                existing[index].ReplaceWith(new XElement(_cells[index].Element));
            for (int index = retained; index < existing.Length; index++)
                existing[index].Remove();
            for (int index = retained; index < _cells.Count; index++)
                clone.Add(new XElement(_cells[index].Element));
            return clone;
        }
    }

    /// <summary>
    /// Typed, source-preserving view of an otherwise unmodeled Visio ShapeSheet section.
    /// </summary>
    public sealed class VisioShapeSheetSection {
        internal static readonly XNamespace VisioNamespace =
            "http://schemas.microsoft.com/office/visio/2012/main";
        private readonly List<VisioShapeSheetRow> _rows;

        internal VisioShapeSheetSection(XElement element) {
            Element = new XElement(element);
            _rows = Element.Elements(VisioNamespace + "Row")
                .Select(row => new VisioShapeSheetRow(row)).ToList();
        }

        /// <summary>Creates a named ShapeSheet section.</summary>
        public VisioShapeSheetSection(string name) {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Section name cannot be empty.", nameof(name));
            Element = new XElement(VisioNamespace + "Section",
                new XAttribute("N", name));
            _rows = new List<VisioShapeSheetRow>();
        }

        internal XElement Element { get; }

        /// <summary>Gets or sets the native section name.</summary>
        public string? Name {
            get => (string?)Element.Attribute("N");
            set => Element.SetAttributeValue("N", value);
        }

        /// <summary>Gets or sets the native section index.</summary>
        public int? Index {
            get => int.TryParse((string?)Element.Attribute("IX"), out int value) ? value : (int?)null;
            set => Element.SetAttributeValue("IX", value);
        }

        /// <summary>Gets the typed rows while unmodeled attributes and child elements remain preserved.</summary>
        public IReadOnlyList<VisioShapeSheetRow> Rows =>
            new ReadOnlyCollection<VisioShapeSheetRow>(_rows);

        /// <summary>Finds a row by universal name.</summary>
        public VisioShapeSheetRow? FindRow(string name) =>
            _rows.FirstOrDefault(row => string.Equals(row.Name, name,
                StringComparison.OrdinalIgnoreCase));

        /// <summary>Gets or creates a named row.</summary>
        public VisioShapeSheetRow GetOrAddRow(string name) {
            VisioShapeSheetRow? row = FindRow(name);
            if (row != null) return row;
            row = new VisioShapeSheetRow(name);
            _rows.Add(row);
            return row;
        }

        internal XElement ToXElement() {
            XElement clone = new(Element);
            XElement[] existing = clone.Elements(VisioNamespace + "Row").ToArray();
            int retained = Math.Min(existing.Length, _rows.Count);
            for (int index = 0; index < retained; index++)
                existing[index].ReplaceWith(_rows[index].ToXElement());
            for (int index = retained; index < existing.Length; index++)
                existing[index].Remove();
            for (int index = retained; index < _rows.Count; index++)
                clone.Add(_rows[index].ToXElement());
            return clone;
        }
    }
}
