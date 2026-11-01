using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Walks worksheet DOM for structural edits. Cell content dominates a worksheet, and a typed descendant search
    /// visits every cell, so metadata searches skip the grid and grid searches use direct child links.
    /// </summary>
    internal static class ExcelWorksheetScan {
        [ThreadStatic]
        private static FormulaScope? _formulaScope;

        /// <summary>Finds the formula of a cell through sibling links, which allocates nothing unlike the typed property.</summary>
        internal static CellFormula? FindFormula(Cell cell) {
            for (OpenXmlElement? child = cell.FirstChild; child != null; child = child.NextSibling()) {
                if (child is CellFormula formula) return formula;
            }
            return null;
        }

        internal static CellValue? FindValue(Cell cell) {
            for (OpenXmlElement? child = cell.FirstChild; child != null; child = child.NextSibling()) {
                if (child is CellValue value) return value;
            }
            return null;
        }

        /// <summary>Sets the cached value of a cell. The typed property walks the schema particle, which allocates per call.</summary>
        internal static void SetValueText(Cell cell, string text) {
            if (FindValue(cell) is { } value) value.Text = text;
            else if (cell.FirstChild == null) cell.AppendChild(new CellValue(text));
            else cell.CellValue = new CellValue(text);
        }

        internal static void RemoveInlineString(Cell cell) {
            for (OpenXmlElement? child = cell.FirstChild; child != null; child = child.NextSibling()) {
                if (child is InlineString) {
                    child.Remove();
                    return;
                }
            }
        }

        internal static IEnumerable<Row> Rows(Worksheet worksheet) {
            for (OpenXmlElement? part = worksheet.FirstChild; part != null; part = part.NextSibling()) {
                if (part is not SheetData) continue;
                for (OpenXmlElement? row = part.FirstChild; row != null; row = row.NextSibling()) {
                    if (row is Row typed) yield return typed;
                }
            }
        }

        internal static IEnumerable<Cell> Cells(Worksheet worksheet) {
            foreach (Row row in Rows(worksheet)) {
                for (OpenXmlElement? cell = row.FirstChild; cell != null; cell = cell.NextSibling()) {
                    if (cell is Cell typed) yield return typed;
                }
            }
        }

        internal static IEnumerable<CellFormula> CellFormulas(Worksheet worksheet) =>
            _formulaScope?.Formulas(worksheet) ?? ScanCellFormulas(worksheet);

        private static IEnumerable<CellFormula> ScanCellFormulas(Worksheet worksheet) {
            foreach (Row row in Rows(worksheet)) {
                for (OpenXmlElement? cell = row.FirstChild; cell != null; cell = cell.NextSibling()) {
                    if (cell is Cell typed && FindFormula(typed) is { } formula) yield return formula;
                }
            }
        }

        /// <summary>
        /// Keeps the formula list of each worksheet for one structural edit, whose checks and rewrites all visit it.
        /// The edit must call <see cref="FormulaScope.Refresh"/> after it adds or removes formula cells.
        /// </summary>
        internal static FormulaScope BeginFormulaScope() => new FormulaScope();

        internal sealed class FormulaScope : IDisposable {
            private readonly FormulaScope? _parent = _formulaScope;
            private readonly Dictionary<Worksheet, List<CellFormula>> _formulas = new Dictionary<Worksheet, List<CellFormula>>();

            internal FormulaScope() => _formulaScope = this;

            internal IEnumerable<CellFormula> Formulas(Worksheet worksheet) {
                if (!_formulas.TryGetValue(worksheet, out List<CellFormula>? formulas)) {
                    formulas = new List<CellFormula>(ScanCellFormulas(worksheet));
                    _formulas.Add(worksheet, formulas);
                }
                return formulas;
            }

            internal void Refresh(Worksheet worksheet) => _formulas.Remove(worksheet);

            public void Dispose() => _formulaScope = _parent;
        }

        internal static bool HasCellFormulas(Worksheet worksheet) {
            foreach (CellFormula _ in CellFormulas(worksheet)) return true;
            return false;
        }

        /// <summary>
        /// Returns the descendants of <paramref name="worksheet"/> outside its cell grid, in document order.
        /// Structural edits ask this for each kind of metadata on each edit, so it follows element links rather than
        /// start a typed descendant search below every top-level element.
        /// </summary>
        internal static IEnumerable<T> MetadataDescendants<T>(this Worksheet worksheet) where T : OpenXmlElement {
            for (OpenXmlElement? part = worksheet.FirstChild; part != null; part = part.NextSibling()) {
                if (part is SheetData) continue;
                for (OpenXmlElement? element = part; element != null; element = NextInDocumentOrder(element, part)) {
                    if (element is T typed) yield return typed;
                }
            }
        }

        /// <summary>Reports whether any text below <paramref name="root"/> holds a carriage return, which SDK serialization does not keep.</summary>
        internal static bool HasCarriageReturnText(OpenXmlElement root) {
            for (OpenXmlElement? element = root; element != null; element = NextInDocumentOrder(element, root)) {
                if (element is OpenXmlLeafTextElement text && text.Text?.IndexOf('\r') >= 0) return true;
            }
            return false;
        }

        private static OpenXmlElement? NextInDocumentOrder(OpenXmlElement element, OpenXmlElement top) {
            if (element.FirstChild is { } child) return child;
            for (OpenXmlElement? current = element; current != null && !ReferenceEquals(current, top); current = current.Parent) {
                if (current.NextSibling() is { } sibling) return sibling;
            }
            return null;
        }
    }
}
